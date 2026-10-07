# -*- coding: utf-8 -*-
"""Xuất thông tin thiếu và thời điểm đăng ký do MPLIS lưu cho từng thửa.

Nguyên tắc bắt buộc của công cụ này:

* Danh sách ban đầu luôn đến từ
  ``GetThongKePhanLoaiThuaDatChiTiet`` với ``traCuu[xaId]``.
* Nội dung thiếu chỉ được chép từ trường mô tả do API trả về. Công cụ không
  kiểm tra trường rỗng để tự kết luận "thiếu ...".
* Thời điểm đăng ký chỉ được chép từ
  ``TinhHinhDangKy.thoiDiemDangKy`` của
  ``GetThongTinDangKyByTinhHinhDangKyIds``. Không có fallback sang
  CreatedDate, ModifiedDate, ngày đồng bộ hay thời gian chạy.
* Giá trị ngày giờ được giữ nguyên chuỗi nguồn, không đổi múi giờ và không
  cắt phần giờ/phút/giây.
* Biến động cuối chỉ lấy từ ``Value.LichSuBienDongs[*].ThongTinBienDong``;
  bỏ ``maBienDong == "DH"`` rồi chọn ``thoiDiemBienDong`` mới nhất.

Chạy không tham số để mở giao diện, hoặc:

    python bao_cao_thong_tin_thieu_mplis.py --ma-xa 24163 --output ket_qua.xlsx
"""

from __future__ import annotations

import argparse
import gzip
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager


BASE_URL = "https://dla.mplis.gov.vn"
REFERER_URL = f"{BASE_URL}/dc/DonDangKy/KeKhaiDangKyV2"
URL_DANH_SACH = (
    f"{BASE_URL}/dc/LamSachDuLieuAjax/"
    "GetThongKePhanLoaiThuaDatChiTiet"
)
URL_TTDK = f"{BASE_URL}/dc/DangKyAjax/GetThongTinDangKyByTinhHinhDangKyIds"
URL_LICH_SU_BIEN_DONG = (
    f"{BASE_URL}/dc/ActionChangingAjax/"
    "GetLichSuBienDongByTinhHinhDangKyId"
)

DEFAULT_TINH_ID = 66
DEFAULT_PAGE_SIZE = 10
DEFAULT_DETAIL_BATCH_SIZE = 100
DEFAULT_HISTORY_WORKERS = 5
DEFAULT_TIMEOUT = 120
DEFAULT_DELAY = 0.6


# Đây là tên trường nguồn, không phải quy tắc suy luận nghiệp vụ. Thứ tự ưu
# tiên dành cho trường mô tả hiển thị trực tiếp; hai trường sau vẫn được giữ
# nguyên dưới dạng JSON nếu MPLIS trả một mảng/đối tượng.
MISSING_SOURCE_FIELDS = (
    "chiTietThongTinThieu",
    "errorMessages",
    "thongTinDangKyChuaDapUngNhom1",
)

# Các alias chỉ ánh xạ tên trường vật lý sang cột báo cáo. Không alias nào
# được dùng để suy ra giá trị từ nội dung của trường khác.
DIRECT_FIELD_ALIASES = {
    "ma_xa": ("xaId", "maXa", "ma_xa"),
    "so_to": ("soHieuToBanDo", "soTo", "so_to"),
    "so_thua": ("soThuTuThua", "soThua", "so_thua"),
    "thong_tin_chu": ("thongTinChu",),
    "phan_loai_2959": (
        "phanLoaiKeHoach2959",
        "phanLoai2959",
        "tenPhanLoai",
        "phanLoai",
    ),
    "trang_thai_lam_sach": (
        "trangThaiLamSach",
        "tinhTrangLamSach",
        "tenTrangThaiLamSach",
        "statusLamSach",
        "type",
    ),
    "id_phan_loai": ("id", "_id", "phanLoaiThuaDatChiTietId"),
    "thua_dat_id": ("thuaDatId",),
    "tinh_hinh_dang_ky_id": ("tinhHinhDangKyId",),
}


class MplisReportError(RuntimeError):
    """Lỗi có thể trình bày trực tiếp cho người dùng."""


@dataclass(frozen=True)
class DirectValue:
    value: Any
    source_field: str


@dataclass(frozen=True)
class LatestChange:
    ma_bien_dong: Any
    ma_loai_bien_dong: Any
    loai_bien_dong_id: Any
    thoi_diem_bien_dong: Any
    ly_do_bien_dong: Any
    thong_tin_bien_dong: Any
    source_path: str
    raw: dict[str, Any]


def _not_blank(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, (list, tuple, dict, set)):
        return bool(value)
    return True


def _as_source_text(value: Any) -> str:
    """Biểu diễn giá trị nguồn mà không diễn giải nội dung nghiệp vụ."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return str(value)


def parse_ma_xa_values(value: str | Iterable[Any]) -> list[str]:
    """Chuẩn hóa một hoặc nhiều mã xã, giữ nguyên thứ tự và bỏ trùng."""
    if isinstance(value, str):
        raw_values = re.split(r"[;,\s]+", value.strip())
    else:
        raw_values = [str(item).strip() for item in value]
    result: list[str] = []
    for raw in raw_values:
        code = str(raw).strip()
        if not code:
            continue
        if not re.fullmatch(r"\d+", code):
            raise MplisReportError(f"Mã xã không hợp lệ: {code!r}")
        if code not in result:
            result.append(code)
    if not result:
        raise MplisReportError("Chưa nhập mã xã.")
    return result


def pick_direct(item: dict[str, Any], aliases: Iterable[str]) -> DirectValue | None:
    """Lấy nguyên giá trị của trường đầu tiên thực sự có trong response."""
    for field in aliases:
        if field in item:
            return DirectValue(item[field], field)
    return None


def extract_missing_information(item: dict[str, Any]) -> DirectValue | None:
    """Chỉ lấy mô tả thiếu trực tiếp; không kiểm tra các trường dữ liệu khác."""
    for field in MISSING_SOURCE_FIELDS:
        if field in item and _not_blank(item[field]):
            return DirectValue(item[field], field)
    return None


def unwrap_thong_tin_dang_ky(response_json: dict[str, Any]) -> dict[str, Any] | None:
    """Mở đúng wrapper đã thấy ở API GetThongTinDangKyBy..."""
    value = response_json.get("value")
    if not isinstance(value, list) or not value or not isinstance(value[0], dict):
        return None
    first = value[0]
    wrapped = first.get("thongTinDangKy")
    return wrapped if isinstance(wrapped, dict) else first


def extract_registration_time(response_json: dict[str, Any]) -> DirectValue | None:
    """Lấy duy nhất TinhHinhDangKy.thoiDiemDangKy, không fallback."""
    value = response_json.get("value")
    first = value[0] if isinstance(value, list) and value else None
    ttdk = unwrap_thong_tin_dang_ky(response_json)
    if not isinstance(ttdk, dict):
        return None
    base_path = (
        "value[0].thongTinDangKy"
        if isinstance(first, dict) and isinstance(first.get("thongTinDangKy"), dict)
        else "value[0]"
    )
    tinh_hinh = ttdk.get("TinhHinhDangKy")
    if not isinstance(tinh_hinh, dict) or "thoiDiemDangKy" not in tinh_hinh:
        return None
    return DirectValue(
        tinh_hinh.get("thoiDiemDangKy"),
        f"{base_path}.TinhHinhDangKy.thoiDiemDangKy",
    )


def format_mplis_date_ddmmyyyy(value: Any) -> str:
    """Đổi timestamp MPLIS sang ngày Việt Nam DD/MM/YYYY."""
    if value is None or value == "":
        return ""
    if not isinstance(value, str):
        return ""
    source = value.strip()
    if not source:
        return ""

    dotnet_match = re.fullmatch(r"/Date\((-?\d+)(?:[+-]\d{4})?\)/", source)
    if dotnet_match:
        dt = datetime.fromtimestamp(
            int(dotnet_match.group(1)) / 1000,
            tz=timezone.utc,
        ).astimezone(timezone(timedelta(hours=7)))
        return dt.strftime("%d/%m/%Y")

    try:
        iso_value = source[:-1] + "+00:00" if source.endswith("Z") else source
        dt = datetime.fromisoformat(iso_value)
    except ValueError:
        return ""
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone(timedelta(hours=7)))
    return dt.strftime("%d/%m/%Y")


def format_registration_date_ddmmyyyy(value: Any) -> str:
    """Đổi chính giá trị thoiDiemDangKy sang ngày Việt Nam DD/MM/YYYY."""
    return format_mplis_date_ddmmyyyy(value)


def _mplis_timestamp_sort_value(value: Any) -> float | None:
    """Đọc timestamp để sắp xếp; không dùng trường ngày khác làm fallback."""
    if not isinstance(value, str) or not value.strip():
        return None
    source = value.strip()
    dotnet_match = re.fullmatch(r"/Date\((-?\d+)(?:[+-]\d{4})?\)/", source)
    if dotnet_match:
        return int(dotnet_match.group(1)) / 1000
    try:
        iso_value = source[:-1] + "+00:00" if source.endswith("Z") else source
        dt = datetime.fromisoformat(iso_value)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.timestamp()


def extract_latest_non_dh_change(
    response_json: dict[str, Any],
) -> LatestChange | None:
    """Lấy biến động có thoiDiemBienDong mới nhất và maBienDong khác DH.

    Cấu trúc và tên trường được lấy đúng từ response thực tế trong ``bd.json``:
    ``Value.LichSuBienDongs[*].ThongTinBienDong``. Hàm không fallback sang
    CreatedDate, ModifiedDate, ngayHopDong hoặc một trường ngày nào khác.
    """
    value = response_json.get("Value")
    histories = value.get("LichSuBienDongs") if isinstance(value, dict) else None
    if not isinstance(histories, list):
        return None

    selected: tuple[float, LatestChange] | None = None
    for index, history in enumerate(histories):
        if not isinstance(history, dict):
            continue
        change = history.get("ThongTinBienDong")
        if not isinstance(change, dict):
            continue
        ma_bien_dong = change.get("maBienDong")
        if not isinstance(ma_bien_dong, str) or not ma_bien_dong.strip():
            continue
        if ma_bien_dong.strip().upper() == "DH":
            continue
        timestamp = _mplis_timestamp_sort_value(change.get("thoiDiemBienDong"))
        if timestamp is None:
            continue
        result = LatestChange(
            ma_bien_dong=ma_bien_dong,
            ma_loai_bien_dong=change.get("maLoaiBienDong"),
            loai_bien_dong_id=change.get("loaiBienDongId"),
            thoi_diem_bien_dong=change.get("thoiDiemBienDong"),
            ly_do_bien_dong=change.get("lyDoBienDong"),
            thong_tin_bien_dong=change.get("thongTinBienDong"),
            source_path=(
                f"Value.LichSuBienDongs[{index}].ThongTinBienDong"
            ),
            raw=change,
        )
        if selected is None or timestamp > selected[0]:
            selected = (timestamp, result)
    return selected[1] if selected is not None else None


GCN_DIRECT_FIELDS = (
    "giayChungNhanId",
    "soPhatHanh",
    "soVaoSo",
    "maVach",
    "ngayVaoSo",
)


def extract_giay_chung_nhan(response_json: dict[str, Any]) -> dict[str, str]:
    """Lấy nguyên ListGiayChungNhan và các trường GCN trực tiếp từ MPLIS."""
    result = {field: "" for field in GCN_DIRECT_FIELDS}
    result["raw"] = ""
    result["source"] = ""
    values = response_json.get("value")
    first = values[0] if isinstance(values, list) and values else None
    ttdk = unwrap_thong_tin_dang_ky(response_json)
    if not isinstance(ttdk, dict):
        return result
    certificates = ttdk.get("ListGiayChungNhan")
    if not isinstance(certificates, list):
        return result
    certificates = [item for item in certificates if isinstance(item, dict)]
    if not certificates:
        return result

    base_path = (
        "value[0].thongTinDangKy"
        if isinstance(first, dict) and isinstance(first.get("thongTinDangKy"), dict)
        else "value[0]"
    )
    result["source"] = f"{base_path}.ListGiayChungNhan"
    result["raw"] = json.dumps(
        certificates, ensure_ascii=False, separators=(",", ":")
    )
    for field in GCN_DIRECT_FIELDS:
        values = [certificate.get(field) for certificate in certificates]
        if len(values) == 1:
            result[field] = _as_source_text(values[0])
        else:
            result[field] = json.dumps(
                values, ensure_ascii=False, separators=(",", ":")
            )
    return result


def extract_top_level_ids(item: dict[str, Any]) -> dict[str, Any]:
    """Giữ toàn bộ ID cấp đầu của bản ghi để phục vụ truy vết."""
    ids: dict[str, Any] = {}
    for key, value in item.items():
        key_lower = str(key).lower()
        if key_lower == "id" or key_lower == "_id" or key_lower.endswith("id"):
            if value is not None and not isinstance(value, (dict, list)):
                ids[str(key)] = value
    return ids


def build_list_payload(
    ma_xa: str,
    *,
    start: int,
    length: int,
    tinh_id: int = DEFAULT_TINH_ID,
) -> dict[str, str]:
    """Payload lấy toàn bộ phân loại của xã, không lọc tờ/thửa hay nhóm."""
    return {
        "traCuu[soThuTuThua]": "",
        "traCuu[soHieuToBanDo]": "",
        "traCuu[tinhHinhDangKyId]": "",
        "traCuu[soPhatHanh]": "",
        "traCuu[hoTenChu]": "",
        "traCuu[phanLoai]": "7",
        "traCuu[type]": "-1",
        "traCuu[loaiTaiSan]": "-1",
        "traCuu[loaiChu]": "-1",
        "traCuu[tuNgay]": "",
        "traCuu[denNgay]": "",
        "traCuu[query]": "",
        "traCuu[xaId]": str(ma_xa),
        "traCuu[huyenId]": "0",
        "traCuu[tinhId]": str(tinh_id),
        "start": str(start),
        "length": str(length),
        "exportWard": "false",
        "subLength": "3000",
        "sort[Field]": "_id",
        "sort[Direction]": "1",
    }


def build_change_history_payload(
    tinh_hinh_dang_ky_id: Any,
    *,
    so_doi: int = 0,
) -> dict[str, int]:
    try:
        tid = int(tinh_hinh_dang_ky_id)
    except (TypeError, ValueError) as exc:
        raise MplisReportError(
            f"tinhHinhDangKyId không hợp lệ: {tinh_hinh_dang_ky_id!r}"
        ) from exc
    return {"tinhHinhDangKyId": tid, "soDoi": int(so_doi)}


def _json_response(response: requests.Response, label: str) -> dict[str, Any]:
    auth_marker = response.headers.get("X-Responded-JSON", "")
    if '"status":401' in auth_marker.replace(" ", ""):
        raise MplisReportError(f"Phiên MPLIS hết hạn khi gọi {label}.")
    try:
        payload = response.json()
    except ValueError as exc:
        raise MplisReportError(
            f"{label} trả HTTP {response.status_code} nhưng không phải JSON: "
            f"{response.text[:300]}"
        ) from exc
    if not response.ok:
        raise MplisReportError(
            f"{label} trả HTTP {response.status_code}: {_as_source_text(payload)[:500]}"
        )
    if payload.get("success") is False:
        raise MplisReportError(f"{label} success=false: {_as_source_text(payload)[:500]}")
    return payload


def fetch_list_page(
    session: requests.Session,
    ma_xa: str,
    *,
    start: int,
    length: int,
    tinh_id: int,
    timeout: int = DEFAULT_TIMEOUT,
) -> dict[str, Any]:
    headers = dict(session.headers)
    headers.update({
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Referer": REFERER_URL,
    })
    response = session.post(
        URL_DANH_SACH,
        data=build_list_payload(
            ma_xa, start=start, length=length, tinh_id=tinh_id
        ),
        headers=headers,
        timeout=timeout,
    )
    return _json_response(response, "GetThongKePhanLoaiThuaDatChiTiet")


def fetch_registration_batch(
    session: requests.Session,
    tinh_hinh_dang_ky_ids: Iterable[Any],
    *,
    timeout: int = DEFAULT_TIMEOUT,
) -> dict[str, Any]:
    ids: list[int] = []
    for raw_id in tinh_hinh_dang_ky_ids:
        try:
            tid = int(raw_id)
        except (TypeError, ValueError) as exc:
            raise MplisReportError(
                f"tinhHinhDangKyId không hợp lệ: {raw_id!r}"
            ) from exc
        if tid not in ids:
            ids.append(tid)
    if not ids:
        raise MplisReportError("Danh sách tinhHinhDangKyId đang rỗng.")

    headers = dict(session.headers)
    headers.update({
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Referer": REFERER_URL,
    })
    response = session.post(
        URL_TTDK,
        data=json.dumps(
            {"tinhHinhDangKyIds": ids, "getHoSoQuet": False},
            ensure_ascii=False,
        ),
        headers=headers,
        timeout=timeout,
    )
    payload = _json_response(response, "GetThongTinDangKyByTinhHinhDangKyIds")
    if not payload.get("value"):
        raise MplisReportError(
            "API chi tiết không có value cho lô ID: "
            + ", ".join(str(value) for value in ids[:10])
        )
    return payload


def fetch_registration_response(
    session: requests.Session,
    tinh_hinh_dang_ky_id: Any,
    *,
    timeout: int = DEFAULT_TIMEOUT,
) -> dict[str, Any]:
    """Tương thích lời gọi cũ cho một ID; luồng báo cáo dùng API theo lô."""
    return fetch_registration_batch(
        session, [tinh_hinh_dang_ky_id], timeout=timeout
    )


def fetch_change_history(
    session: requests.Session,
    tinh_hinh_dang_ky_id: Any,
    *,
    so_doi: int = 0,
    timeout: int = DEFAULT_TIMEOUT,
) -> dict[str, Any]:
    """Gọi đúng endpoint/payload lịch sử biến động quan sát từ Network MPLIS."""
    headers = dict(session.headers)
    headers.update({
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Referer": REFERER_URL,
    })
    response = session.post(
        URL_LICH_SU_BIEN_DONG,
        data=json.dumps(
            build_change_history_payload(
                tinh_hinh_dang_ky_id, so_doi=so_doi
            ),
            ensure_ascii=False,
        ),
        headers=headers,
        timeout=timeout,
    )
    return _json_response(response, "GetLichSuBienDongByTinhHinhDangKyId")


def _copy_read_session(session: requests.Session) -> requests.Session:
    """Tạo session chỉ đọc độc lập để các request lịch sử chạy song song an toàn."""
    copied = requests.Session()
    copied.headers.update(session.headers)
    copied.cookies.update(session.cookies)
    return copied


def registration_responses_by_id(
    response_json: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    """Tách response nhiều ID thành response một phần tử cho từng ID."""
    mapped: dict[str, dict[str, Any]] = {}
    values = response_json.get("value")
    if not isinstance(values, list):
        return mapped
    for value in values:
        if not isinstance(value, dict):
            continue
        ttdk = value.get("thongTinDangKy")
        if not isinstance(ttdk, dict):
            ttdk = value
        tinh_hinh = ttdk.get("TinhHinhDangKy")
        if not isinstance(tinh_hinh, dict):
            continue
        tid = tinh_hinh.get("tinhHinhDangKyId")
        if tid is None:
            continue
        mapped[str(tid)] = {
            "success": response_json.get("success", True),
            "value": [value],
        }
    return mapped


def fetch_all_list_items(
    session: requests.Session,
    ma_xa: str,
    *,
    tinh_id: int = DEFAULT_TINH_ID,
    page_size: int = DEFAULT_PAGE_SIZE,
    delay: float = DEFAULT_DELAY,
    raw_dir: Path | None = None,
    log: Callable[[str], None] = print,
) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    start = 0
    page_no = 0
    previous_signature: str | None = None

    while True:
        page_no += 1
        payload = fetch_list_page(
            session,
            ma_xa,
            start=start,
            length=page_size,
            tinh_id=tinh_id,
        )
        if raw_dir is not None:
            raw_dir.mkdir(parents=True, exist_ok=True)
            (raw_dir / f"danh_sach_trang_{page_no:04d}.json").write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

        page_items = payload.get("data") or []
        if not isinstance(page_items, list):
            raise MplisReportError("Response danh sách có trường data không phải mảng.")
        if not page_items:
            break
        if not all(isinstance(row, dict) for row in page_items):
            raise MplisReportError("Response danh sách chứa phần tử không phải object.")

        signature = json.dumps(page_items[:2], ensure_ascii=False, sort_keys=True)
        if signature == previous_signature:
            raise MplisReportError(
                "MPLIS lặp lại cùng một trang; đã dừng để tránh nhân đôi dữ liệu."
            )
        previous_signature = signature
        items.extend(page_items)
        log(f"Đã lấy trang {page_no}: {len(page_items)} bản ghi, tổng {len(items)}.")

        total = payload.get("recordsFiltered")
        if total is None:
            total = payload.get("recordsTotal")
        try:
            total_int = int(total) if total is not None else None
        except (TypeError, ValueError):
            total_int = None

        start += len(page_items)
        if total_int is not None and start >= total_int:
            break
        if len(page_items) < page_size:
            break
        if delay > 0:
            time.sleep(delay)

    return items


def build_report_row(
    item: dict[str, Any],
    *,
    ma_xa_requested: str,
    detail_response: dict[str, Any] | None,
    detail_error: str = "",
    change_history_response: dict[str, Any] | None = None,
    change_history_error: str = "",
    change_history_enabled: bool | None = None,
) -> dict[str, Any]:
    direct = {
        name: pick_direct(item, aliases)
        for name, aliases in DIRECT_FIELD_ALIASES.items()
    }
    missing = extract_missing_information(item)
    registration = (
        extract_registration_time(detail_response)
        if isinstance(detail_response, dict)
        else None
    )
    gcn = (
        extract_giay_chung_nhan(detail_response)
        if isinstance(detail_response, dict)
        else extract_giay_chung_nhan({})
    )
    latest_change = (
        extract_latest_non_dh_change(change_history_response)
        if isinstance(change_history_response, dict)
        else None
    )
    if change_history_enabled is None:
        change_history_enabled = (
            isinstance(change_history_response, dict)
            or bool(change_history_error)
        )

    notes: list[str] = []
    if missing is None:
        notes.append("Response danh sách không có trường mô tả thiếu trực tiếp đã biết")
    if detail_error:
        notes.append(detail_error)
    elif registration is None:
        notes.append("API chi tiết không có đúng trường TinhHinhDangKy.thoiDiemDangKy")
    elif registration.value is None or registration.value == "":
        notes.append("MPLIS trả thoiDiemDangKy rỗng/null")
    if change_history_error:
        notes.append(change_history_error)
    elif isinstance(change_history_response, dict) and latest_change is None:
        notes.append(
            "API lịch sử không có biến động với maBienDong khác DH "
            "và thoiDiemBienDong hợp lệ"
        )

    def value_of(name: str) -> str:
        selected = direct.get(name)
        return _as_source_text(selected.value) if selected is not None else ""

    def source_of(name: str) -> str:
        selected = direct.get(name)
        return selected.source_field if selected is not None else ""

    return {
        "ma_xa": value_of("ma_xa") or str(ma_xa_requested),
        "ma_xa_source": source_of("ma_xa") or "tham số traCuu[xaId]",
        "so_to": value_of("so_to"),
        "so_to_source": source_of("so_to"),
        "so_thua": value_of("so_thua"),
        "so_thua_source": source_of("so_thua"),
        "thong_tin_chu": value_of("thong_tin_chu"),
        "thong_tin_chu_source": source_of("thong_tin_chu"),
        "phan_loai_2959": value_of("phan_loai_2959"),
        "phan_loai_source": source_of("phan_loai_2959"),
        "trang_thai_lam_sach": value_of("trang_thai_lam_sach"),
        "trang_thai_source": source_of("trang_thai_lam_sach"),
        "id_phan_loai": value_of("id_phan_loai"),
        "thua_dat_id": value_of("thua_dat_id"),
        "tinh_hinh_dang_ky_id": value_of("tinh_hinh_dang_ky_id"),
        "cac_id_nguon": json.dumps(
            extract_top_level_ids(item), ensure_ascii=False, separators=(",", ":")
        ),
        "gcn_id": gcn["giayChungNhanId"],
        "gcn_so_phat_hanh": gcn["soPhatHanh"],
        "gcn_so_vao_so": gcn["soVaoSo"],
        "gcn_ma_vach": gcn["maVach"],
        "gcn_ngay_vao_so": gcn["ngayVaoSo"],
        "gcn_source": gcn["source"],
        "gcn_raw": gcn["raw"],
        "thong_tin_thieu": _as_source_text(missing.value) if missing else "",
        "thong_tin_thieu_source": missing.source_field if missing else "",
        "thoi_diem_dang_ky": (
            _as_source_text(registration.value) if registration is not None else ""
        ),
        "ngay_dang_ky_ddmmyyyy": (
            format_registration_date_ddmmyyyy(registration.value)
            if registration is not None
            else ""
        ),
        "thoi_diem_dang_ky_source": (
            registration.source_field if registration is not None else ""
        ),
        "ma_bien_dong_cuoi": (
            _as_source_text(latest_change.ma_bien_dong) if latest_change else ""
        ),
        "ma_loai_bien_dong_cuoi": (
            _as_source_text(latest_change.ma_loai_bien_dong)
            if latest_change else ""
        ),
        "loai_bien_dong_id_cuoi": (
            _as_source_text(latest_change.loai_bien_dong_id)
            if latest_change else ""
        ),
        "thoi_diem_bien_dong_cuoi": (
            _as_source_text(latest_change.thoi_diem_bien_dong)
            if latest_change else ""
        ),
        "ngay_bien_dong_cuoi_ddmmyyyy": (
            format_mplis_date_ddmmyyyy(latest_change.thoi_diem_bien_dong)
            if latest_change else ""
        ),
        "ly_do_bien_dong_cuoi": (
            _as_source_text(latest_change.ly_do_bien_dong)
            if latest_change else ""
        ),
        "thong_tin_bien_dong_cuoi": (
            _as_source_text(latest_change.thong_tin_bien_dong)
            if latest_change else ""
        ),
        "bien_dong_cuoi_source": (
            latest_change.source_path if latest_change else ""
        ),
        "bien_dong_cuoi_raw": (
            json.dumps(
                latest_change.raw, ensure_ascii=False, separators=(",", ":")
            )
            if latest_change else ""
        ),
        "trang_thai_lich_su_bien_dong": (
            "Đã tắt"
            if not change_history_enabled
            else "Lỗi"
            if change_history_error
            else "Đã truy vấn"
            if isinstance(change_history_response, dict)
            else "Chưa truy vấn"
        ),
        "trang_thai_truy_van": (
            "Lỗi" if detail_error or change_history_error else "Đã truy vấn"
        ),
        "ghi_chu": "; ".join(notes),
        "raw_item": json.dumps(item, ensure_ascii=False, separators=(",", ":")),
    }


REPORT_HEADERS = (
    ("Mã xã", "ma_xa"),
    ("Số tờ", "so_to"),
    ("Số thửa", "so_thua"),
    ("Thông tin chủ", "thong_tin_chu"),
    ("Phân loại Kế hoạch 2959", "phan_loai_2959"),
    ("Trạng thái làm sạch", "trang_thai_lam_sach"),
    ("ID bản ghi phân loại", "id_phan_loai"),
    ("thuaDatId", "thua_dat_id"),
    ("tinhHinhDangKyId", "tinh_hinh_dang_ky_id"),
    ("Các ID cấp đầu từ MPLIS", "cac_id_nguon"),
    ("giayChungNhanId", "gcn_id"),
    ("Số phát hành GCN", "gcn_so_phat_hanh"),
    ("Số vào sổ GCN", "gcn_so_vao_so"),
    ("Mã vạch GCN", "gcn_ma_vach"),
    ("Ngày vào sổ GCN", "gcn_ngay_vao_so"),
    ("Trường nguồn GCN", "gcn_source"),
    ("Danh sách GCN gốc (JSON)", "gcn_raw"),
    ("Thông tin thiếu MPLIS", "thong_tin_thieu"),
    ("Trường nguồn thông tin thiếu", "thong_tin_thieu_source"),
    ("Thời điểm đăng ký", "thoi_diem_dang_ky"),
    ("Ngày đăng ký (DD/MM/YYYY)", "ngay_dang_ky_ddmmyyyy"),
    ("Trường nguồn thời điểm đăng ký", "thoi_diem_dang_ky_source"),
    ("Mã biến động cuối (khác DH)", "ma_bien_dong_cuoi"),
    ("Mã loại biến động cuối", "ma_loai_bien_dong_cuoi"),
    ("loaiBienDongId cuối", "loai_bien_dong_id_cuoi"),
    ("Thời điểm biến động cuối", "thoi_diem_bien_dong_cuoi"),
    ("Ngày biến động cuối (DD/MM/YYYY)", "ngay_bien_dong_cuoi_ddmmyyyy"),
    ("Loại/lý do biến động cuối từ MPLIS", "ly_do_bien_dong_cuoi"),
    ("Nội dung biến động cuối từ MPLIS", "thong_tin_bien_dong_cuoi"),
    ("Trường nguồn biến động cuối", "bien_dong_cuoi_source"),
    ("Biến động cuối gốc (JSON)", "bien_dong_cuoi_raw"),
    ("Trạng thái lịch sử biến động", "trang_thai_lich_su_bien_dong"),
    ("Trạng thái truy vấn chi tiết", "trang_thai_truy_van"),
    ("Ghi chú", "ghi_chu"),
)

SOURCE_HEADERS = (
    ("Mã xã", "ma_xa_source"),
    ("Số tờ", "so_to_source"),
    ("Số thửa", "so_thua_source"),
    ("Thông tin chủ", "thong_tin_chu_source"),
    ("Phân loại Kế hoạch 2959", "phan_loai_source"),
    ("Trạng thái làm sạch", "trang_thai_source"),
)


def _style_sheet(ws: Any, widths: list[int]) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 34
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.number_format = "@"
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def write_report(rows: list[dict[str, Any]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "KetQua"
    ws.append(["STT", *(label for label, _ in REPORT_HEADERS)])
    for index, row in enumerate(rows, start=1):
        ws.append([index, *(row.get(key, "") for _, key in REPORT_HEADERS)])
    _style_sheet(
        ws,
        [
            8, 12, 12, 14, 32, 24, 22, 20, 16, 20, 42,
            20, 24, 24, 22, 28, 48, 80,
            48, 34, 30, 24, 58,
            24, 18, 18, 30, 24, 36, 80, 58, 80, 26,
            22, 55,
        ],
    )

    source_ws = wb.create_sheet("TruongNguon")
    source_ws.append([
        "STT",
        *(label for label, _ in SOURCE_HEADERS),
        "Raw item từ endpoint danh sách",
    ])
    for index, row in enumerate(rows, start=1):
        source_ws.append([
            index,
            *(row.get(key, "") for _, key in SOURCE_HEADERS),
            row.get("raw_item", ""),
        ])
    _style_sheet(source_ws, [8, 24, 24, 24, 30, 30, 30, 100])
    wb.save(output_path)


def token_from_page(driver: webdriver.Chrome) -> str:
    return driver.execute_script(
        """
        return (
            document.querySelector('input[name="__RequestVerificationToken"]')?.value ||
            document.querySelector('input[name="__requestverificationtoken"]')?.value ||
            document.querySelector('meta[name="__RequestVerificationToken"]')?.content ||
            document.querySelector('meta[name="__requestverificationtoken"]')?.content ||
            document.querySelector('meta[name="RequestVerificationToken"]')?.content || ''
        );
        """
    )


def session_from_driver(driver: webdriver.Chrome) -> requests.Session:
    token = token_from_page(driver)
    if not token:
        raise MplisReportError(
            "Không lấy được RequestVerificationToken. Hãy mở đúng màn hình MPLIS sau đăng nhập."
        )
    session = requests.Session()
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Referer": REFERER_URL,
        "__requestverificationtoken": token,
        "__RequestVerificationToken": token,
        "RequestVerificationToken": token,
    })
    for cookie in driver.get_cookies():
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain"),
            path=cookie.get("path", "/"),
        )
    return session


def _strip_copied_header(value: str, header_names: Iterable[str]) -> str:
    """Cho phép dán cả dòng ``Header-Name: value`` từ DevTools."""
    cleaned = str(value or "").strip()
    for name in header_names:
        match = re.match(
            rf"^\s*{re.escape(name)}\s*:\s*(.*)$",
            cleaned,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if match:
            return match.group(1).strip()
    return cleaned


def session_from_token_cookie(token: str, cookie_raw: str) -> requests.Session:
    """Tạo session từ Token/Cookie được sao chép ở Request Headers."""
    token = _strip_copied_header(token, (
        "__requestverificationtoken",
        "__RequestVerificationToken",
        "RequestVerificationToken",
    ))
    cookie_raw = _strip_copied_header(cookie_raw, ("Cookie",))
    if not token:
        raise MplisReportError("Chưa nhập Token MPLIS.")
    if not cookie_raw:
        raise MplisReportError("Chưa nhập Cookie MPLIS.")

    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/129.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "vi-VN,vi;q=0.9",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE_URL,
        "Referer": REFERER_URL,
        "__requestverificationtoken": token,
        "__RequestVerificationToken": token,
        "RequestVerificationToken": token,
    })

    cookie_count = 0
    for part in cookie_raw.replace("\r", "").replace("\n", "").split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, value = part.split("=", 1)
        name = name.strip()
        if not name:
            continue
        session.cookies.set(name, value.strip())
        cookie_count += 1
    if cookie_count == 0:
        raise MplisReportError(
            "Không đọc được cookie nào. Hãy dán nguyên giá trị Request Headers > Cookie."
        )
    return session


def login_mplis(username: str = "", password: str = "") -> webdriver.Chrome:
    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(
        options=options,
        service=Service(ChromeDriverManager().install()),
    )
    driver.get(REFERER_URL)
    if username and password:
        try:
            time.sleep(2)
            inputs = driver.find_elements(By.CSS_SELECTOR, "input")
            user_box = next(
                (
                    element
                    for element in inputs
                    if (element.get_attribute("type") or "").lower() in ("text", "email")
                ),
                None,
            )
            pass_box = next(
                (
                    element
                    for element in inputs
                    if (element.get_attribute("type") or "").lower() == "password"
                ),
                None,
            )
            if user_box is not None and pass_box is not None:
                user_box.clear()
                user_box.send_keys(username)
                pass_box.clear()
                pass_box.send_keys(password)
                pass_box.send_keys(Keys.ENTER)
        except Exception:
            pass
    return driver


def run_report(
    *,
    ma_xa: str | Iterable[Any],
    output_path: Path,
    username: str = "",
    password: str = "",
    token: str = "",
    cookie: str = "",
    tinh_id: int = DEFAULT_TINH_ID,
    page_size: int = DEFAULT_PAGE_SIZE,
    delay: float = DEFAULT_DELAY,
    include_change_history: bool = True,
    log: Callable[[str], None] = print,
    wait_for_login: Callable[[], None] | None = None,
) -> Path:
    driver: webdriver.Chrome | None = None
    ma_xas = parse_ma_xa_values(ma_xa)
    raw_dir = output_path.with_name(f"{output_path.stem}_raw")
    try:
        has_token = bool(str(token or "").strip())
        has_cookie = bool(str(cookie or "").strip())
        if has_token != has_cookie:
            raise MplisReportError(
                "Phải nhập đồng thời cả Token và Cookie, hoặc để trống cả hai để đăng nhập Chrome."
            )

        if has_token and has_cookie:
            session = session_from_token_cookie(token, cookie)
            log("Đã tạo session từ Token/Cookie; không mở Chrome.")
        else:
            driver = login_mplis(username, password)
            if wait_for_login is not None:
                wait_for_login()
            else:
                input(
                    "Hoàn tất đăng nhập/xác thực trên Chrome, mở màn hình MPLIS rồi nhấn Enter..."
                )
            session = session_from_driver(driver)

        scoped_items: list[tuple[str, dict[str, Any]]] = []
        for xa_index, current_ma_xa in enumerate(ma_xas, start=1):
            log(
                "Đang lấy danh sách từ GetThongKePhanLoaiThuaDatChiTiet "
                f"cho xã {current_ma_xa} ({xa_index}/{len(ma_xas)})..."
            )
            commune_items = fetch_all_list_items(
                session,
                current_ma_xa,
                tinh_id=tinh_id,
                page_size=page_size,
                delay=delay,
                raw_dir=raw_dir / f"xa_{current_ma_xa}",
                log=log,
            )
            scoped_items.extend(
                (current_ma_xa, item) for item in commune_items
            )
            if delay > 0 and xa_index < len(ma_xas):
                time.sleep(delay)
        if not scoped_items:
            raise MplisReportError(
                "MPLIS không trả bản ghi nào cho các mã xã: " + ", ".join(ma_xas)
            )

        unique_tids: list[int] = []
        for _, item in scoped_items:
            selected = pick_direct(
                item, DIRECT_FIELD_ALIASES["tinh_hinh_dang_ky_id"]
            )
            if selected is None or not _not_blank(selected.value):
                continue
            try:
                tid = int(selected.value)
            except (TypeError, ValueError):
                continue
            if tid not in unique_tids:
                unique_tids.append(tid)

        detail_by_tid: dict[str, dict[str, Any]] = {}
        detail_error_by_tid: dict[str, str] = {}
        batches = [
            unique_tids[start:start + DEFAULT_DETAIL_BATCH_SIZE]
            for start in range(0, len(unique_tids), DEFAULT_DETAIL_BATCH_SIZE)
        ]
        for batch_no, batch_ids in enumerate(batches, start=1):
            try:
                batch_response = fetch_registration_batch(session, batch_ids)
                (raw_dir / f"chi_tiet_ttdk_lo_{batch_no:04d}.json").write_text(
                    json.dumps(batch_response, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                mapped = registration_responses_by_id(batch_response)
                detail_by_tid.update(mapped)
                for tid in batch_ids:
                    if str(tid) not in mapped:
                        detail_error_by_tid[str(tid)] = (
                            "API chi tiết không trả bản ghi tương ứng trong response theo lô"
                        )
            except Exception as exc:
                for tid in batch_ids:
                    detail_error_by_tid[str(tid)] = str(exc)
            log(
                f"Đã đọc chi tiết lô {batch_no}/{len(batches)} "
                f"({len(batch_ids)} ID)."
            )
            if delay > 0 and batch_no < len(batches):
                time.sleep(delay)

        change_history_by_tid: dict[str, dict[str, Any]] = {}
        change_history_error_by_tid: dict[str, str] = {}

        def load_change_history(tid: int) -> dict[str, Any]:
            read_session = _copy_read_session(session)
            try:
                return fetch_change_history(read_session, tid, so_doi=0)
            finally:
                read_session.close()

        if include_change_history and unique_tids:
            history_dir = raw_dir / "lich_su_bien_dong"
            history_dir.mkdir(parents=True, exist_ok=True)
            worker_count = min(DEFAULT_HISTORY_WORKERS, len(unique_tids))
            log(
                "Đang đọc lịch sử biến động theo từng tinhHinhDangKyId "
                f"với {worker_count} luồng..."
            )
            with ThreadPoolExecutor(max_workers=worker_count) as executor:
                futures = {
                    executor.submit(load_change_history, tid): tid
                    for tid in unique_tids
                }
                for completed, future in enumerate(
                    as_completed(futures), start=1
                ):
                    tid = futures[future]
                    tid_key = str(tid)
                    try:
                        history_response = future.result()
                        change_history_by_tid[tid_key] = history_response
                        with gzip.open(
                            history_dir / f"tinh_hinh_dang_ky_{tid}.json.gz",
                            "wt",
                            encoding="utf-8",
                        ) as raw_file:
                            json.dump(
                                history_response,
                                raw_file,
                                ensure_ascii=False,
                                indent=2,
                            )
                    except Exception as exc:
                        change_history_error_by_tid[tid_key] = (
                            f"Lỗi lịch sử biến động: {exc}"
                        )
                    if completed % 10 == 0 or completed == len(unique_tids):
                        log(
                            "Đã đọc lịch sử biến động "
                            f"{completed}/{len(unique_tids)} ID."
                        )
        elif not include_change_history:
            log("Đã tắt lấy lịch sử biến động; bỏ qua endpoint lịch sử.")

        rows: list[dict[str, Any]] = []
        for current_ma_xa, item in scoped_items:
            tid_value = pick_direct(
                item, DIRECT_FIELD_ALIASES["tinh_hinh_dang_ky_id"]
            )
            detail: dict[str, Any] | None = None
            detail_error = ""
            change_history: dict[str, Any] | None = None
            change_history_error = ""
            if tid_value is None or not _not_blank(tid_value.value):
                detail_error = (
                    "Bản ghi danh sách không có tinhHinhDangKyId; "
                    "không thể truy vấn thời điểm đăng ký"
                )
                change_history_error = (
                    "Bản ghi danh sách không có tinhHinhDangKyId; "
                    "không thể truy vấn lịch sử biến động"
                )
            else:
                try:
                    tid_key = str(int(tid_value.value))
                    detail = detail_by_tid.get(tid_key)
                    detail_error = detail_error_by_tid.get(tid_key, "")
                    change_history = change_history_by_tid.get(tid_key)
                    change_history_error = change_history_error_by_tid.get(
                        tid_key, ""
                    )
                except (TypeError, ValueError):
                    detail_error = (
                        f"tinhHinhDangKyId không hợp lệ: {tid_value.value!r}"
                    )
                    change_history_error = detail_error

            rows.append(build_report_row(
                item,
                ma_xa_requested=current_ma_xa,
                detail_response=detail,
                detail_error=detail_error,
                change_history_response=change_history,
                change_history_error=change_history_error,
                change_history_enabled=include_change_history,
            ))

        write_report(rows, output_path)
        log(f"Đã ghi báo cáo: {output_path}")
        log(f"Đã lưu response gốc để đối chiếu: {raw_dir}")
        return output_path
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


def _default_output(ma_xa: str | Iterable[Any]) -> Path:
    ma_xas = parse_ma_xa_values(ma_xa)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    xa_label = "_".join(ma_xas[:5])
    if len(ma_xas) > 5:
        xa_label += f"_va_{len(ma_xas) - 5}_xa_khac"
    return Path.cwd() / f"thong_tin_thieu_mplis_{xa_label}_{stamp}.xlsx"


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class ReportApp(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title("Báo cáo thông tin thiếu và thời điểm đăng ký MPLIS")
            self.geometry("920x640")
            self.log_queue: queue.Queue[Any] = queue.Queue()
            self.worker: threading.Thread | None = None

            self.username = tk.StringVar()
            self.password = tk.StringVar()
            self.token = tk.StringVar()
            self.ma_xa = tk.StringVar()
            self.output = tk.StringVar()
            self.tinh_id = tk.IntVar(value=DEFAULT_TINH_ID)
            self.page_size = tk.IntVar(value=DEFAULT_PAGE_SIZE)
            self.delay = tk.DoubleVar(value=DEFAULT_DELAY)
            self.include_change_history = tk.BooleanVar(value=True)
            self._build()
            self.after(150, self._drain_log)

        def _build(self) -> None:
            form = ttk.LabelFrame(self, text="Thông tin tra cứu", padding=10)
            form.pack(fill="x", padx=10, pady=10)
            labels = (
                ("Username", self.username, False),
                ("Password", self.password, True),
                ("Token MPLIS", self.token, False),
            )
            for row, (label, variable, secret) in enumerate(labels):
                ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", pady=4)
                ttk.Entry(
                    form,
                    textvariable=variable,
                    width=45,
                    show="*" if secret else "",
                ).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
            ttk.Label(form, text="Cookie MPLIS").grid(row=3, column=0, sticky="nw", pady=4)
            self.cookie_box = tk.Text(form, width=70, height=3, wrap="word")
            self.cookie_box.grid(row=3, column=1, columnspan=2, sticky="ew", padx=8, pady=4)

            compact = ttk.Frame(form)
            compact.grid(row=4, column=0, columnspan=3, sticky="ew", pady=4)
            compact_fields = (
                ("Mã xã (;)", self.ma_xa, 18),
                ("Mã tỉnh", self.tinh_id, 7),
                ("Số bản ghi/trang", self.page_size, 9),
                ("Giãn request (giây)", self.delay, 7),
            )
            for index, (label, variable, width) in enumerate(compact_fields):
                column = index * 2
                ttk.Label(compact, text=label).grid(
                    row=0, column=column, sticky="w", padx=(0 if index == 0 else 12, 4)
                )
                ttk.Entry(compact, textvariable=variable, width=width).grid(
                    row=0, column=column + 1, sticky="w"
                )

            ttk.Checkbutton(
                form,
                text="Lấy lịch sử biến động (bỏ mã DH)",
                variable=self.include_change_history,
            ).grid(row=5, column=1, sticky="w", padx=8, pady=4)

            ttk.Label(form, text="File kết quả").grid(row=6, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=self.output, width=70).grid(
                row=6, column=1, sticky="ew", padx=8, pady=4
            )
            ttk.Button(form, text="Chọn...", command=self._choose_output).grid(
                row=6, column=2, pady=4
            )
            form.columnconfigure(1, weight=1)

            ttk.Label(
                self,
                text=(
                    "Có thể dán Token + Cookie từ F12 > Network > Request Headers để bỏ qua Chrome. "
                    "Công cụ chỉ đọc. Nội dung thiếu và thời điểm đăng ký được chép "
                    "nguyên từ trường nguồn MPLIS; biến động cuối bỏ mã DH và lấy "
                    "theo thoiDiemBienDong. Không tự suy luận."
                ),
                foreground="#1F4E78",
                wraplength=860,
            ).pack(fill="x", padx=14, pady=(0, 8))
            self.start_button = ttk.Button(self, text="Lấy báo cáo", command=self._start)
            self.start_button.pack(anchor="w", padx=12, pady=4)
            self.log_box = tk.Text(self, wrap="word", font=("Consolas", 9))
            self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

        def _choose_output(self) -> None:
            path = filedialog.asksaveasfilename(
                title="Chọn nơi lưu báo cáo",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
            )
            if path:
                self.output.set(path)

        def _log(self, message: str) -> None:
            self.log_queue.put(message)

        def _wait_for_login(self) -> None:
            event = threading.Event()

            def ask() -> None:
                messagebox.showinfo(
                    "Đăng nhập MPLIS",
                    "Hoàn tất đăng nhập và Authenticator trên Chrome.\n"
                    "Sau khi vào được MPLIS, bấm OK để công cụ đọc dữ liệu.",
                )
                event.set()

            self.after(0, ask)
            event.wait()

        def _start(self) -> None:
            if self.worker and self.worker.is_alive():
                return
            ma_xa = self.ma_xa.get().strip()
            try:
                ma_xas = parse_ma_xa_values(ma_xa)
            except MplisReportError as exc:
                messagebox.showerror("Mã xã không hợp lệ", str(exc))
                return
            output = (
                Path(self.output.get().strip())
                if self.output.get().strip()
                else _default_output(ma_xas)
            )
            self.output.set(str(output))
            token = self.token.get().strip()
            cookie = self.cookie_box.get("1.0", "end").strip()
            if bool(token) != bool(cookie):
                messagebox.showerror(
                    "Thiếu Token/Cookie",
                    "Phải dán đồng thời cả Token và Cookie, hoặc để trống cả hai để đăng nhập Chrome.",
                )
                return
            username = self.username.get().strip()
            password = self.password.get()
            tinh_id = int(self.tinh_id.get())
            page_size = max(1, int(self.page_size.get()))
            delay = max(0.0, float(self.delay.get()))
            include_change_history = bool(self.include_change_history.get())
            self.start_button.config(state="disabled")

            def work() -> None:
                try:
                    run_report(
                        ma_xa=ma_xas,
                        output_path=output,
                        username=username,
                        password=password,
                        token=token,
                        cookie=cookie,
                        tinh_id=tinh_id,
                        page_size=page_size,
                        delay=delay,
                        include_change_history=include_change_history,
                        log=self._log,
                        wait_for_login=self._wait_for_login,
                    )
                    self.log_queue.put(("done", str(output)))
                except Exception as exc:
                    traceback.print_exc()
                    self.log_queue.put(("error", str(exc)))

            self.worker = threading.Thread(target=work, daemon=True)
            self.worker.start()

        def _drain_log(self) -> None:
            try:
                while True:
                    message = self.log_queue.get_nowait()
                    if isinstance(message, tuple):
                        self.start_button.config(state="normal")
                        kind, text = message
                        if kind == "done":
                            messagebox.showinfo("Hoàn tất", f"Đã lưu báo cáo:\n{text}")
                        else:
                            messagebox.showerror("Lỗi", text)
                    else:
                        self.log_box.insert("end", str(message) + "\n")
                        self.log_box.see("end")
            except queue.Empty:
                pass
            self.after(150, self._drain_log)

    ReportApp().mainloop()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ma-xa")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--username", default=os.environ.get("MPLIS_USERNAME", ""))
    parser.add_argument("--password", default=os.environ.get("MPLIS_PASSWORD", ""))
    parser.add_argument("--token", default=os.environ.get("MPLIS_TOKEN", ""))
    parser.add_argument("--cookie", default=os.environ.get("MPLIS_COOKIE", ""))
    parser.add_argument("--tinh-id", type=int, default=DEFAULT_TINH_ID)
    parser.add_argument("--page-size", type=int, default=DEFAULT_PAGE_SIZE)
    parser.add_argument("--delay", type=float, default=DEFAULT_DELAY)
    parser.add_argument(
        "--khong-lay-lich-su-bien-dong",
        action="store_true",
        help="Bỏ qua endpoint lịch sử biến động để chạy nhanh hơn.",
    )
    return parser.parse_args()


def main() -> int:
    # Windows có thể để stdout ở cp1252 dù mã nguồn/terminal dùng Unicode.
    # Chỉ đổi encoding hiển thị; không ảnh hưởng dữ liệu gửi tới MPLIS.
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")
    args = parse_args()
    if not args.ma_xa:
        launch_gui()
        return 0
    ma_xas = parse_ma_xa_values(args.ma_xa)
    output = args.output or _default_output(ma_xas)
    run_report(
        ma_xa=ma_xas,
        output_path=output,
        username=args.username,
        password=args.password,
        token=args.token,
        cookie=args.cookie,
        tinh_id=args.tinh_id,
        page_size=max(1, args.page_size),
        delay=max(0.0, args.delay),
        include_change_history=not args.khong_lay_lich_su_bien_dong,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
