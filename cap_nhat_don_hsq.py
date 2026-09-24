# -*- coding: utf-8 -*-
"""
Tool cập nhật hồ sơ quét (CHUACOGIAY -> thật) + thông tin đăng ký/chủ theo tờ/thửa.

KHÔNG có chức năng xóa đơn — chỉ đọc và cập nhật.

Luồng xử lý:
1. Mở Chrome và đăng nhập MPLIS.
2. Lấy cookie + __RequestVerificationToken từ Chrome.
3. Đọc file Excel có cột soto, sothua, loaidat, tenfile, sogcn.
4. Tra AdvancedSearchTinhHinhDangKy theo mã xã, số tờ, số thửa -> lấy tất cả
   tinhHinhDangKyId khớp (1 thửa có thể có nhiều đơn — xử lý tất cả).
5. Lấy chi tiết đơn (GetThongTinDangKyByTinhHinhDangKyIds, kèm hồ sơ quét)
   trong 1 lần gọi.
6. Tìm file CHUACOGIAY trong hồ sơ quét:
   - Có tick "Đẩy hồ sơ quét": upload file PDF thật thay file CHUACOGIAY.
   - Không tick: chỉ sửa moTa (hậu tố -DDK) + loaiHoSoQuet của file đó,
     KHÔNG đính kèm file mới.
7. Cập nhật thông tin đăng ký (coQuyenQuanLy=True, thoiDiemDangKyLanDau)
   + đổi chủ (nếu nhập ID thông tin chủ).
8. Ghi kết quả ra Excel, tự lưu sau mỗi 5 dòng.

Cài thư viện:
    python -m pip install requests selenium webdriver-manager openpyxl
"""

from __future__ import annotations

import copy
import json
import os
import re
import threading
import time
import unicodedata
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# ============================ CẤU HÌNH ============================

BASE_URL = "https://dla.mplis.gov.vn"

REFERER_LOGIN = f"{BASE_URL}/dc/DonDangKy/KeKhaiDangKyV2"
REFERER_API = REFERER_LOGIN

URL_TIM_DON = f"{BASE_URL}/dc/DangKyAjax/AdvancedSearchTinhHinhDangKy"
URL_CHI_TIET = f"{BASE_URL}/dc/DangKyAjax/GetThongTinDangKyByTinhHinhDangKyIds"
URL_UPDATE_HOSOQUET = f"{BASE_URL}/dc/HoSoQuetAjax/UpdateHoSoQuetExistFile"
URL_UPDATE_THONG_TIN_DANG_KY = f"{BASE_URL}/dc/DangKyAjax/UpdateThongTinDangKy"

PAGE_SIZE = 10
TIMEOUT = 120
SAVE_EVERY_ROWS = 5
REQUEST_DELAY_SECONDS = 0.15

# True = chỉ kiểm tra, không update thật.
DRY_RUN = False

# Các loại hồ sơ quét (theo dropdown loaiHoSoQuet trên MPLIS).
LOAI_HO_SO_QUET_OPTIONS = {
    0: "Giấy tờ",
    1: "Giấy chứng nhận",
    2: "Đơn đăng ký",
    3: "Thông báo xác nhận đăng ký",
}

OUTPUT_HEADERS = [
    "stt",
    "soto",
    "sothua",
    "loaidat",
    "tenfile",
    "sogcn",
    "tinhhinhdangkyid",
    "hosoquetid",
    "mota moi",
    "chu su dung",
    "trang thai",
    "ghi chu",
]


# ============================ HELPER ============================


def lay_token_tu_trang(driver: webdriver.Chrome) -> str:
    js = """
    return (
        document.querySelector('input[name="__RequestVerificationToken"]')?.value ||
        document.querySelector('input[name="__requestverificationtoken"]')?.value ||
        document.querySelector('meta[name="__RequestVerificationToken"]')?.content ||
        document.querySelector('meta[name="__requestverificationtoken"]')?.content ||
        document.querySelector('meta[name="RequestVerificationToken"]')?.content ||
        ''
    );
    """
    value = driver.execute_script(js)
    return str(value or "").strip()


def chuan_hoa_gia_tri_excel(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def rut_gon_text(value: Any, limit: int = 500) -> str:
    text = str(value or "").strip().replace("\r", " ").replace("\n", " ")
    if len(text) > limit:
        return text[:limit] + "..."
    return text


def safe_int(value: Any, default: int = 0) -> int:
    try:
        if value is None or value == "":
            return default
        return int(value)
    except Exception:
        return default


def now_iso_z() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def lay_thong_tin_ho_so_id(detail: dict[str, Any] | None, hoso: dict[str, Any] | None = None) -> int:
    """Lấy thongTinHoSoId từ hoso hoặc từ detail (DanhSachThongTinHoSo, thongTinHoSoId, ThongTinHoSo)."""
    if hoso:
        tid = safe_int(hoso.get("thongTinHoSoId"))
        if tid:
            return tid

    if not detail or not isinstance(detail, dict):
        return 0

    # 1. Từ DanhSachThongTinHoSo trong detail
    ds = detail.get("DanhSachThongTinHoSo") or []
    if isinstance(ds, list) and ds:
        for item in ds:
            if isinstance(item, dict):
                tid = safe_int(item.get("thongTinHoSoId"))
                if tid:
                    return tid

    # 2. Từ thongTinHoSoId top-level trong detail
    tid = safe_int(detail.get("thongTinHoSoId"))
    if tid:
        return tid

    # 3. Từ ThongTinHoSo object
    tths = detail.get("ThongTinHoSo") or {}
    if isinstance(tths, dict):
        tid = safe_int(tths.get("thongTinHoSoId"))
        if tid:
            return tid

    # 4. Từ TinhHinhDangKy
    thdk = detail.get("TinhHinhDangKy") or {}
    if isinstance(thdk, dict):
        tid = safe_int(thdk.get("thongTinHoSoId"))
        if tid:
            return tid

    # 5. Từ ListHoSoQuet nếu có
    for h in detail.get("ListHoSoQuet") or []:
        if isinstance(h, dict):
            tid = safe_int(h.get("thongTinHoSoId"))
            if tid:
                return tid

    return 0


def dotnet_date_to_iso(value: Any) -> Any:
    """Chuyển /Date(1780843978377)/ sang ISO UTC. Nếu không phải /Date(...)/ thì giữ nguyên."""
    if not isinstance(value, str):
        return value

    match = re.search(r"/Date\((-?\d+)\)/", value)
    if not match:
        return value

    ms = int(match.group(1))
    dt = datetime.fromtimestamp(ms / 1000, tz=timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{int(dt.microsecond / 1000):03d}Z"


def convert_dates_recursive(obj: Any) -> Any:
    """Convert toàn bộ ngày /Date(...)/ trong dict/list sang ISO."""
    if isinstance(obj, dict):
        return {k: convert_dates_recursive(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [convert_dates_recursive(x) for x in obj]
    return dotnet_date_to_iso(obj)


def ddmmyyyy_to_iso_utc_start_of_day_vn(date_str: str) -> str:
    """Nhập dd/mm/yyyy theo ngày Việt Nam. Ví dụ 23/04/2026 -> 2026-04-22T17:00:00.000Z."""
    date_str = str(date_str).strip()
    dt_vn = datetime.strptime(date_str, "%d/%m/%Y")
    tz_vn = timezone(timedelta(hours=7))
    dt_vn = dt_vn.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=tz_vn)
    dt_utc = dt_vn.astimezone(timezone.utc)
    return dt_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")


def remove_accents(input_str: str) -> str:
    """Loại bỏ dấu tiếng Việt để so khớp mềm."""
    if not input_str:
        return ""
    nfkd_form = unicodedata.normalize("NFKD", input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])


def tach_thong_tin_tu_chuoi_madon(val: Any) -> tuple[str, str]:
    """
    Tự động bóc tách mã đơn (tinhHinhDangKyId) và thông tin GCN (nếu có)
    từ chuỗi mô tả/thông báo kiểm tra hoặc lỗi.
    Ví dụ:
    "Tình hình đăng ký 13456477 có giấy chứng nhận 2502899_11 có hồ sơ quét không có dữ liệu tập tin; Chưa đồng bộ thông tin ba khối (không có hồ sơ quét)"
    -> id_don = "13456477", gcn_hint = "2502899_11"
    """
    if val is None:
        return "", ""
    s = str(val).strip()
    if not s:
        return "", ""

    # 1. Trường hợp là số nguyên hoặc float thuần túy (e.g. 13456477 hoặc 13456477.0)
    if re.fullmatch(r"\d+(\.0+)?", s):
        return str(int(float(s))), ""

    id_don = ""
    gcn_hint = ""

    s_clean = remove_accents(s).lower()

    # Pattern 1: Tình hình đăng ký <mã> / Đơn đăng ký <mã> / Mã đơn <mã> / THDK <mã>
    m_id = re.search(
        r"(?:tinh\s*hinh\s*dang\s*ky|don\s*dang\s*ky|ma\s*don|id\s*don|thdk|tinhhinhdangky)\D*?(\d{5,10})",
        s_clean,
    )
    if m_id:
        id_don = m_id.group(1)
    else:
        # Pattern 2: Tìm chữ số sau 'id' hoặc 'don'
        m_id2 = re.search(r"\b(?:id|don)\D*?(\d{5,10})\b", s_clean)
        if m_id2:
            id_don = m_id2.group(1)
        else:
            # Pattern 3: Tìm dãy số 6-10 chữ số đầu tiên trong chuỗi
            m_digits = re.search(r"\b\d{6,10}\b", s)
            if m_digits:
                id_don = m_digits.group(0)

    # 2. Tách thông tin GCN nếu có trong chuỗi (ví dụ: 'giấy chứng nhận 2502899_11')
    m_gcn = re.search(r"gi[aấ]y\s*ch[uứ]ng\s*nh[aậ]n\s*([A-Za-z0-9_\-]+)", s, re.IGNORECASE)
    if m_gcn:
        gcn_hint = m_gcn.group(1).strip()

    return id_don, gcn_hint


def normalize_code(s: Any) -> str:
    """Bỏ toàn bộ khoảng trắng, gạch nối, gạch dưới, dấu chấm và chuyển hoa để so sánh mã."""
    if not s:
        return ""
    return re.sub(r"[\s\-_.]+", "", str(s)).upper()


def normalize_sogcn_key(sogcn: Any) -> str:
    """Gom nhóm theo số GCN, bỏ khác biệt hoa/thường và khoảng trắng thừa."""
    return " ".join(str(sogcn or "").strip().upper().split())


def make_group_key(item: dict[str, Any]) -> str:
    """
    Tạo khóa gom nhóm: phân biệt theo ID đơn / tờ thửa VÀ số GCN / tên file
    để không bị gộp nhầm các GCN khác nhau trong cùng 1 đơn.
    """
    id_don = str(item.get("id_don") or "").strip()
    so_to = str(item.get("soto") or "").strip()
    so_thua = str(item.get("sothua") or "").strip()
    so_gcn = normalize_code(item.get("sogcn"))
    ten_file = normalize_code(Path(item.get("tenfile") or "").stem)
    so_vao_so = normalize_code(item.get("sovaoso"))
    gcn_id = str(item.get("gcn_id") or "").strip()

    gcn_token = gcn_id or so_gcn or so_vao_so or ten_file

    if id_don:
        if gcn_token:
            return f"ID_{id_don}_{gcn_token}"
        return f"ROW_{item.get('excel_row')}"
    elif so_to and so_thua:
        if gcn_token:
            return f"TO_{so_to}_THUA_{so_thua}_{gcn_token}"
        return f"ROW_{item.get('excel_row')}"
    return f"ROW_{item.get('excel_row')}"


def tim_file_pdf_that(folder_upload: str, ten_file: str, so_gcn: str = "") -> str | None:
    """
    Tìm đường dẫn file PDF thực tế trong folder_upload.
    Hỗ trợ:
    - ten_file có hoặc không có đuôi .pdf
    - so_gcn có hoặc không có đuôi .pdf
    - tìm case-insensitive và bỏ qua khác biệt dấu cách/ký tự gạch nối.
    """
    if ten_file and os.path.isabs(ten_file) and os.path.isfile(ten_file):
        return ten_file

    if not folder_upload or not os.path.isdir(folder_upload):
        return None

    candidates = []
    if ten_file:
        candidates.append(ten_file)
        if not ten_file.lower().endswith(".pdf"):
            candidates.append(f"{ten_file}.pdf")
            candidates.append(f"{ten_file}.PDF")
    if so_gcn:
        candidates.append(so_gcn)
        if not so_gcn.lower().endswith(".pdf"):
            candidates.append(f"{so_gcn}.pdf")
            candidates.append(f"{so_gcn}.PDF")

    for c in candidates:
        full_p = os.path.join(folder_upload, c)
        if os.path.isfile(full_p):
            return full_p

    def normalize_name(s: str) -> str:
        base = Path(s).stem
        return re.sub(r"[\s_-]+", "", base).lower()

    target_norms = set()
    if ten_file:
        target_norms.add(normalize_name(ten_file))
    if so_gcn:
        target_norms.add(normalize_name(so_gcn))

    try:
        for fname in os.listdir(folder_upload):
            if fname.lower().endswith(".pdf"):
                if normalize_name(fname) in target_norms:
                    return os.path.join(folder_upload, fname)
    except Exception:
        pass

    return None


ALIASES_ID_DON = {
    "madon", "ma_don", "tinhhinhdangkyid", "id_don", "iddon", "id",
    "tinh_hinh_dang_ky_id", "mã đơn", "id đơn", "mã_đơn", "id_đơn"
}
ALIASES_SO_TO = {"soto", "so_to", "số tờ", "sohieutobando", "so_hieu_to_ban_do", "số_tờ"}
ALIASES_SO_THUA = {"sothua", "so_thua", "số thửa", "sothututhua", "so_thu_tu_thua", "số_thửa"}
ALIASES_TEN_FILE = {"tenfile", "ten_file", "tên file", "file", "teptin", "tep_tin", "tên_file"}
ALIASES_SO_GCN = {
    "sogcn", "so_gcn", "số gcn", "sophathanh", "so_phat_hanh", "gcn", "số_gcn",
    "so_phat_hanh_gcn", "sophathanhgcn", "so_gcn_moi", "so_seri", "soseri"
}
ALIASES_SO_VAO_SO = {
    "sovaoso", "so_vao_so", "số vào sổ", "sovaosocu", "so_vao_so_cu"
}
ALIASES_GCN_ID = {
    "giaychungnhanid", "id_gcn", "gcn_id", "giay_chung_nhan_id", "idgcn"
}
ALIASES_LOAI_DAT = {"loaidat", "loai_dat", "loại đất", "mucdich", "muc_dich", "loại_đất"}


def doc_danh_sach(file_path: str) -> list[dict[str, Any]]:
    """
    Đọc file Excel hỗ trợ 2 dạng:
    - Dạng 1: có cột ID đơn (madon, tinhhinhdangkyid, id_don) + tenfile/sogcn.
    - Dạng 2: có cặp cột soto, sothua + loaidat/tenfile/sogcn.
    Hỗ trợ thêm các cột nhận diện GCN cụ thể: sogcn/sophathanh, sovaoso, gcn_id.
    """
    extension = Path(file_path).suffix.lower()
    if extension not in {".xlsx", ".xlsm"}:
        raise ValueError("Chỉ hỗ trợ file Excel .xlsx hoặc .xlsm.")

    try:
        workbook = load_workbook(file_path, data_only=True)
    except BadZipFile as exc:
        raise ValueError(
            "File được chọn không phải file Excel .xlsx/.xlsm hợp lệ hoặc file đã bị hỏng."
        ) from exc

    try:
        worksheet = workbook.active

        headers: dict[str, int] = {}
        for column_index in range(1, worksheet.max_column + 1):
            value = worksheet.cell(row=1, column=column_index).value
            if value:
                clean_header = str(value).strip().lower()
                headers[clean_header] = column_index

        def tim_cot(aliases: set[str]) -> int | None:
            for alias in aliases:
                if alias in headers:
                    return headers[alias]
            return None

        col_id_don    = tim_cot(ALIASES_ID_DON)
        col_so_to     = tim_cot(ALIASES_SO_TO)
        col_so_thua   = tim_cot(ALIASES_SO_THUA)
        col_ten_file  = tim_cot(ALIASES_TEN_FILE)
        col_so_gcn    = tim_cot(ALIASES_SO_GCN)
        col_so_vao_so = tim_cot(ALIASES_SO_VAO_SO)
        col_gcn_id    = tim_cot(ALIASES_GCN_ID)
        col_loai_dat  = tim_cot(ALIASES_LOAI_DAT)

        has_id_don = col_id_don is not None
        has_to_thua = col_so_to is not None and col_so_thua is not None
        has_file = col_ten_file is not None or col_so_gcn is not None or col_gcn_id is not None

        if not has_id_don and not has_to_thua:
            raise ValueError(
                "File Excel cần có cột Mã đơn/ID đơn ('madon' hoặc 'tinhhinhdangkyid') "
                "HOẶC cặp cột ('soto', 'sothua')."
            )

        if not has_file and not has_id_don:
            raise ValueError(
                "File Excel cần có ít nhất một trong các cột tên file hoặc số GCN ('tenfile', 'sogcn')."
            )

        rows: list[dict[str, Any]] = []

        for row_index in range(2, worksheet.max_row + 1):
            id_don_val    = worksheet.cell(row=row_index, column=col_id_don).value if col_id_don else None
            so_to_val     = worksheet.cell(row=row_index, column=col_so_to).value if col_so_to else None
            so_thua_val   = worksheet.cell(row=row_index, column=col_so_thua).value if col_so_thua else None
            loai_dat_val  = worksheet.cell(row=row_index, column=col_loai_dat).value if col_loai_dat else None
            ten_file_val  = worksheet.cell(row=row_index, column=col_ten_file).value if col_ten_file else None
            so_gcn_val    = worksheet.cell(row=row_index, column=col_so_gcn).value if col_so_gcn else None
            so_vao_so_val = worksheet.cell(row=row_index, column=col_so_vao_so).value if col_so_vao_so else None
            gcn_id_val    = worksheet.cell(row=row_index, column=col_gcn_id).value if col_gcn_id else None

            if not any([id_don_val, so_to_val, so_thua_val, loai_dat_val, ten_file_val, so_gcn_val, so_vao_so_val, gcn_id_val]):
                continue

            raw_id_don_str = chuan_hoa_gia_tri_excel(id_don_val)
            extracted_id, extracted_gcn = tach_thong_tin_tu_chuoi_madon(id_don_val)
            id_don_str    = extracted_id if extracted_id else raw_id_don_str

            so_to_str     = chuan_hoa_gia_tri_excel(so_to_val)
            so_thua_str   = chuan_hoa_gia_tri_excel(so_thua_val)
            loai_dat_str  = chuan_hoa_gia_tri_excel(loai_dat_val)
            ten_file_str  = chuan_hoa_gia_tri_excel(ten_file_val)
            so_gcn_str    = chuan_hoa_gia_tri_excel(so_gcn_val)
            so_vao_so_str = chuan_hoa_gia_tri_excel(so_vao_so_val)
            gcn_id_str    = chuan_hoa_gia_tri_excel(gcn_id_val)

            # Nếu trong chuỗi mã đơn có số GCN và cột GCN chưa có, tự điền
            if not so_gcn_str and extracted_gcn:
                so_gcn_str = extracted_gcn

            # Nếu tenfile rỗng nhưng có sogcn, gán tenfile mặc định = sogcn.pdf
            if not ten_file_str and so_gcn_str:
                ten_file_str = f"{so_gcn_str}.pdf"

            rows.append(
                {
                    "excel_row": row_index,
                    "id_don":    id_don_str,
                    "soto":      so_to_str,
                    "sothua":    so_thua_str,
                    "loaidat":   loai_dat_str,
                    "tenfile":   ten_file_str,
                    "sogcn":     so_gcn_str,
                    "sovaoso":   so_vao_so_str,
                    "gcn_id":    gcn_id_str,
                }
            )

        if not rows:
            raise ValueError("File Excel không có dữ liệu để xử lý.")

        return rows
    finally:
        workbook.close()



class ExcelResultWriter:
    def __init__(self, output_path: str):
        self.output_path = output_path
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "KetQua"
        self.stt = 0

        self._setup_sheet()

    def _setup_sheet(self) -> None:
        self.worksheet.append(OUTPUT_HEADERS)

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="B7B7B7")

        for cell in self.worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        widths = {1: 6, 2: 10, 3: 10, 4: 14, 5: 30, 6: 16, 7: 16, 8: 14, 9: 30, 10: 24, 11: 16, 12: 55}
        for column_index, width in widths.items():
            self.worksheet.column_dimensions[get_column_letter(column_index)].width = width

        self.worksheet.freeze_panes = "A2"
        self.worksheet.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}1"
        self.worksheet.row_dimensions[1].height = 30

    def append_result(self, row: dict[str, Any]) -> None:
        self.stt += 1
        self.worksheet.append(
            [
                self.stt,
                row.get("soto"),
                row.get("sothua"),
                row.get("loaidat"),
                row.get("tenfile"),
                row.get("sogcn"),
                row.get("tinhhinhdangkyid"),
                row.get("hosoquetid"),
                row.get("mota_moi"),
                row.get("chu_su_dung"),
                row.get("trang_thai"),
                row.get("ghi_chu"),
            ]
        )

        row_index = self.worksheet.max_row
        thin = Side(style="thin", color="D9D9D9")

        for cell in self.worksheet[row_index]:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        trang_thai = str(row.get("trang_thai") or "")

        if trang_thai == "Thành công":
            fill = PatternFill("solid", fgColor="C6EFCE")
            font = Font(color="006100")
        elif trang_thai == "Bỏ qua":
            fill = PatternFill("solid", fgColor="FFF2CC")
            font = None
        elif trang_thai in ("Lỗi", "Lỗi ngoài"):
            fill = PatternFill("solid", fgColor="F4CCCC")
            font = Font(color="9C0006")
        elif trang_thai == "DRY_RUN":
            fill = PatternFill("solid", fgColor="D9EAF7")
            font = None
        else:
            fill = None
            font = None

        cell_trang_thai = self.worksheet.cell(row=row_index, column=11)
        if fill:
            cell_trang_thai.fill = fill
        if font:
            cell_trang_thai.font = font

    def save(self) -> None:
        output_parent = Path(self.output_path).resolve().parent
        output_parent.mkdir(parents=True, exist_ok=True)
        self.workbook.save(self.output_path)

    def close(self) -> None:
        try:
            self.workbook.close()
        except Exception:
            pass


# ============================ CORE API ============================


class MplisClient:
    def __init__(self, log_fn):
        self.log = log_fn
        self.session: requests.Session | None = None
        self.driver: webdriver.Chrome | None = None

    # ---------- login ----------
    def open_browser_and_fill_login(self, username: str, password: str) -> None:
        options = Options()
        options.add_argument("--start-maximized")

        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.get(REFERER_LOGIN)
        time.sleep(2)

        try:
            inputs = self.driver.find_elements(By.CSS_SELECTOR, "input")
            user_box = None
            pass_box = None

            for inp in inputs:
                input_type = (inp.get_attribute("type") or "").lower()

                if user_box is None and input_type in {"text", "email"}:
                    user_box = inp

                if pass_box is None and input_type == "password":
                    pass_box = inp

            if user_box and pass_box:
                user_box.clear()
                user_box.send_keys(username)
                pass_box.clear()
                pass_box.send_keys(password)
                pass_box.send_keys(Keys.ENTER)
                self.log("Đã điền thông tin đăng nhập, chờ trang tải...")
            else:
                self.log("Không nhận dạng được form đăng nhập, hãy đăng nhập tay trên Chrome.")

        except Exception as exc:
            self.log(f"Không tự điền được form đăng nhập ({exc}), hãy đăng nhập tay.")

    def build_session_from_browser(self) -> None:
        if not self.driver:
            raise RuntimeError("Chưa mở trình duyệt.")

        token = lay_token_tu_trang(self.driver)

        if not token:
            raise RuntimeError(
                "Không lấy được token. Hãy bảo đảm đã đăng nhập thành công và đang mở trang MPLIS."
            )

        session = requests.Session()
        user_agent = self.driver.execute_script("return navigator.userAgent;")

        session.headers.update(
            {
                "User-Agent": user_agent,
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Accept-Language": "vi-VN,vi;q=0.9",
                "X-Requested-With": "XMLHttpRequest",
                "Origin": BASE_URL,
                "Referer": REFERER_API,
                "__requestverificationtoken": token,
            }
        )

        for cookie in self.driver.get_cookies():
            session.cookies.set(
                name=cookie["name"],
                value=cookie["value"],
                domain=cookie.get("domain"),
                path=cookie.get("path", "/"),
            )

        self.session = session
        self.log("✅ Đã lấy session + token thành công.")

    def build_session_from_manual(self, token: str, cookie_raw: str) -> None:
        """
        Tạo session trực tiếp từ token + cookie dán thủ công (copy từ F12 > Network
        > Request Headers), không cần mở Chrome/Selenium.
        """
        token = token.strip()
        cookie_raw = cookie_raw.strip()

        if not token:
            raise RuntimeError("Chưa nhập token.")
        if not cookie_raw:
            raise RuntimeError("Chưa nhập cookie.")

        session = requests.Session()
        session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
                ),
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Accept-Language": "vi-VN,vi;q=0.9",
                "X-Requested-With": "XMLHttpRequest",
                "Origin": BASE_URL,
                "Referer": REFERER_API,
                "__requestverificationtoken": token,
            }
        )

        so_cookie_da_nhan = 0
        for part in cookie_raw.split(";"):
            part = part.strip()
            if not part or "=" not in part:
                continue
            name, _, value = part.partition("=")
            name = name.strip()
            value = value.strip()
            if not name:
                continue
            session.cookies.set(name=name, value=value)
            so_cookie_da_nhan += 1

        if so_cookie_da_nhan == 0:
            raise RuntimeError("Không đọc được cookie nào — kiểm tra lại chuỗi cookie đã dán.")

        self.session = session
        self.log(f"✅ Đã tạo session từ token/cookie thủ công ({so_cookie_da_nhan} cookie).")

    def close_browser(self) -> None:
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ---------- request helpers ----------
    def _require_session(self) -> requests.Session:
        if self.session is None:
            raise RuntimeError("Chưa lấy session từ Chrome.")
        return self.session

    @staticmethod
    def _check_response(response: requests.Response, api_name: str) -> None:
        if response.status_code in {301, 302, 303, 307, 308} or response.headers.get("Location"):
            raise RuntimeError(
                f"{api_name} bị chuyển hướng (HTTP {response.status_code}). "
                "Phiên đăng nhập có thể đã hết hạn."
            )

        if response.status_code == 404:
            raise RuntimeError(f"{api_name}: endpoint không tồn tại (404).")

        response.raise_for_status()

    def _response_json(self, response: requests.Response, api_name: str) -> Any:
        self._check_response(response, api_name)

        try:
            return response.json()
        except requests.JSONDecodeError as exc:
            raise RuntimeError(
                f"{api_name} không trả JSON (HTTP {response.status_code}): {response.text[:500]}"
            ) from exc

    # ---------- search đơn ----------
    @staticmethod
    def build_search_payload(
        xa_id: str,
        so_to: str,
        so_thua: str,
        start: int = 0,
        length: int = PAGE_SIZE,
        draw: int = 1,
    ) -> dict[str, str]:
        payload = {
            "draw": str(draw),
            "order[0][column]": "5",
            "order[0][dir]": "desc",
            "start": str(start),
            "length": str(length),
            "search[value]": "",
            "search[regex]": "false",
            "model[xaId]": xa_id,
            "model[huyenId]": "",
            "model[tinhHinhDangKyId]": "",
            "model[maDon]": "",
            "model[soThuTu]": "",
            "model[ngayTiepNhan]": "",
            "model[thoiDiemDangKy]": "",
            "model[loaiGiayChungNhanId]": "",
            "model[soPhatHanh]": "",
            "model[maVach]": "",
            "model[soVaoSo]": "",
            "model[soVaoSoCu]": "",
            "model[ngayVaoSo]": "",
            "model[soHoSoGoc]": "",
            "model[soHoSoGocCu]": "",
            "model[hoTen]": "",
            "model[soGiayTo]": "",
            "model[namSinh]": "",
            "model[soThuTuThua]": so_thua,
            "model[soHieuToBanDo]": so_to,
            "model[soThuTuThuaCu]": "",
            "model[soHieuToBanDoCu]": "",
            "model[soNha]": "",
            "model[diaChiChiTiet]": "",
            "model[dieuKienCapGiay]": "",
            "model[phucHoiDuLieu]": "false",
        }

        columns = [
            ("", "", "true", "false"),
            ("tinhHinhDangKyId", "tinhHinhDangKyId", "true", "true"),
            ("maDon", "maDon", "true", "true"),
            ("soThuTu", "soThuTu", "true", "true"),
            ("DaiDienKhaiTrinh", "DaiDienKhaiTrinh", "true", "false"),
            ("ngayTiepNhan", "ngayTiepNhan", "true", "true"),
            ("thoiDiemDangKy", "thoiDiemDangKy", "true", "true"),
        ]

        for index, (data, name, searchable, orderable) in enumerate(columns):
            payload[f"columns[{index}][data]"] = data
            payload[f"columns[{index}][name]"] = name
            payload[f"columns[{index}][searchable]"] = searchable
            payload[f"columns[{index}][orderable]"] = orderable
            payload[f"columns[{index}][search][value]"] = ""
            payload[f"columns[{index}][search][regex]"] = "false"

        return payload

    def tim_don_mot_trang(
        self, xa_id: str, so_to: str, so_thua: str, start: int, length: int, draw: int
    ) -> tuple[list[dict[str, Any]], int | None]:
        session = self._require_session()
        payload = self.build_search_payload(
            xa_id=xa_id, so_to=so_to, so_thua=so_thua, start=start, length=length, draw=draw
        )

        response = session.post(URL_TIM_DON, data=payload, timeout=TIMEOUT, allow_redirects=False)
        result = self._response_json(response, "AdvancedSearchTinhHinhDangKy")
        rows = result.get("data") or []

        if not isinstance(rows, list):
            raise RuntimeError("API tìm đơn trả trường data không phải danh sách.")

        total_value = result.get("recordsFiltered")
        try:
            total = int(total_value)
        except (TypeError, ValueError):
            total = None

        return rows, total

    def tim_tat_ca_don(self, xa_id: str, so_to: str, so_thua: str) -> list[dict[str, Any]]:
        all_rows: list[dict[str, Any]] = []
        start = 0
        draw = 1

        while True:
            rows, total = self.tim_don_mot_trang(
                xa_id=xa_id, so_to=so_to, so_thua=so_thua, start=start, length=PAGE_SIZE, draw=draw
            )

            if not rows:
                break

            all_rows.extend(rows)
            start += len(rows)
            draw += 1

            if total is not None and start >= total:
                break

            if len(rows) < PAGE_SIZE:
                break

        unique_rows: list[dict[str, Any]] = []
        seen_ids: set[str] = set()

        for row in all_rows:
            tinh_hinh_id = row.get("tinhHinhDangKyId")
            id_text = str(tinh_hinh_id or "").strip()

            if not id_text or id_text in seen_ids:
                continue

            seen_ids.add(id_text)
            unique_rows.append(row)

        return unique_rows

    def lay_chi_tiet_don(self, tinh_hinh_dang_ky_id: int) -> dict[str, Any] | None:
        session = self._require_session()
        payload = {"getHoSoQuet": True, "tinhHinhDangKyIds": [int(tinh_hinh_dang_ky_id)]}

        response = session.post(URL_CHI_TIET, json=payload, timeout=TIMEOUT, allow_redirects=False)
        result = self._response_json(response, "GetThongTinDangKyByTinhHinhDangKyIds")

        values = result.get("value") or []
        if isinstance(values, dict):
            values = [values]

        if not isinstance(values, list) or not values:
            return None

        for item in values:
            if not isinstance(item, dict):
                continue

            tinh_hinh = item.get("TinhHinhDangKy") or {}
            current_id = tinh_hinh.get("tinhHinhDangKyId")

            if str(current_id) == str(tinh_hinh_dang_ky_id):
                return item

        if len(values) == 1 and isinstance(values[0], dict):
            return values[0]

        return None

    # ---------- cập nhật hồ sơ quét ----------
    def api_update_hosoquet_exist_file(
        self,
        file_path: str,
        hoso: dict[str, Any],
        mo_ta_moi: str,
        loai_ho_so_quet: int = 1,
        giay_chung_nhan_id: str = "",
        la_giay_chung_nhan: bool = False,
        target_file: dict[str, Any] | None = None,
        version_gcn: int = 11,
        bien_dong_id: int = 0,
        detail: dict[str, Any] | None = None,
    ) -> str:
        """
        Upload file PDF thật thay cho file hiện có trong HSQ hoặc tạo mới HSQ rỗng theo chuẩn MPLIS.
        Nếu HSQ rỗng (chưa có file hoặc hoSoQuetId == 0), đẩy payload tạo mới.
        """
        if not os.path.isfile(file_path):
            raise RuntimeError(f"Không tìm thấy file PDF: {file_path}")

        session = self._require_session()

        ho_so_quet_id_int = safe_int(hoso.get("hoSoQuetId") or hoso.get("Title"), 0)
        gcn_id_int = safe_int(giay_chung_nhan_id) if giay_chung_nhan_id else None

        tf = target_file or {}
        tp_id = safe_int(tf.get("thanhPhanHoSoQuetId")) or None
        tp_nid = tf.get("thanhPhanHoSoQuetNId") or str(uuid.uuid4())
        ver_gcn = safe_int(tf.get("versionGiayChungNhan") or version_gcn, 11)

        is_gcn = la_giay_chung_nhan or (safe_int(loai_ho_so_quet, 1) == 1)

        # Kiểm tra chế độ: Nếu HSQ rỗng (chưa có file hoặc hoSoQuetId == 0) -> Đẩy payload tạo mới
        is_empty_or_new = (target_file is None) or (tp_id is None) or (ho_so_quet_id_int == 0)

        if is_empty_or_new:
            thong_tin_ho_so_id = safe_int(hoso.get("thongTinHoSoId") or lay_thong_tin_ho_so_id(detail, hoso))
            ho_so_quet = {
                "thongTinHoSoId": thong_tin_ho_so_id,
                "TuiHoSo": None,
                "tuiHoSoId": 0,
                "hoSoQuetId": 0,
            }
            info_ho_so_quet = {
                "loaiHoSoQuet": safe_int(loai_ho_so_quet, 1),
                "laGiayToVeNguonGoc": True,
                "giayChungNhanId": str(giay_chung_nhan_id).strip() if giay_chung_nhan_id else "",
                "moTa": mo_ta_moi,
                "tenGiayTo": "",
                "trichYeu": "",
                "versionGiayChungNhan": ver_gcn,
                "laGiayChungNhan": True if is_gcn else False,
                "__id": str(uuid.uuid4()),
                "files": None,
            }
        else:
            ho_so_quet = {
                "hoSoQuetId":       ho_so_quet_id_int,
                "thongTinHoSoId":   safe_int(hoso.get("thongTinHoSoId")),
                "tinhHinhDangKyId": safe_int(hoso.get("tinhHinhDangKyId")),
                "bienDongId":       safe_int(hoso.get("bienDongId") or bien_dong_id),
                "xaId":             safe_int(hoso.get("xaId")),
                "CreatedDate":      hoso.get("CreatedDate") or now_iso_z(),
                "ModifiedDate":     now_iso_z(),
                "Id":               str(hoso.get("Id") or uuid.uuid4()),
                "Title":            str(hoso.get("Title") or ho_so_quet_id_int),
                "Name":             hoso.get("Name"),
                "Path":             hoso.get("Path"),
                "ParentPath":       hoso.get("ParentPath"),
                "_id":              1,
                "TuiHoSo":          None,
                "tuiHoSoId":        safe_int(hoso.get("tuiHoSoId"), 0),
            }

            info_ho_so_quet = {
                "nodeId":               "",
                "deleteId":             None,
                "moTa":                 mo_ta_moi,
                "tenGiayTo":            None,
                "trichYeu":             None,
                "isOldFile":            False,
                "laGiayToVeNguonGoc":   True if is_gcn else False,
                "laGiayChungNhan":      is_gcn,
                "giayChungNhanId":      gcn_id_int,
                "workflowNodeId":       "",
                "versionGiayChungNhan": ver_gcn,
                "loaiHoSoQuet":         safe_int(loai_ho_so_quet, 1 if is_gcn else 2),
                "daKySo":               False,
                "duongDan":             None,
                "duLieu":               None,
                "GiayChungNhan":        None,
                "hoSoQuetId":           ho_so_quet_id_int,
                "thanhPhanHoSoQuetNId": tp_nid,
                "thanhPhanHoSoQuetId":  tp_id,
                "_id":                  1,
                "files":                None,
                "__id":                 tf.get("__id") or str(uuid.uuid4()),
            }

        data = {
            "hoSoQuet":         json.dumps(ho_so_quet, ensure_ascii=False),
            "infoHoSoQuet_1":   json.dumps(info_ho_so_quet, ensure_ascii=False),
            "count":            "1",
            "isLuuKhoHoSoQuet": "false",
        }

        headers = dict(session.headers)
        headers.pop("Content-Type", None)

        with open(file_path, "rb") as f:
            files = {"fileHoSoQuet_1": (os.path.basename(file_path), f, "application/pdf")}
            response = session.post(
                URL_UPDATE_HOSOQUET, data=data, files=files, headers=headers, timeout=TIMEOUT
            )

        result = self._response_json(response, "UpdateHoSoQuetExistFile (upload)")

        if not isinstance(result, dict) or not result.get("success"):
            raise RuntimeError(f"Upload HSQ không thành công: {rut_gon_text(result)}")

        return "Upload HSQ OK"

    def api_update_hosoquet_metadata_only(
        self,
        hoso: dict[str, Any],
        target_file: dict[str, Any] | None,
        mo_ta_moi: str,
        loai_ho_so_quet: int = 2,
        giay_chung_nhan_id: str = "",
        la_giay_chung_nhan: bool = False,
        version_gcn: int = 11,
        bien_dong_id: int = 0,
    ) -> str:
        """
        Sửa moTa + loaiHoSoQuet của ĐÚNG file target trong HSQ, KHÔNG đính kèm file
        mới. Gửi lại toàn bộ node hiện có kèm files=null — server chỉ cập nhật
        metadata, giữ nguyên file vật lý cũ. Các node/file khác giữ nguyên.
        """
        session = self._require_session()

        wrapper = hoso.get("ListFileHoSoQuet") or {}
        list_file = wrapper.get("ListFileHoSoQuet") or []
        if not list_file:
            raise RuntimeError("HSQ không có ListFileHoSoQuet để cập nhật metadata")

        ho_so_quet_id_int = safe_int(hoso.get("hoSoQuetId") or hoso.get("Title"))
        gcn_id_int = safe_int(giay_chung_nhan_id) if giay_chung_nhan_id else None
        is_gcn = la_giay_chung_nhan or (safe_int(loai_ho_so_quet, 2) == 1)

        hoso_core = copy.deepcopy(hoso)
        hoso_core.pop("ListFileHoSoQuet", None)
        hoso_core["hoSoQuetId"] = ho_so_quet_id_int
        hoso_core["bienDongId"] = safe_int(hoso.get("bienDongId") or bien_dong_id)

        data = {
            "hoSoQuet":         json.dumps(hoso_core, ensure_ascii=False),
            "count":            str(len(list_file)),
            "isLuuKhoHoSoQuet": "false",
        }

        for i, node in enumerate(list_file, start=1):
            n = copy.deepcopy(node)
            n["files"] = None
            is_target = False
            if target_file is not None:
                if node is target_file:
                    is_target = True
                elif (
                    safe_int(node.get("thanhPhanHoSoQuetId"))
                    and safe_int(node.get("thanhPhanHoSoQuetId")) == safe_int(target_file.get("thanhPhanHoSoQuetId"))
                ):
                    is_target = True
                elif (
                    node.get("thanhPhanHoSoQuetNId")
                    and node.get("thanhPhanHoSoQuetNId") == target_file.get("thanhPhanHoSoQuetNId")
                ):
                    is_target = True

            if is_target:
                n["moTa"] = mo_ta_moi
                n["loaiHoSoQuet"] = safe_int(loai_ho_so_quet, 1 if is_gcn else 2)
                n["laGiayChungNhan"] = is_gcn
                n["laGiayToVeNguonGoc"] = True if is_gcn else False
                if gcn_id_int:
                    n["giayChungNhanId"] = gcn_id_int
                n["versionGiayChungNhan"] = safe_int(version_gcn, 11)
            data[f"infoHoSoQuet_{i}"] = json.dumps(n, ensure_ascii=False)

        headers = dict(session.headers)
        headers.pop("Content-Type", None)

        response = session.post(URL_UPDATE_HOSOQUET, data=data, headers=headers, timeout=TIMEOUT)
        result = self._response_json(response, "UpdateHoSoQuetExistFile (metadata)")

        if not isinstance(result, dict) or not result.get("success"):
            raise RuntimeError(f"Update metadata HSQ không thành công: {rut_gon_text(result)}")

        return "Update metadata HSQ OK"

    def api_update_thong_tin_dang_ky(self, payload: dict[str, Any], debug_key: str = "") -> str:
        """
        Cập nhật coQuyenQuanLy/thoiDiemDangKyLanDau (+ chủ nếu payload đã set).
        Endpoint nhận body bọc trong key "thongTinDangKy" — thử đúng dạng bọc
        trước, fallback raw không bọc nếu thất bại.
        """
        session = self._require_session()

        headers = dict(session.headers)
        headers["Content-Type"] = "application/json; charset=UTF-8"

        payload_variants = [
            ("json_wrapper_thongTinDangKy", {"thongTinDangKy": payload}),
            ("json_raw_top_level", payload),
        ]

        last_error = ""

        for mode, body in payload_variants:
            try:
                debug_dir = Path("debug_payload")
                debug_dir.mkdir(parents=True, exist_ok=True)
                suffix = f"_{debug_key}" if debug_key else ""
                with open(
                    debug_dir / f"UpdateThongTinDangKy_{mode}{suffix}.json", "w", encoding="utf-8"
                ) as f:
                    json.dump(body, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

            response = session.post(
                URL_UPDATE_THONG_TIN_DANG_KY,
                data=json.dumps(body, ensure_ascii=False),
                headers=headers,
                timeout=TIMEOUT,
                allow_redirects=False,
            )

            try:
                js = self._response_json(response, "UpdateThongTinDangKy")
            except Exception as exc:
                last_error = str(exc)
                continue

            if isinstance(js, dict) and (
                js.get("success") is True or js.get("Success") is True or js.get("ok") is True
            ):
                return f"Update TTĐK OK ({mode})"

            last_error = rut_gon_text(js)

        raise RuntimeError(f"UpdateThongTinDangKy không thành công: {last_error}")


# ============================ XỬ LÝ 1 ĐƠN ============================


def tim_gcn_trong_don(
    list_gcn: list[dict[str, Any]],
    so_gcn: str = "",
    ten_file: str = "",
    so_vao_so: str = "",
    gcn_id: str = "",
) -> tuple[dict[str, Any] | None, str]:
    """
    Tìm chính xác GCN tương ứng trong list_gcn của đơn khi đơn có 1 hoặc nhiều GCN.
    Trả về: (gcn_dict_or_None, thong_tin_khop_hoac_loi)
    """
    if not list_gcn:
        return None, "Đơn không có Giấy chứng nhận nào (ListGiayChungNhan rỗng)."

    norm_gcn_id = str(gcn_id).strip() if gcn_id else ""
    norm_so_gcn = normalize_code(so_gcn)
    norm_file = normalize_code(Path(ten_file).stem) if ten_file else ""
    norm_so_vao_so = normalize_code(so_vao_so)

    # 1. Khớp theo giayChungNhanId nếu có
    if norm_gcn_id:
        for gcn in list_gcn:
            if str(gcn.get("giayChungNhanId") or "").strip() == norm_gcn_id:
                return gcn, f"Khớp theo ID GCN {norm_gcn_id}"

    # 2. Khớp theo số phát hành (soPhatHanh)
    targets = [t for t in [norm_so_gcn, norm_file] if t]
    for target in targets:
        for gcn in list_gcn:
            sph = normalize_code(gcn.get("soPhatHanh"))
            if sph and sph == target:
                return gcn, f"Khớp số phát hành: {gcn.get('soPhatHanh')}"

    # 3. Khớp theo số vào sổ (soVaoSo)
    targets_svs = [t for t in [norm_so_vao_so, norm_so_gcn, norm_file] if t]
    for target in targets_svs:
        for gcn in list_gcn:
            svs = normalize_code(gcn.get("soVaoSo"))
            if svs and svs == target:
                return gcn, f"Khớp số vào sổ: {gcn.get('soVaoSo')}"

    # 4. Khớp theo mã vạch (maVach)
    for target in targets:
        for gcn in list_gcn:
            mv = normalize_code(gcn.get("maVach"))
            if mv and mv == target:
                return gcn, f"Khớp mã vạch: {gcn.get('maVach')}"

    # 5. Khớp mềm theo phần số (nếu người dùng chỉ gõ số, ví dụ '828383' trong 'DĐ 828383')
    for target in targets:
        digits_only = re.sub(r"\D", "", target)
        if len(digits_only) >= 5:
            matched = []
            for gcn in list_gcn:
                sph_digits = re.sub(r"\D", "", str(gcn.get("soPhatHanh") or ""))
                if digits_only in sph_digits:
                    matched.append(gcn)
            if len(matched) == 1:
                return matched[0], f"Khớp phần số {digits_only} của GCN {matched[0].get('soPhatHanh')}"

    # 6. Nếu trong đơn CHỈ CÓ DUY NHẤT 1 GCN
    if len(list_gcn) == 1:
        return list_gcn[0], f"Đơn chỉ có 1 GCN duy nhất ({list_gcn[0].get('soPhatHanh') or list_gcn[0].get('giayChungNhanId')})"

    # 7. Nếu đơn có nhiều GCN mà không khớp được
    ds_gcn_str = ", ".join(
        f"'{g.get('soPhatHanh') or g.get('soVaoSo') or g.get('giayChungNhanId')}'"
        for g in list_gcn
    )
    return None, f"Đơn có {len(list_gcn)} GCN [{ds_gcn_str}] nhưng không xác định được GCN nào khớp với '{so_gcn or ten_file}'."


def tim_hoso_de_cap_nhat(
    detail: dict[str, Any],
    target_gcn: dict[str, Any] | None = None,
    thay_moi_gcn: bool = True,
) -> dict[str, Any] | None:
    """
    detail: kết quả lay_chi_tiet_don() — ListHoSoQuet nằm ở TOP-LEVEL.
    target_gcn: dict thông tin GCN cụ thể cần gán (giúp gán đúng file trong đơn có nhiều GCN).
    """
    list_ho_so_quet = detail.get("ListHoSoQuet")
    if not isinstance(list_ho_so_quet, list) or not list_ho_so_quet:
        tths_id = lay_thong_tin_ho_so_id(detail)
        return {
            "hoso": {
                "thongTinHoSoId": tths_id,
                "TuiHoSo": None,
                "tuiHoSoId": 0,
                "hoSoQuetId": 0,
            },
            "file": None,
            "khop": "tao_moi_ho_so_quet_rong",
        }

    target_gcn_id = safe_int(target_gcn.get("giayChungNhanId")) if target_gcn else 0
    target_sph = normalize_code(target_gcn.get("soPhatHanh")) if target_gcn else ""

    if not thay_moi_gcn:
        # Chế độ cũ: Chỉ tìm file có CHUACOGIAY
        for hoso in list_ho_so_quet:
            if not isinstance(hoso, dict):
                continue
            wrapper = hoso.get("ListFileHoSoQuet") or {}
            files = wrapper.get("ListFileHoSoQuet") or []
            for f in files:
                mo_ta = (f.get("moTa") or "").upper()
                if "CHUACOGIAY" in mo_ta:
                    return {"hoso": hoso, "file": f}
        first = list_ho_so_quet[0]
        if isinstance(first, dict):
            return {"hoso": first, "file": None}
        return None

    # Chế độ mới: Gán hồ sơ quét cho đúng GCN trong đơn
    # Mức 1: Tìm file đã từng liên kết đúng với giayChungNhanId của GCN này
    if target_gcn_id:
        for hoso in list_ho_so_quet:
            if not isinstance(hoso, dict):
                continue
            wrapper = hoso.get("ListFileHoSoQuet") or {}
            files = wrapper.get("ListFileHoSoQuet") or []
            for f in files:
                if safe_int(f.get("giayChungNhanId")) == target_gcn_id:
                    return {"hoso": hoso, "file": f, "khop": "dung_gcn_id"}

    # Mức 2: Tìm file có mô tả chứa đúng số phát hành của GCN này
    if target_sph:
        for hoso in list_ho_so_quet:
            if not isinstance(hoso, dict):
                continue
            wrapper = hoso.get("ListFileHoSoQuet") or {}
            files = wrapper.get("ListFileHoSoQuet") or []
            for f in files:
                if target_sph in normalize_code(f.get("moTa")):
                    return {"hoso": hoso, "file": f, "khop": "so_phat_hanh_trong_mota"}

    # Mức 3: Tìm file mang mô tả CHUACOGIAY và CHƯA bị gắn vào GCN khác trong đơn
    for hoso in list_ho_so_quet:
        if not isinstance(hoso, dict):
            continue
        wrapper = hoso.get("ListFileHoSoQuet") or {}
        files = wrapper.get("ListFileHoSoQuet") or []
        for f in files:
            mo_ta = (f.get("moTa") or "").upper()
            f_gid = safe_int(f.get("giayChungNhanId"))
            if "CHUACOGIAY" in mo_ta and (not f_gid or f_gid == target_gcn_id):
                return {"hoso": hoso, "file": f, "khop": "chuacogiay"}

    # Mức 4: Tìm file tự do chưa gắn với bất kỳ GCN nào (giayChungNhanId rỗng và laGiayChungNhan False)
    for hoso in list_ho_so_quet:
        if not isinstance(hoso, dict):
            continue
        wrapper = hoso.get("ListFileHoSoQuet") or {}
        files = wrapper.get("ListFileHoSoQuet") or []
        for f in files:
            f_gid = safe_int(f.get("giayChungNhanId"))
            is_gcn = f.get("laGiayChungNhan") is True
            if not f_gid and not is_gcn:
                return {"hoso": hoso, "file": f, "khop": "file_tu_do"}

    # Mức 5: Nếu đơn chỉ có 1 GCN duy nhất và có file trong HSQ, lấy file GCN hoặc file đầu tiên
    if not target_gcn or (len(detail.get("ListGiayChungNhan") or []) <= 1):
        for hoso in list_ho_so_quet:
            if not isinstance(hoso, dict):
                continue
            wrapper = hoso.get("ListFileHoSoQuet") or {}
            files = wrapper.get("ListFileHoSoQuet") or []
            for f in files:
                if f.get("laGiayChungNhan") is True:
                    return {"hoso": hoso, "file": f, "khop": "gcn_duy_nhat"}
            if files:
                return {"hoso": hoso, "file": files[0], "khop": "file_dau_tien"}

    # Mức 6: Tất cả file hiện có đã thuộc về GCN khác, hoặc HSQ chưa có file:
    # Trả về hoso phù hợp với file = None để upload thành phần file mới cho GCN này
    for hoso in list_ho_so_quet:
        if isinstance(hoso, dict):
            return {"hoso": hoso, "file": None, "khop": "upload_file_moi_cho_gcn"}

    return None


def build_payload_update_ttdk(
    detail: dict[str, Any],
    ngay_dang_ky_lan_dau_ddmmyyyy: str = "",
    chu_id: str | None = None,
) -> dict[str, Any]:
    """
    Build payload UpdateThongTinDangKy trực tiếp từ 'chi tiết đơn' đã lấy sẵn.
    Chỉ sửa coQuyenQuanLy, thoiDiemDangKyLanDau (nếu nhập), và chủ (nếu có chu_id).
    """
    payload = {
        "TinhHinhDangKy": copy.deepcopy(detail.get("TinhHinhDangKy") or {}),
        "ChuSoHuu": copy.deepcopy(detail.get("ChuSoHuu") or {}),
        "TaiSan": copy.deepcopy(detail.get("TaiSan") or {}),
    }

    payload = convert_dates_recursive(payload)
    payload["TinhHinhDangKy"]["coQuyenQuanLy"] = True
    if ngay_dang_ky_lan_dau_ddmmyyyy:
        payload["TinhHinhDangKy"]["thoiDiemDangKyLanDau"] = ddmmyyyy_to_iso_utc_start_of_day_vn(
            ngay_dang_ky_lan_dau_ddmmyyyy
        )

    if chu_id:
        chu_so_huu = payload.get("ChuSoHuu") or {}
        to_chucs = chu_so_huu.get("ToChucs") or []
        ca_nhans = chu_so_huu.get("CaNhans") or []
        if to_chucs:
            to_chucs[0]["toChucId"] = safe_int(chu_id)
        elif ca_nhans:
            ca_nhans[0]["caNhanId"] = safe_int(chu_id)

    return payload


def xu_ly_mot_don(
    client: MplisClient,
    tinh_hinh_dang_ky_id: int,
    so_to: str,
    so_thua: str,
    loai_dat: str,
    ten_file: str,
    so_gcn: str,
    xa_id: str,
    folder_upload: str,
    ngay_dang_ky_lan_dau: str,
    day_hsq: bool,
    chu_id: str | None,
    loai_ho_so_quet: int,
    thay_moi_gcn: bool = True,
    so_vao_so: str = "",
    gcn_id: str = "",
) -> dict[str, Any]:
    """
    Xử lý 1 đơn (1 tinhHinhDangKyId): lấy chi tiết (kèm HSQ), tìm HSQ phù hợp,
    đẩy file thật (hoặc sửa metadata), rồi cập nhật thông tin đăng ký/chủ.
    Hỗ trợ đơn có nhiều GCN: xác định chính xác GCN cần gán và không ghi đè nhầm file khác.
    """
    result: dict[str, Any] = {
        "soto": so_to,
        "sothua": so_thua,
        "loaidat": loai_dat,
        "tenfile": ten_file,
        "sogcn": so_gcn,
        "tinhhinhdangkyid": tinh_hinh_dang_ky_id,
        "hosoquetid": "",
        "mota_moi": "",
        "chu_su_dung": "",
        "trang_thai": "Lỗi",
        "ghi_chu": "",
    }

    try:
        detail = client.lay_chi_tiet_don(tinh_hinh_dang_ky_id)
    except Exception as exc:
        result["ghi_chu"] = f"Lấy chi tiết đơn lỗi: {rut_gon_text(exc)}"
        return result

    if detail is None:
        result["ghi_chu"] = "Không lấy được chi tiết đơn."
        return result

    # Trích xuất số tờ, số thửa, mã xã nếu Excel chưa có
    tinh_hinh_dk = detail.get("TinhHinhDangKy") or {}
    if not so_to:
        so_to = str(tinh_hinh_dk.get("soHieuToBanDo") or "")
        result["soto"] = so_to
    if not so_thua:
        so_thua = str(tinh_hinh_dk.get("soThuTuThua") or "")
        result["sothua"] = so_thua
    if not xa_id:
        xa_id = str(tinh_hinh_dk.get("xaId") or "")

    chu_so_huu = detail.get("ChuSoHuu") or {}
    ca_nhans = chu_so_huu.get("CaNhans") or []
    to_chucs = chu_so_huu.get("ToChucs") or []
    if ca_nhans:
        result["chu_su_dung"] = ca_nhans[0].get("hoTen") or ""
    elif to_chucs:
        result["chu_su_dung"] = to_chucs[0].get("tenToChuc") or to_chucs[0].get("ten") or ""

    # Kiểm tra liên kết GCN trong đơn (hỗ trợ đơn có 1 hoặc nhiều GCN)
    list_gcn = detail.get("ListGiayChungNhan") or []
    target_gcn = None
    target_gcn_id = ""
    target_gcn_sph = ""
    target_ver_gcn = 11
    bien_dong_id = safe_int(detail.get("bienDongId") or (detail.get("BienDong") or {}).get("bienDongId"))

    if list_gcn:
        target_gcn, reason_gcn = tim_gcn_trong_don(
            list_gcn=list_gcn,
            so_gcn=so_gcn,
            ten_file=ten_file,
            so_vao_so=so_vao_so,
            gcn_id=gcn_id,
        )
        if target_gcn is None:
            # Đơn có nhiều GCN nhưng không khớp được -> Dừng để tránh gán nhầm
            result["trang_thai"] = "Lỗi"
            result["ghi_chu"] = f"Lỗi xác định GCN: {reason_gcn}"
            return result

        target_gcn_id = str(target_gcn.get("giayChungNhanId") or "").strip()
        target_gcn_sph = str(target_gcn.get("soPhatHanh") or "").strip()
        target_ver_gcn = safe_int(target_gcn.get("versionGiayChungNhan") or target_gcn.get("version"), 11)

    if not so_gcn and target_gcn_sph:
        so_gcn = target_gcn_sph
        result["sogcn"] = so_gcn

    # Xác định mô tả mới chuẩn theo MPLIS
    is_gcn_mode = (loai_ho_so_quet == 1 or thay_moi_gcn or bool(target_gcn_id))
    if is_gcn_mode:
        raw_gcn = target_gcn_sph or so_gcn or Path(ten_file).stem or str(tinh_hinh_dang_ky_id)
        if "giấy chứng nhận" in raw_gcn.lower() or "gcn" in raw_gcn.lower():
            mo_ta_moi = raw_gcn
        else:
            mo_ta_moi = f"Giấy chứng nhận {raw_gcn}"
    elif xa_id and loai_dat:
        mo_ta_moi = f"CHUACOGIAY_{xa_id}_{loai_dat}-DDK"
    else:
        mo_ta_moi = f"Đơn đăng ký {tinh_hinh_dang_ky_id}"

    result["mota_moi"] = mo_ta_moi

    found = tim_hoso_de_cap_nhat(detail, target_gcn=target_gcn, thay_moi_gcn=thay_moi_gcn)
    if found is None:
        result["ghi_chu"] = "Đơn không có hồ sơ quét (ListHoSoQuet rỗng)."
        return result

    hoso = found["hoso"]
    target_file = found["file"]
    hsq_id_raw = hoso.get("hoSoQuetId") or hoso.get("Title")
    result["hosoquetid"] = str(hsq_id_raw) if (hsq_id_raw and str(hsq_id_raw) != "0") else "0 (Tạo mới)"

    if target_file is None and not day_hsq:
        result["trang_thai"] = "Bỏ qua"
        result["ghi_chu"] = "Bỏ qua — HSQ chưa có file để sửa mô tả (cần tick 'Đẩy hồ sơ quét' để upload)."
        return result

    if not thay_moi_gcn and target_file is None:
        result["trang_thai"] = "Bỏ qua"
        result["ghi_chu"] = "Bỏ qua — không còn HSQ nào mang mô tả CHUACOGIAY để sửa."
        return result

    if DRY_RUN:
        result["trang_thai"] = "DRY_RUN"
        result["ghi_chu"] = "DRY_RUN — chưa update thật."
        return result

    la_gcn = (safe_int(loai_ho_so_quet, 2) == 1 or bool(target_gcn_id))

    try:
        if day_hsq:
            file_path = tim_file_pdf_that(folder_upload, ten_file, so_gcn)
            if not file_path:
                result["ghi_chu"] = f"Không tìm thấy file PDF cho '{ten_file or so_gcn}' trong thư mục {folder_upload}"
                return result
            result["tenfile"] = os.path.basename(file_path)
            client.api_update_hosoquet_exist_file(
                file_path=file_path,
                hoso=hoso,
                mo_ta_moi=mo_ta_moi,
                loai_ho_so_quet=loai_ho_so_quet,
                giay_chung_nhan_id=target_gcn_id,
                la_giay_chung_nhan=la_gcn,
                target_file=target_file,
                version_gcn=target_ver_gcn,
                bien_dong_id=bien_dong_id,
                detail=detail,
            )
        else:
            client.api_update_hosoquet_metadata_only(
                hoso=hoso,
                target_file=target_file,
                mo_ta_moi=mo_ta_moi,
                loai_ho_so_quet=loai_ho_so_quet,
                giay_chung_nhan_id=target_gcn_id,
                la_giay_chung_nhan=la_gcn,
                version_gcn=target_ver_gcn,
                bien_dong_id=bien_dong_id,
            )
    except Exception as exc:
        result["ghi_chu"] = f"Update HSQ lỗi: {rut_gon_text(exc)}"
        return result

    if ngay_dang_ky_lan_dau or chu_id:
        try:
            payload = build_payload_update_ttdk(detail, ngay_dang_ky_lan_dau, chu_id=chu_id)
            client.api_update_thong_tin_dang_ky(payload, debug_key=str(tinh_hinh_dang_ky_id))
        except Exception as exc:
            result["ghi_chu"] = f"Update HSQ OK nhưng UpdateThongTinDangKy lỗi: {rut_gon_text(exc)}"
            return result

    gcn_info_str = f" cho GCN {target_gcn_sph} (ID {target_gcn_id})" if target_gcn_sph else ""
    result["trang_thai"] = "Thành công"
    result["ghi_chu"] = f"Đã cập nhật HSQ{gcn_info_str}" + (" + TTĐK/chủ." if (ngay_dang_ky_lan_dau or chu_id) else ".")
    return result


# ============================ TKINTER UI ============================


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Cập nhật hồ sơ quét + thông tin đăng ký MPLIS")
        root.geometry("980x760")
        root.minsize(900, 680)

        self.client = MplisClient(self.log)
        self.running = False
        self.stop_flag = False

        self.var_input_file = tk.StringVar()
        self.var_output_file = tk.StringVar()
        self.var_token = tk.StringVar()
        self.var_cookie = tk.StringVar()
        self.var_ngay_dk = tk.StringVar()
        self.var_chu_id = tk.StringVar()
        self.var_day_hsq = tk.BooleanVar(value=True)
        self.var_thay_moi_gcn = tk.BooleanVar(value=True)
        self.var_folder_upload = tk.StringVar()
        self.var_dry_run = tk.BooleanVar(value=False)
        self.loai_hsq_display_to_id = {
            f"{k} - {v}": k for k, v in LOAI_HO_SO_QUET_OPTIONS.items()
        }
        self.var_loai_hsq = tk.StringVar(value="1 - Giấy chứng nhận")

        self._build_ui()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        login_frame = ttk.LabelFrame(self.root, text="Đăng nhập MPLIS", padding=10)
        login_frame.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(login_frame, text="Username:").grid(row=0, column=0, sticky="w")
        self.ent_user = ttk.Entry(login_frame, width=32)
        self.ent_user.grid(row=0, column=1, sticky="ew", padx=5, pady=3)

        ttk.Label(login_frame, text="Password:").grid(row=0, column=2, sticky="w")
        self.ent_pass = ttk.Entry(login_frame, width=32, show="*")
        self.ent_pass.grid(row=0, column=3, sticky="ew", padx=5, pady=3)

        ttk.Label(login_frame, text="Mã xã (xaId):").grid(row=1, column=0, sticky="w")
        self.ent_xa_id = ttk.Entry(login_frame, width=20)
        self.ent_xa_id.grid(row=1, column=1, sticky="w", padx=5, pady=3)

        ttk.Separator(login_frame, orient="horizontal").grid(
            row=2, column=0, columnspan=4, sticky="ew", pady=8
        )

        ttk.Label(
            login_frame,
            text="Hoặc dán Token/Cookie thủ công (bỏ qua bước mở Chrome):",
            foreground="gray",
        ).grid(row=3, column=0, columnspan=4, sticky="w")

        ttk.Label(login_frame, text="Token:").grid(row=4, column=0, sticky="w")
        self.ent_token = ttk.Entry(login_frame, textvariable=self.var_token)
        self.ent_token.grid(row=4, column=1, columnspan=3, sticky="ew", padx=5, pady=3)

        ttk.Label(login_frame, text="Cookie:").grid(row=5, column=0, sticky="w")
        self.ent_cookie = ttk.Entry(login_frame, textvariable=self.var_cookie)
        self.ent_cookie.grid(row=5, column=1, columnspan=3, sticky="ew", padx=5, pady=3)

        self.btn_dung_token_cookie = ttk.Button(
            login_frame, text="Dùng Token/Cookie này → Lấy session", command=self.dung_token_cookie
        )
        self.btn_dung_token_cookie.grid(row=6, column=1, sticky="w", padx=5, pady=(3, 0))

        login_frame.columnconfigure(1, weight=1)
        login_frame.columnconfigure(3, weight=1)

        file_frame = ttk.LabelFrame(self.root, text="File Excel", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(file_frame, text="File đầu vào:").grid(row=0, column=0, sticky="w")
        self.ent_input_file = ttk.Entry(file_frame, textvariable=self.var_input_file)
        self.ent_input_file.grid(row=0, column=1, sticky="ew", padx=5, pady=3)
        ttk.Button(file_frame, text="Duyệt file...", command=self.chon_file_input).grid(
            row=0, column=2, padx=5, pady=3
        )

        ttk.Label(file_frame, text="File kết quả:").grid(row=1, column=0, sticky="w")
        self.ent_output_file = ttk.Entry(file_frame, textvariable=self.var_output_file)
        self.ent_output_file.grid(row=1, column=1, sticky="ew", padx=5, pady=3)
        ttk.Button(file_frame, text="Chọn nơi lưu...", command=self.chon_file_output).grid(
            row=1, column=2, padx=5, pady=3
        )

        ttk.Label(
            file_frame,
            text="Hỗ trợ Excel: [madon (hoặc tinhhinhdangkyid) + tenfile (hoặc sogcn)] HOẶC [soto + sothua + ...]",
            foreground="blue",
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=0, pady=(4, 0))

        file_frame.columnconfigure(1, weight=1)

        options_frame = ttk.LabelFrame(self.root, text="Tuỳ chọn cập nhật", padding=10)
        options_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(options_frame, text="Ngày ĐK lần đầu:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.ent_ngay_dk = ttk.Entry(options_frame, textvariable=self.var_ngay_dk, width=14)
        self.ent_ngay_dk.grid(row=0, column=1, sticky="w", padx=5, pady=3)
        ttk.Label(options_frame, text="dd/mm/yyyy (để trống nếu không đổi)", foreground="gray").grid(row=0, column=2, sticky="w")

        ttk.Label(options_frame, text="ID thông tin chủ:").grid(row=0, column=3, sticky="w", padx=5, pady=3)
        self.ent_chu_id = ttk.Entry(options_frame, textvariable=self.var_chu_id, width=16)
        self.ent_chu_id.grid(row=0, column=4, sticky="w", padx=5, pady=3)
        ttk.Label(options_frame, text="(để trống nếu không đổi chủ)", foreground="gray").grid(
            row=0, column=5, sticky="w", padx=5, pady=3
        )

        ttk.Label(options_frame, text="Loại hồ sơ quét:").grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.cbo_loai_hsq = ttk.Combobox(
            options_frame,
            textvariable=self.var_loai_hsq,
            values=list(self.loai_hsq_display_to_id.keys()),
            state="readonly",
            width=26,
        )
        self.cbo_loai_hsq.grid(row=1, column=1, columnspan=2, sticky="w", padx=5, pady=3)

        self.chk_day_hsq = ttk.Checkbutton(
            options_frame,
            text="Đẩy hồ sơ quét (upload file PDF)",
            variable=self.var_day_hsq,
            command=self._on_toggle_day_hsq,
        )
        self.chk_day_hsq.grid(row=1, column=3, columnspan=2, sticky="w", padx=5, pady=3)

        self.chk_dry_run = ttk.Checkbutton(
            options_frame, text="DRY RUN (chỉ kiểm tra, không update)", variable=self.var_dry_run
        )
        self.chk_dry_run.grid(row=1, column=5, sticky="w", padx=5, pady=3)

        ttk.Label(options_frame, text="Folder PDF:").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.ent_folder_upload = ttk.Entry(options_frame, textvariable=self.var_folder_upload, width=60)
        self.ent_folder_upload.grid(row=2, column=1, columnspan=4, sticky="ew", padx=5, pady=3)
        self.btn_browse_folder_upload = ttk.Button(
            options_frame, text="Duyệt...", command=self.chon_folder_upload
        )
        self.btn_browse_folder_upload.grid(row=2, column=5, sticky="w", padx=5, pady=3)

        self.chk_thay_moi_gcn = ttk.Checkbutton(
            options_frame,
            text="Thay thế file HSQ theo số GCN / tên file (không lọc CHUACOGIAY)",
            variable=self.var_thay_moi_gcn,
        )
        self.chk_thay_moi_gcn.grid(row=3, column=0, columnspan=6, sticky="w", padx=5, pady=3)

        ttk.Label(
            options_frame,
            text="Hỗ trợ: Xử lý trực tiếp theo ID đơn (tinhHinhDangKyId / madon) hoặc tra theo Tờ / Thửa.",
            foreground="blue",
        ).grid(row=4, column=0, columnspan=6, sticky="w", padx=5, pady=(4, 0))

        for col_index in (1, 4):
            options_frame.columnconfigure(col_index, weight=1)

        button_frame = ttk.Frame(self.root, padding=(10, 5))
        button_frame.pack(fill="x")

        self.btn_login = ttk.Button(button_frame, text="1. Mở Chrome đăng nhập", command=self.mo_chrome)
        self.btn_login.pack(side="left", padx=5)

        self.btn_session = ttk.Button(
            button_frame,
            text="2. Đã đăng nhập xong → Lấy session",
            command=self.lay_session,
            state="disabled",
        )
        self.btn_session.pack(side="left", padx=5)

        self.btn_run = ttk.Button(
            button_frame, text="3. Bắt đầu xử lý Excel", command=self.chay, state="disabled"
        )
        self.btn_run.pack(side="left", padx=5)

        self.btn_stop = ttk.Button(button_frame, text="Dừng", command=self.dung, state="disabled")
        self.btn_stop.pack(side="left", padx=5)

        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(5, 0))

        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="x", padx=10, pady=(3, 0))

        self.lbl_status = ttk.Label(status_frame, text="Chưa chạy")
        self.lbl_status.pack(side="left")

        self.lbl_count = ttk.Label(status_frame, text="", foreground="blue")
        self.lbl_count.pack(side="right")

        log_frame = ttk.LabelFrame(self.root, text="Nhật ký xử lý", padding=5)
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)

        self.txt_log = tk.Text(log_frame, wrap="word", height=25)
        self.txt_log.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        scrollbar.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=scrollbar.set)

    # ---------- UI helpers ----------
    def log(self, message: str) -> None:
        def append() -> None:
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.txt_log.insert("end", f"{timestamp}  {message}\n")
            self.txt_log.see("end")

        self.root.after(0, append)

    def set_status(self, message: str) -> None:
        self.root.after(0, lambda: self.lbl_status.config(text=message))

    def set_count(self, processed: int, total: int) -> None:
        self.root.after(0, lambda: self.lbl_count.config(text=f"Dòng: {processed} / {total}"))

    def set_progress(self, value: int, maximum: int) -> None:
        def update() -> None:
            self.progress.config(maximum=max(maximum, 1), value=value)

        self.root.after(0, update)

    def set_controls_state(self, state: str) -> None:
        widgets = [
            self.ent_ngay_dk, self.ent_chu_id, self.cbo_loai_hsq,
            self.chk_day_hsq, self.chk_dry_run, self.ent_folder_upload,
            self.btn_browse_folder_upload,
        ]
        for widget in widgets:
            widget.config(state=state)

        if state == "normal":
            self.cbo_loai_hsq.config(state="readonly")
            self._on_toggle_day_hsq()

    def _on_toggle_day_hsq(self) -> None:
        state = "normal" if self.var_day_hsq.get() else "disabled"
        self.ent_folder_upload.config(state=state)
        self.btn_browse_folder_upload.config(state=state)

    def chon_folder_upload(self) -> None:
        folder = filedialog.askdirectory(title="Chọn folder chứa PDF")
        if folder:
            self.var_folder_upload.set(folder)

    # ---------- file actions ----------
    def chon_file_input(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel có cột soto, sothua, loaidat, tenfile, sogcn",
            filetypes=[
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro-Enabled", "*.xlsm"),
                ("Tất cả file", "*.*"),
            ],
        )

        if not file_path:
            return

        self.var_input_file.set(file_path)

        input_path = Path(file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = input_path.with_name(f"ket_qua_cap_nhat_don_hsq_{timestamp}.xlsx")
        self.var_output_file.set(str(output_path))

    def chon_file_output(self) -> None:
        initial_name = Path(self.var_output_file.get() or "ket_qua_cap_nhat_don_hsq.xlsx").name

        file_path = filedialog.asksaveasfilename(
            title="Chọn nơi lưu file kết quả",
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )

        if file_path:
            self.var_output_file.set(file_path)

    # ---------- login actions ----------
    def mo_chrome(self) -> None:
        username = self.ent_user.get().strip()
        password = self.ent_pass.get()

        if not username or not password:
            messagebox.showwarning("Thiếu thông tin", "Nhập username và password trước.")
            return

        self.btn_login.config(state="disabled")

        def work() -> None:
            try:
                self.log("Đang mở Chrome...")
                self.client.open_browser_and_fill_login(username, password)
                self.log(
                    "Chrome đã mở. Hoàn tất đăng nhập/OTP nếu có, "
                    "sau đó bấm nút 2 để lấy session."
                )
                self.root.after(0, lambda: self.btn_session.config(state="normal"))
            except Exception as exc:
                self.log(f"❌ Lỗi mở Chrome: {exc}")
                self.root.after(0, lambda: self.btn_login.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    def lay_session(self) -> None:
        self.btn_session.config(state="disabled")

        def work() -> None:
            try:
                self.client.build_session_from_browser()
                self.client.close_browser()
                self.log("✅ Đã đóng Chrome sau khi lấy session.")
                self.root.after(0, lambda: self.btn_run.config(state="normal"))
            except Exception as exc:
                self.log(f"❌ {exc}")
                self.root.after(0, lambda: self.btn_session.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    def dung_token_cookie(self) -> None:
        token = self.var_token.get().strip()
        cookie = self.var_cookie.get().strip()

        if not token or not cookie:
            messagebox.showwarning("Thiếu thông tin", "Nhập cả Token và Cookie trước.")
            return

        try:
            self.client.build_session_from_manual(token, cookie)
        except Exception as exc:
            messagebox.showerror("Lỗi", str(exc))
            return

        self.btn_run.config(state="normal")
        messagebox.showinfo("Thành công", "Đã tạo session từ Token/Cookie thủ công. Có thể bắt đầu chạy.")

    # ---------- processing actions ----------
    def dung(self) -> None:
        self.stop_flag = True
        self.log("⏸ Đã yêu cầu dừng. Sẽ dừng sau request hiện tại và lưu kết quả đã có.")

    def chay(self) -> None:
        if self.running:
            return

        xa_id = self.ent_xa_id.get().strip()
        input_file = self.var_input_file.get().strip()
        output_file = self.var_output_file.get().strip()
        ngay_dk = self.var_ngay_dk.get().strip()
        chu_id = self.var_chu_id.get().strip() or None
        day_hsq = self.var_day_hsq.get()
        thay_moi_gcn = self.var_thay_moi_gcn.get()
        folder_upload = self.var_folder_upload.get().strip()
        loai_ho_so_quet = self.loai_hsq_display_to_id.get(self.var_loai_hsq.get(), 1)
        dry_run = self.var_dry_run.get()

        if not input_file:
            messagebox.showwarning("Thiếu file", "Chọn file Excel đầu vào.")
            return
        if not os.path.isfile(input_file):
            messagebox.showwarning("Không tìm thấy file", input_file)
            return
        if not output_file:
            messagebox.showwarning("Thiếu file", "Chọn đường dẫn file kết quả.")
            return
        if Path(input_file).resolve() == Path(output_file).resolve():
            messagebox.showwarning(
                "Sai đường dẫn", "File kết quả không được trùng với file Excel đầu vào."
            )
            return

        if ngay_dk:
            try:
                ddmmyyyy_to_iso_utc_start_of_day_vn(ngay_dk)
            except Exception:
                messagebox.showerror(
                    "Sai định dạng", "Ngày đăng ký lần đầu phải là dd/mm/yyyy. Ví dụ: 23/04/2026"
                )
                return

        if day_hsq and not os.path.isdir(folder_upload):
            messagebox.showwarning("Sai đường dẫn", "Folder PDF không tồn tại (cần khi có đẩy HSQ).")
            return

        try:
            rows = doc_danh_sach(input_file)
        except Exception as exc:
            messagebox.showerror("Lỗi đọc Excel", str(exc))
            return

        is_mode_id_don = any(bool(item.get("id_don")) for item in rows)

        if not is_mode_id_don:
            if not xa_id:
                messagebox.showwarning("Thiếu thông tin", "Nhập mã xã (xaId) khi tra cứu theo Tờ / Thửa.")
                return
            if not xa_id.isdigit():
                messagebox.showwarning("Sai mã xã", "Mã xã phải là số nguyên.")
                return
        else:
            if xa_id and not xa_id.isdigit():
                messagebox.showwarning("Sai mã xã", "Mã xã nếu nhập phải là số nguyên.")
                return

        # Gom nhóm: phân biệt theo ID đơn / tờ thửa VÀ số GCN / tên file để tránh gộp nhầm các GCN khác nhau trong cùng 1 đơn
        grouped: dict[str, list[dict[str, Any]]] = {}
        for item in rows:
            key = make_group_key(item)
            grouped.setdefault(key, []).append(item)
        groups = list(grouped.items())

        global DRY_RUN
        DRY_RUN = dry_run

        che_do_str = "THEO MÃ / ID ĐƠN (tinhHinhDangKyId)" if is_mode_id_don else "THEO TỜ / THỬA"
        confirm_text = (
            f"Chế độ: {che_do_str}\n"
            f"File có {len(rows)} dòng | gom còn {len(groups)} nhóm duy nhất.\n\n"
            f"Đẩy file PDF: {'CÓ' if day_hsq else 'KHÔNG (chỉ sửa mô tả/loại HSQ)'}\n"
            f"Thay thế HSQ: {'Số GCN / file PDF (không lọc CHUACOGIAY)' if thay_moi_gcn else 'CHUACOGIAY'}\n"
            f"Loại HSQ: {loai_ho_so_quet} ({LOAI_HO_SO_QUET_OPTIONS.get(loai_ho_so_quet, '')})\n"
            f"Ngày ĐK lần đầu: {ngay_dk or '(không đổi)'}\n"
            f"Đổi chủ: {chu_id or '(không đổi)'}\n"
            f"DRY RUN: {'CÓ — chỉ kiểm tra' if dry_run else 'KHÔNG — update thật'}\n\n"
            "Tiếp tục?"
        )
        if not messagebox.askyesno("Xác nhận cập nhật", confirm_text):
            return

        self.running = True
        self.stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.set_controls_state("disabled")
        self.progress.config(value=0, maximum=max(len(groups), 1))
        self.set_count(0, len(groups))

        threading.Thread(
            target=self._run_process,
            args=(xa_id, groups, output_file, day_hsq, chu_id, loai_ho_so_quet, folder_upload, ngay_dk, thay_moi_gcn),
            daemon=True,
        ).start()

    def _xu_ly_1_row(
        self,
        xa_id: str,
        row: dict[str, Any],
        day_hsq: bool,
        chu_id: str | None,
        loai_ho_so_quet: int,
        folder_upload: str,
        ngay_dang_ky_lan_dau: str,
        thay_moi_gcn: bool = True,
    ) -> list[dict[str, Any]]:
        """Xử lý cho 1 dòng Excel: nếu có ID đơn thì xử lý trực tiếp; nếu không thì tra cứu theo tờ/thửa."""
        id_don = row.get("id_don")
        so_to = row.get("soto", "")
        so_thua = row.get("sothua", "")
        loai_dat = row.get("loaidat", "")
        ten_file = row.get("tenfile", "")
        so_gcn = row.get("sogcn", "")
        so_vao_so = row.get("sovaoso", "")
        gcn_id = row.get("gcn_id", "")

        # Nhánh 1: Có sẵn ID đơn (tinhHinhDangKyId / madon)
        if id_don:
            try:
                tinh_hinh_id_int = int(id_don)
            except (TypeError, ValueError):
                ext_id, ext_gcn = tach_thong_tin_tu_chuoi_madon(id_don)
                if ext_id:
                    tinh_hinh_id_int = int(ext_id)
                    if not so_gcn and ext_gcn:
                        so_gcn = ext_gcn
                    if not ten_file and so_gcn:
                        ten_file = f"{so_gcn}.pdf"
                else:
                    return [
                        {
                            "soto": so_to, "sothua": so_thua, "loaidat": loai_dat,
                            "tenfile": ten_file, "sogcn": so_gcn,
                            "tinhhinhdangkyid": id_don,
                            "trang_thai": "Lỗi", "ghi_chu": f"ID đơn '{id_don}' không phải số nguyên hợp lệ.",
                        }
                    ]

            self.log(f"   → Xử lý trực tiếp theo ID đơn {tinh_hinh_id_int}")
            row_result = xu_ly_mot_don(
                client=self.client,
                tinh_hinh_dang_ky_id=tinh_hinh_id_int,
                so_to=so_to,
                so_thua=so_thua,
                loai_dat=loai_dat,
                ten_file=ten_file,
                so_gcn=so_gcn,
                xa_id=xa_id,
                folder_upload=folder_upload,
                ngay_dang_ky_lan_dau=ngay_dang_ky_lan_dau,
                day_hsq=day_hsq,
                chu_id=chu_id,
                loai_ho_so_quet=loai_ho_so_quet,
                thay_moi_gcn=thay_moi_gcn,
                so_vao_so=so_vao_so,
                gcn_id=gcn_id,
            )
            self.log(f"      {row_result.get('trang_thai')}: {row_result.get('ghi_chu')}")
            time.sleep(REQUEST_DELAY_SECONDS)
            return [row_result]

        # Nhánh 2: Tra cứu theo Tờ / Thửa
        if not so_to or not so_thua:
            return [
                {
                    "soto": so_to, "sothua": so_thua, "loaidat": loai_dat,
                    "tenfile": ten_file, "sogcn": so_gcn,
                    "trang_thai": "Bỏ qua", "ghi_chu": "Thiếu số tờ hoặc số thửa (và không có ID đơn).",
                }
            ]

        try:
            registrations = self.client.tim_tat_ca_don(xa_id=xa_id, so_to=so_to, so_thua=so_thua)
        except Exception as exc:
            self.log(f"   ❌ Lỗi tra cứu: {exc}")
            return [
                {
                    "soto": so_to, "sothua": so_thua, "loaidat": loai_dat,
                    "tenfile": ten_file, "sogcn": so_gcn,
                    "trang_thai": "Lỗi", "ghi_chu": f"Lỗi tra cứu: {rut_gon_text(exc)}",
                }
            ]

        if not registrations:
            self.log("   Không tìm thấy đơn đăng ký.")
            return [
                {
                    "soto": so_to, "sothua": so_thua, "loaidat": loai_dat,
                    "tenfile": ten_file, "sogcn": so_gcn,
                    "trang_thai": "Lỗi", "ghi_chu": "Không tìm thấy đơn đăng ký.",
                }
            ]

        if len(registrations) > 1:
            self.log(f"   ⚠ Tìm thấy {len(registrations)} đơn — sẽ xử lý TẤT CẢ.")
        else:
            self.log("   Tìm thấy 1 đơn đăng ký.")

        results: list[dict[str, Any]] = []

        for registration in registrations:
            if self.stop_flag:
                break

            tinh_hinh_id = registration.get("tinhHinhDangKyId")

            try:
                tinh_hinh_id_int = int(tinh_hinh_id)
            except (TypeError, ValueError):
                results.append(
                    {
                        "soto": so_to, "sothua": so_thua, "loaidat": loai_dat,
                        "tenfile": ten_file, "sogcn": so_gcn,
                        "tinhhinhdangkyid": tinh_hinh_id or "",
                        "trang_thai": "Lỗi", "ghi_chu": "tinhHinhDangKyId không hợp lệ.",
                    }
                )
                continue

            self.log(f"   → Xử lý ID {tinh_hinh_id_int}")

            row_result = xu_ly_mot_don(
                client=self.client,
                tinh_hinh_dang_ky_id=tinh_hinh_id_int,
                so_to=so_to,
                so_thua=so_thua,
                loai_dat=loai_dat,
                ten_file=ten_file,
                so_gcn=so_gcn,
                xa_id=xa_id,
                folder_upload=folder_upload,
                ngay_dang_ky_lan_dau=ngay_dang_ky_lan_dau,
                day_hsq=day_hsq,
                chu_id=chu_id,
                loai_ho_so_quet=loai_ho_so_quet,
                thay_moi_gcn=thay_moi_gcn,
                so_vao_so=so_vao_so,
                gcn_id=gcn_id,
            )
            results.append(row_result)
            self.log(f"      {row_result.get('trang_thai')}: {row_result.get('ghi_chu')}")
            time.sleep(REQUEST_DELAY_SECONDS)

        return results

    @staticmethod
    def _make_copy_result(
        master_result: dict[str, Any], sibling_row: dict[str, Any], master_excel_row: Any, group_key: str
    ) -> dict[str, Any]:
        copied = dict(master_result)
        copied["soto"] = sibling_row.get("soto") or master_result.get("soto")
        copied["sothua"] = sibling_row.get("sothua") or master_result.get("sothua")
        copied["loaidat"] = sibling_row.get("loaidat") or master_result.get("loaidat")
        copied["tenfile"] = sibling_row.get("tenfile") or master_result.get("tenfile")
        copied["sogcn"] = sibling_row.get("sogcn") or master_result.get("sogcn")
        copied["ghi_chu"] = (
            f"Cùng nhóm '{group_key}' với dòng {master_excel_row}; "
            "không gọi API lần nữa. Kết quả dùng chung từ dòng đại diện. "
            + str(master_result.get("ghi_chu") or "")
        )
        return copied

    def _run_process(
        self,
        xa_id: str,
        groups: list[tuple[str, list[dict[str, Any]]]],
        output_file: str,
        day_hsq: bool,
        chu_id: str | None,
        loai_ho_so_quet: int,
        folder_upload: str,
        ngay_dang_ky_lan_dau: str,
        thay_moi_gcn: bool = True,
    ) -> None:
        writer = ExcelResultWriter(output_file)
        processed = 0
        total_rows_in = sum(len(items) for _, items in groups)
        total_results = 0
        thanh_cong = 0
        loi = 0
        bo_qua = 0

        try:
            self.log("=" * 70)
            self.log("BẮT ĐẦU - Cập nhật hồ sơ quét + thông tin đăng ký")
            if xa_id:
                self.log(f"Mã xã: {xa_id}")
            self.log(
                f"Đẩy HSQ: {day_hsq} | Thay mới file GCN: {thay_moi_gcn} | Loại HSQ: {loai_ho_so_quet} "
                f"({LOAI_HO_SO_QUET_OPTIONS.get(loai_ho_so_quet, '')})"
            )
            self.log(f"ID chủ: {chu_id or '(không đổi)'} | DRY_RUN: {DRY_RUN}")
            self.log(f"Số dòng đầu vào: {total_rows_in} | Gom còn {len(groups)} nhóm duy nhất.")
            self.log(f"File kết quả: {output_file}")

            for group_index, (group_key, items_in_group) in enumerate(groups, start=1):
                if self.stop_flag:
                    self.log("⏹ Đã dừng theo yêu cầu.")
                    break

                representative = items_in_group[0]
                excel_row = representative["excel_row"]
                id_don_repr = representative.get("id_don")

                if id_don_repr:
                    self.set_status(f"Đang xử lý ID đơn {id_don_repr}")
                    self.log(
                        f"--- [{group_index}/{len(groups)}] "
                        f"Excel dòng {excel_row}: ID đơn {id_don_repr} | File/GCN: {representative.get('tenfile') or representative.get('sogcn')} ---"
                    )
                else:
                    self.set_status(
                        f"Đang xử lý tờ {representative['soto'] or '?'} - thửa {representative['sothua'] or '?'}"
                    )
                    self.log(
                        f"--- [{group_index}/{len(groups)}] "
                        f"Excel dòng {excel_row}: tờ {representative['soto']}, thửa {representative['sothua']} ---"
                    )

                master_results = self._xu_ly_1_row(
                    xa_id=xa_id,
                    row=representative,
                    day_hsq=day_hsq,
                    chu_id=chu_id,
                    loai_ho_so_quet=loai_ho_so_quet,
                    folder_upload=folder_upload,
                    ngay_dang_ky_lan_dau=ngay_dang_ky_lan_dau,
                    thay_moi_gcn=thay_moi_gcn,
                )

                for r in master_results:
                    writer.append_result(r)
                    total_results += 1
                    st = r.get("trang_thai")
                    if st == "Thành công":
                        thanh_cong += 1
                    elif st == "Bỏ qua":
                        bo_qua += 1
                    else:
                        loi += 1

                if len(items_in_group) > 1:
                    self.log(
                        f"🔁 Nhóm soGCN '{group_key}' có {len(items_in_group)} dòng; "
                        f"đã xử lý 1 lần ở dòng {excel_row}, các dòng còn lại dùng chung kết quả."
                    )
                    for sibling in items_in_group[1:]:
                        for r in master_results:
                            copied = self._make_copy_result(r, sibling, excel_row, group_key)
                            writer.append_result(copied)
                            total_results += 1
                            st = copied.get("trang_thai")
                            if st == "Thành công":
                                thanh_cong += 1
                            elif st == "Bỏ qua":
                                bo_qua += 1
                            else:
                                loi += 1

                processed += 1
                self.set_progress(processed, len(groups))
                self.set_count(processed, len(groups))

                if processed % SAVE_EVERY_ROWS == 0:
                    writer.save()
                    self.log(f"💾 Đã tự lưu kết quả sau {processed} nhóm.")

            writer.save()
            self.log("💾 Đã lưu file kết quả cuối cùng.")

            self.log("=" * 70)
            self.log(f"Đã xử lý nhóm: {processed}/{len(groups)} (tổng {total_rows_in} dòng Excel)")
            self.log(f"Tổng dòng kết quả: {total_results}")
            self.log(f"Thành công: {thanh_cong}")
            self.log(f"Lỗi: {loi}")
            self.log(f"Bỏ qua: {bo_qua}")
            self.log(f"File kết quả: {output_file}")

            if self.stop_flag:
                self.set_status(f"Đã dừng: xử lý {processed}/{len(groups)} nhóm, đã lưu kết quả")
            else:
                self.set_status(f"Hoàn tất: {processed} nhóm, {thanh_cong} thành công")

        except PermissionError:
            self.log(
                "❌ Không lưu được file kết quả. Có thể file đang được mở trong Excel. "
                "Hãy đóng file rồi chạy lại."
            )
            self.set_status("Lỗi lưu file kết quả")

        except Exception as exc:
            self.log(f"❌ Lỗi chương trình: {exc}")
            self.set_status("Chương trình gặp lỗi")

            try:
                writer.save()
                self.log("💾 Đã cố gắng lưu phần kết quả đang có.")
            except Exception as save_exc:
                self.log(f"❌ Không lưu được kết quả dở dang: {save_exc}")

        finally:
            writer.close()
            self.running = False

            def reset_buttons() -> None:
                self.btn_run.config(state="normal")
                self.btn_stop.config(state="disabled")
                self.set_controls_state("normal")

            self.root.after(0, reset_buttons)

    def on_close(self) -> None:
        if self.running:
            confirm = messagebox.askyesno(
                "Đang xử lý",
                "Chương trình đang chạy. Thoát ngay có thể làm mất tối đa 4 dòng chưa kịp lưu. Thoát?",
            )
            if not confirm:
                return

        self.stop_flag = True
        self.client.close_browser()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
