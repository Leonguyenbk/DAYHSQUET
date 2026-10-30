# -*- coding: utf-8 -*-
"""
Đơn đã đạt nhóm 1 (làm sạch dữ liệu MPLIS) hay chưa — dùng chung cho tool cập nhật thông tin thiếu, tool tải &
tách HSQ và các tool đẩy HSQ.

Tra màn làm sạch (LamSachDuLieuAjax/GetThongKePhanLoaiThuaDatChiTiet) theo tờ / thửa + xã của từng thửa trong
đơn (phanLoai -1), giữ bản ghi thuộc đơn. Đơn đạt nhóm 1 = có bản ghi và MỌI bản ghi thongTinDangKyDapUngNhom1
= true (đơn nhiều thửa: còn thửa chưa đáp ứng thì vẫn xử lý).

Không import AI / PDF: các tool đẩy HSQ (bản Flet) dùng được mà không phải đóng gói thêm.
"""

from __future__ import annotations

from collections import deque
from typing import Any

import cap_nhat_don_hsq as core
import tai_hsq_don as thd

log = thd.log


def payload_tra_cuu_truc_tiep_theo_id(xa_id: int, id_don: int, tinh_id: int = 66) -> dict[str, str]:
    """Payload tương ứng màn tra cứu MPLIS khi chỉ nhập tinhHinhDangKyId."""
    import bao_cao_thong_tin_thieu_mplis as bc

    payload = bc.build_list_payload(str(xa_id), start=0, length=10, tinh_id=tinh_id)
    payload.update({
        "traCuu[soThuTuThua]": "",
        "traCuu[soHieuToBanDo]": "",
        "traCuu[tinhHinhDangKyId]": str(id_don),
        "traCuu[phanLoai]": "-1",
    })
    return payload


def payload_tra_cuu_theo_to_thua(xa_id: int, so_to: Any, so_thua: Any, tinh_id: int = 66) -> dict[str, str]:
    """Payload màn tra cứu làm sạch MPLIS khi nhập tờ / thửa + xã (tinhHinhDangKyId để trống, phanLoai -1)."""
    import bao_cao_thong_tin_thieu_mplis as bc

    payload = bc.build_list_payload(str(xa_id), start=0, length=10, tinh_id=tinh_id)
    payload.update({
        "traCuu[soThuTuThua]": str(so_thua),
        "traCuu[soHieuToBanDo]": str(so_to),
        "traCuu[tinhHinhDangKyId]": "",
        "traCuu[phanLoai]": "-1",
    })
    return payload


def to_thua_cua_don(detail: dict) -> list[tuple[int, str, str]]:
    """Các thửa của đơn → [(xaId, số tờ, số thửa)] không trùng (xaId thửa, không có thì xaId đơn)."""
    xa_don = core.safe_int((detail.get("TinhHinhDangKy") or {}).get("xaId"))
    ket_qua: list[tuple[int, str, str]] = []
    hang: deque = deque([detail.get("TaiSan") or detail])
    while hang:
        o = hang.popleft()
        if isinstance(o, dict):
            to, thua = o.get("soHieuToBanDo"), o.get("soThuTuThua")
            if "thuaDatId" in o and to not in (None, "") and thua not in (None, ""):
                x = (core.safe_int(o.get("xaId")) or xa_don, str(to).strip(), str(thua).strip())
                if x not in ket_qua:
                    ket_qua.append(x)
            hang.extend(v for v in o.values() if isinstance(v, (dict, list)))
        elif isinstance(o, list):
            hang.extend(o)
    return ket_qua


def _cua_don(ban_ghi: dict, id_don: int) -> bool:
    ids = {core.safe_int(x) for x in (ban_ghi.get("tinhHinhDangKyIds") or [])}
    return core.safe_int(ban_ghi.get("tinhHinhDangKyId")) == id_don or id_don in ids


def lay_ban_ghi_lam_sach(client, id_don: int, detail: dict, theo_id: bool = False) -> list[dict]:
    """
    Bản ghi làm sạch MPLIS của 1 đơn (LamSachDuLieuAjax/GetThongKePhanLoaiThuaDatChiTiet): tra theo tờ / thửa
    + xã của từng thửa trong đơn (phanLoai -1), giữ bản ghi có tinhHinhDangKyId / tinhHinhDangKyIds là đơn này.
    Mỗi bản ghi có id, chiTietThongTinThieu, thongTinDangKyDapUngNhom1, thongTinDangKyChuaDapUngNhom1.
    Đơn không có tờ / thửa trong chi tiết (hoặc theo_id) → tra theo tinhHinhDangKyId. Hết phiên → core.HetPhien.
    """
    import bao_cao_thong_tin_thieu_mplis as bc

    th = detail.get("TinhHinhDangKy") or {}
    xa_id = core.safe_int(th.get("xaId"))
    tinh_id = core.safe_int(th.get("tinhId")) or bc.DEFAULT_TINH_ID
    headers = {"Accept": "application/json, text/javascript, */*; q=0.01",
               "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
               "X-Requested-With": "XMLHttpRequest", "Origin": bc.BASE_URL, "Referer": bc.REFERER_URL}

    def tra(payload: dict) -> list[dict]:
        resp = client._post_doc(bc.URL_DANH_SACH, data=payload, headers=headers)
        data = client._response_json(resp, "GetThongKePhanLoaiThuaDatChiTiet")
        return [r for r in (data.get("data") or []) if isinstance(r, dict)]

    if theo_id:
        if not xa_id:
            raise RuntimeError("Chi tiết đơn không có xaId — không tra được thông tin thiếu trên MPLIS")
        return [r for r in tra(payload_tra_cuu_truc_tiep_theo_id(xa_id, id_don, tinh_id))
                if core.safe_int(r.get("tinhHinhDangKyId"), id_don) == id_don]

    cac_thua = [t for t in to_thua_cua_don(detail) if t[0]]
    if not cac_thua:
        if not xa_id:
            raise RuntimeError("Chi tiết đơn không có tờ/thửa và xaId — không tra được thông tin thiếu trên MPLIS")
        payload = bc.build_list_payload(str(xa_id), start=0, length=50, tinh_id=tinh_id)
        payload["traCuu[tinhHinhDangKyId]"] = str(id_don)
        return [r for r in tra(payload) if core.safe_int(r.get("tinhHinhDangKyId"), id_don) == id_don]
    ket_qua: dict[str, dict] = {}
    for xa, so_to, so_thua in cac_thua:
        for r in tra(payload_tra_cuu_theo_to_thua(xa, so_to, so_thua, tinh_id)):
            if _cua_don(r, id_don):
                ket_qua.setdefault(str(r.get("id") or len(ket_qua)), r)
    return list(ket_qua.values())


def dap_ung_nhom_1(ban_ghi: dict) -> bool:
    value = ban_ghi.get("thongTinDangKyDapUngNhom1")
    return value is True or str(value).strip().lower() in {"true", "1", "x", "yes"}


def loc_chua_dap_ung(ban_ghi: list[dict]) -> tuple[bool, list[dict]]:
    """
    → (đơn đã đáp ứng nhóm 1 → bỏ qua đơn, các bản ghi CHƯA đáp ứng). Đơn đáp ứng = có bản ghi và mọi bản ghi
    thongTinDangKyDapUngNhom1 = true (đơn nhiều thửa: thửa đã đáp ứng thì bỏ, thửa chưa đáp ứng vẫn xử lý).
    """
    con = [bg for bg in ban_ghi if not dap_ung_nhom_1(bg)]
    return bool(ban_ghi) and not con, con


LY_DO_DA_DAT = "MPLIS: đơn đã đạt nhóm 1 (thongTinDangKyDapUngNhom1 = true) — bỏ qua đơn"


def kiem_tra(client, id_don: int, detail: dict) -> tuple[bool, list[dict] | None]:
    """
    → (đơn đã đạt nhóm 1, bản ghi làm sạch). Tra lỗi (không phải hết phiên) → (False, None): coi như chưa đạt để
    vẫn xử lý đơn, chỉ ghi cảnh báo. Hết phiên → core.HetPhien.
    """
    try:
        ban_ghi = lay_ban_ghi_lam_sach(client, id_don, detail)
    except core.HetPhien:
        raise
    except Exception as exc:
        log(f"   ⚠ Không kiểm tra được nhóm 1 trên MPLIS (vẫn xử lý đơn): {core.rut_gon_text(exc, 200)}")
        return False, None
    da_dat = loc_chua_dap_ung(ban_ghi)[0]
    if da_dat:
        log(f"   {LY_DO_DA_DAT}")
    return da_dat, ban_ghi


def da_dat(client, id_don: int, detail: dict) -> bool:
    """Đơn đã đạt nhóm 1 → True (đã ghi log). Dùng ở đầu xu_ly_don của các tool đẩy HSQ."""
    return kiem_tra(client, id_don, detail)[0]
