# -*- coding: utf-8 -*-
"""
Bổ sung thành phần HSQ cho các GCN còn thiếu, dùng chung 1 file GCN đã có.

Trường hợp: đơn có nhiều GCN (vd 3) nhưng HSQ chỉ có 1 file GCN (file đó chứa cả 3 GCN).
Với mỗi đơn:
  1. Lấy chi tiết đơn → GCN hiện hành (ListGiayChungNhan) + thành phần GCN trong HSQ.
  2. GCN thiếu = GCN chưa có thành phần nào (loại GCN) trỏ tới giayChungNhanId của nó.
  3. File nguồn = thành phần GCN duy nhất đang có (nhiều hơn 1 → bỏ qua vì không rõ file nào).
     Tải file nguồn về <thư mục>/<madon>/ (FileHandler.ashx?DocId=<nodeId>).
  4. Gửi 1 request cho HSQ chứa file nguồn, kèm ĐỦ thành phần (xem day_hsq_ky_so):
       - thành phần hiện có: giữ nguyên (nodeId cũ, files=null)
       - mỗi GCN thiếu: thành phần mới = metadata thành phần nguồn, đổi giayChungNhanId/version/moTa,
         nodeId="" + fileHoSoQuet_i = file nguồn.
  5. Xác nhận: số thành phần = cũ + số GCN bổ sung, thành phần cũ còn nguyên nodeId,
     mọi GCN đều đã có thành phần. Số thành phần giảm → DỪNG TOÀN BỘ.
"""

from __future__ import annotations

import contextlib
import json
import os
import time
import uuid
from typing import Any, Callable

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import tai_hsq_don as thd

log = thd.log

EXCEL_HEADERS = [
    "stt", "madon", "so_gcn", "gcn_da_co", "gcn_bo_sung", "hosoquetid", "file_nguon",
    "so_thanh_phan_truoc", "so_thanh_phan_sau", "trang_thai", "ghi_chu",
]
EXCEL_WIDTHS = {"stt": 6, "madon": 12, "so_gcn": 8, "gcn_da_co": 28, "gcn_bo_sung": 28, "hosoquetid": 11,
                "file_nguon": 50, "so_thanh_phan_truoc": 10, "so_thanh_phan_sau": 10, "trang_thai": 16,
                "ghi_chu": 60}


def la_thanh_phan_gcn(node: dict[str, Any]) -> bool:
    return bool(node.get("giayChungNhanId")) and (
        bool(node.get("laGiayChungNhan")) or core.safe_int(node.get("loaiHoSoQuet"), -1) == 1
    )


def gcn_hien_hanh(detail: dict[str, Any]) -> list[dict[str, Any]]:
    """GCN của đơn, bỏ bản không còn hiện hành (isLastest = False)."""
    return [g for g in detail.get("ListGiayChungNhan") or []
            if isinstance(g, dict) and g.get("giayChungNhanId") and g.get("isLastest") is not False]


def mo_ta_gcn(gcn: dict[str, Any]) -> str:
    return f"Giấy chứng nhận {str(gcn.get('soPhatHanh') or '').strip()}".strip()


def info_thanh_phan_moi(nguon: dict[str, Any], gcn: dict[str, Any]) -> dict[str, Any]:
    """
    Thành phần mới cho 1 GCN theo mẫu "tạo mới" (vi_du_payload_cap_nhat_hosoquet.json, TRƯỜNG HỢP 1 —
    đã chạy thật khi khôi phục đơn 12830794). Không có nodeId/thanhPhanHoSoQuet*: gửi dạng thành phần
    cũ (nodeId="", tp=0) thì server trả success nhưng bỏ qua, không tạo thành phần.
    """
    return {
        "loaiHoSoQuet": 1,
        "laGiayToVeNguonGoc": True,
        "giayChungNhanId": str(core.safe_int(gcn.get("giayChungNhanId"))),
        "moTa": mo_ta_gcn(gcn),
        "tenGiayTo": "",
        "trichYeu": "",
        "versionGiayChungNhan": core.safe_int(gcn.get("version"), 11),
        "laGiayChungNhan": True,
        "daKySo": bool(nguon.get("daKySo")),
        "__id": str(uuid.uuid4()),
        "files": None,
    }


def mo_ta_thanh_phan(n: dict[str, Any]) -> str:
    return (f"{n.get('moTa')} | loai={n.get('loaiHoSoQuet')} laGCN={n.get('laGiayChungNhan')}"
            f" gcnId={n.get('giayChungNhanId')} ver={n.get('versionGiayChungNhan')}")


def log_trang_thai_sau(detail: dict[str, Any], hsq_id: int, thieu: list[dict], thu_muc: str) -> None:
    """Chẩn đoán khi xác nhận lệch: in mọi thành phần + GCN cần có, lưu JSON trạng thái."""
    log(f"   GCN cần bổ sung: " + ", ".join(
        f"{g.get('soPhatHanh')} (gcnId={g.get('giayChungNhanId')} ver={g.get('version')})" for g in thieu))
    for h in detail.get("ListHoSoQuet") or []:
        if not isinstance(h, dict):
            continue
        log(f"   HSQ {h.get('hoSoQuetId')}{' (HSQ đã gửi)' if core.safe_int(h.get('hoSoQuetId')) == hsq_id else ''}:")
        for n in day.cac_node(h):
            log(f"     - {mo_ta_thanh_phan(n)} | tp={n.get('thanhPhanHoSoQuetId')} nodeId={n.get('nodeId')}")
    path = os.path.join(thu_muc, f"sau_khi_gui_{time.strftime('%H%M%S')}.json")
    with open(path, "w", encoding="utf-8") as fp:
        json.dump({"ListHoSoQuet": detail.get("ListHoSoQuet"), "ListGiayChungNhan": detail.get("ListGiayChungNhan")},
                  fp, ensure_ascii=False, indent=2, default=str)
    log(f"   Đã lưu trạng thái: {path}")


def gui_hsq(client: core.MplisClient, hoso: dict[str, Any], items: list[tuple[dict, str | None]]) -> None:
    """1 request cho cả HSQ: items = [(infoHoSoQuet, file đính kèm hoặc None)] theo thứ tự."""
    session = client._require_session()
    data = {
        "hoSoQuet": json.dumps(day.ho_so_quet_payload(hoso), ensure_ascii=False),
        "count": str(len(items)),
        "isLuuKhoHoSoQuet": "false",
    }
    headers = dict(session.headers)
    headers.pop("Content-Type", None)
    with contextlib.ExitStack() as stack:
        files = {}
        for i, (info, path) in enumerate(items, 1):
            data[f"infoHoSoQuet_{i}"] = json.dumps(info, ensure_ascii=False)
            if path:
                files[f"fileHoSoQuet_{i}"] = (os.path.basename(path), stack.enter_context(open(path, "rb")),
                                              "application/pdf")
        response = session.post(core.URL_UPDATE_HOSOQUET, data=data, files=files, headers=headers,
                                timeout=core.TIMEOUT, allow_redirects=False)
    result = client._response_json(response, "UpdateHoSoQuetExistFile")
    if not isinstance(result, dict) or not result.get("success"):
        raise RuntimeError(f"Bổ sung HSQ không thành công: {core.rut_gon_text(result)}")


# ============================ XỬ LÝ 1 ĐƠN ============================

def xu_ly_don(
    client: core.MplisClient,
    downloader: thd.Downloader,
    madon: int,
    out_dir: str,
    chay_thu: bool,
) -> dict[str, Any]:
    kq: dict[str, Any] = {"madon": madon, "trang_thai": "Lỗi", "ghi_chu": ""}

    detail = client.lay_chi_tiet_don(madon)
    if detail is None:
        kq["ghi_chu"] = "Không lấy được chi tiết đơn"
        return kq

    ds_gcn = gcn_hien_hanh(detail)
    kq["so_gcn"] = len(ds_gcn)
    hsq_list = [h for h in detail.get("ListHoSoQuet") or [] if isinstance(h, dict)]

    thanh_phan_gcn = [(h, n) for h in hsq_list for n in day.cac_node(h) if la_thanh_phan_gcn(n)]
    id_da_co = {core.safe_int(n.get("giayChungNhanId")) for _, n in thanh_phan_gcn}
    da_co = [g for g in ds_gcn if core.safe_int(g.get("giayChungNhanId")) in id_da_co]
    thieu = [g for g in ds_gcn if core.safe_int(g.get("giayChungNhanId")) not in id_da_co]
    kq["gcn_da_co"] = ", ".join(str(g.get("soPhatHanh")) for g in da_co)
    kq["gcn_bo_sung"] = ", ".join(str(g.get("soPhatHanh")) for g in thieu)
    log(f"Đơn {madon}: {len(ds_gcn)} GCN | đã có file: {kq['gcn_da_co'] or '-'} | thiếu: {kq['gcn_bo_sung'] or '-'}")

    if len(ds_gcn) < 2:
        kq.update(trang_thai="Bỏ qua", ghi_chu="Đơn có dưới 2 GCN")
        return kq
    if not thieu:
        kq.update(trang_thai="Đủ", ghi_chu="Mọi GCN đã có thành phần HSQ")
        return kq
    if len(thanh_phan_gcn) != 1:
        kq.update(trang_thai="Bỏ qua",
                  ghi_chu=f"Có {len(thanh_phan_gcn)} thành phần GCN trong HSQ — cần đúng 1 file GCN làm nguồn")
        return kq

    hoso, nguon = thanh_phan_gcn[0]
    hsq_id = core.safe_int(hoso.get("hoSoQuetId"))
    nodes = day.cac_node(hoso)
    kq["hosoquetid"] = hsq_id
    kq["so_thanh_phan_truoc"] = len(nodes)

    # ---- tải file nguồn
    thu_muc = os.path.join(out_dir, str(madon))
    os.makedirs(thu_muc, exist_ok=True)
    data, ten_server = downloader.tai(nguon)
    file_nguon = os.path.join(
        thu_muc, f"{thd.ten_file_an_toan(nguon.get('moTa') or 'GCN')}{thd.doan_duoi_file(data, ten_server)}"
    )
    with open(file_nguon, "wb") as fp:
        fp.write(data)
    kq["file_nguon"] = file_nguon
    log(f"   Tải file nguồn: {nguon.get('moTa')} → {os.path.basename(file_nguon)} ({len(data) // 1024} KB)")

    # ---- kế hoạch
    items: list[tuple[dict, str | None]] = [(day.info_giu_nguyen(n), None) for n in nodes]
    items += [(info_thanh_phan_moi(nguon, g), file_nguon) for g in thieu]
    log(f"   HSQ {hsq_id}: {len(nodes)} thành phần → {len(items)}")
    for i, (info, path) in enumerate(items, 1):
        log(f"     {i}. {'THÊM' if path else 'giữ '}  {mo_ta_thanh_phan(info)}")

    if chay_thu:
        kq.update(trang_thai="Chạy thử", ghi_chu=f"Sẽ thêm {len(thieu)} thành phần GCN")
        return kq

    node_ids_truoc = {str(n.get("nodeId") or "") for n in nodes}
    gui_hsq(client, hoso, items)

    # ---- xác nhận
    detail_moi = client.lay_chi_tiet_don(madon) or {}
    hoso_moi = next((h for h in detail_moi.get("ListHoSoQuet") or []
                     if isinstance(h, dict) and core.safe_int(h.get("hoSoQuetId")) == hsq_id), None)
    nodes_moi = day.cac_node(hoso_moi) if hoso_moi else []
    kq["so_thanh_phan_sau"] = len(nodes_moi)
    if len(nodes_moi) < len(nodes):
        kq["ghi_chu"] = f"MPLIS còn {len(nodes_moi)}/{len(nodes)} thành phần sau khi gửi!"
        raise day.DungKhanCap(
            f"HSQ {hsq_id} (đơn {madon}) giảm từ {len(nodes)} còn {len(nodes_moi)} thành phần → DỪNG TOÀN BỘ."
        )
    ids_moi = {str(n.get("nodeId") or "") for n in nodes_moi}
    mat = node_ids_truoc - ids_moi
    id_gcn_moi = {core.safe_int(n.get("giayChungNhanId")) for n in nodes_moi if la_thanh_phan_gcn(n)}
    con_thieu = [g.get("soPhatHanh") for g in thieu if core.safe_int(g.get("giayChungNhanId")) not in id_gcn_moi]
    if con_thieu or mat:
        kq["ghi_chu"] = (f"Còn thiếu: {con_thieu}. " if con_thieu else "") + \
                        (f"Thành phần cũ đổi nodeId: {len(mat)}" if mat else "")
        log(f"   ❌ {kq['ghi_chu']}")
        log_trang_thai_sau(detail_moi, hsq_id, thieu, thu_muc)
        return kq
    kq.update(trang_thai="Đã bổ sung", ghi_chu=f"Thêm {len(thieu)} thành phần GCN")
    log(f"   ✅ Đã bổ sung {kq['gcn_bo_sung']} ({len(nodes)} → {len(nodes_moi)} thành phần)")
    return kq


def ghi_excel(path: str, rows: list[dict[str, Any]]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "KetQua"
    ws.append(EXCEL_HEADERS)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    mau = {"Đã bổ sung": "C6EFCE", "Lỗi": "F4CCCC", "Bỏ qua": "FFF2CC", "Chạy thử": "D9EAF7", "Đủ": "E2EFDA"}
    col_tt = EXCEL_HEADERS.index("trang_thai") + 1
    for stt, row in enumerate(rows, 1):
        ws.append([stt if h == "stt" else row.get(h, "") for h in EXCEL_HEADERS])
        fill = mau.get(str(row.get("trang_thai") or ""))
        if fill:
            ws.cell(row=ws.max_row, column=col_tt).fill = PatternFill("solid", fgColor=fill)
    for i, h in enumerate(EXCEL_HEADERS, 1):
        ws.column_dimensions[get_column_letter(i)].width = EXCEL_WIDTHS.get(h, 14)
    ws.freeze_panes = "C2"
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)


def chay(
    client: core.MplisClient,
    ids: list[int],
    out_dir: str,
    chay_thu: bool = True,
    nen_dung: Callable[[], bool] = lambda: False,
    bao_tien_do: Callable[[int, int, str], None] = lambda done, total, label: None,
) -> str:
    """Chạy cho danh sách đơn; trả về đường dẫn Excel kết quả."""
    token = client.session.headers.get("__requestverificationtoken", "")
    downloader = thd.Downloader(None, client.session, token, thd.doc_mau_da_luu(), thd.MAU_MAC_DINH)
    os.makedirs(out_dir, exist_ok=True)
    excel = os.path.join(out_dir, f"ket_qua_bo_sung_gcn_{'thu_' if chay_thu else ''}"
                                  f"{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    rows: list[dict[str, Any]] = []
    log(f"{'[CHẠY THỬ] ' if chay_thu else ''}Bổ sung GCN cho {len(ids)} đơn...")
    try:
        for n, madon in enumerate(ids, 1):
            if nen_dung():
                log("⏹ Đã dừng theo yêu cầu.")
                break
            bao_tien_do(n - 1, len(ids), f"Đơn {madon} ({n}/{len(ids)})")
            log(f"--- [{n}/{len(ids)}] Đơn {madon} ---")
            try:
                kq = xu_ly_don(client, downloader, madon, out_dir, chay_thu)
            except day.DungKhanCap:
                rows.append({"madon": madon, "trang_thai": "Lỗi", "ghi_chu": "MPLIS giảm thành phần — ĐÃ DỪNG"})
                raise
            except Exception as exc:
                kq = {"madon": madon, "trang_thai": "Lỗi", "ghi_chu": core.rut_gon_text(exc, 300)}
                log(f"   ❌ {kq['ghi_chu']}")
            rows.append(kq)
            ghi_excel(excel, rows)
            time.sleep(core.REQUEST_DELAY_SECONDS)
    finally:
        ghi_excel(excel, rows)

    dem = {}
    for r in rows:
        dem[r.get("trang_thai")] = dem.get(r.get("trang_thai"), 0) + 1
    bao_tien_do(len(rows), len(ids), "Hoàn tất")
    log("=" * 50)
    log(("[CHẠY THỬ — chưa gửi gì] " if chay_thu else "✅ ") + " | ".join(f"{k}: {v}" for k, v in dem.items()))
    log(f"Excel kết quả: {excel}")
    return excel
