# -*- coding: utf-8 -*-
"""
Đẩy file GCN theo danh sách mã đơn trong Excel — file tìm trong 1 folder (kể cả folder con).

    Excel: cột 'madon' (hoặc tinhhinhdangkyid...; không có header thì lấy cột A).
    <folder>/.../24221_GCN_DH 770214_signed.pdf  ← tên file CHỨA số phát hành GCN; không phân biệt hoa thường,
                                                   dấu cách / '_' / '-' / '.' giữa các ký tự ('DH770214',
                                                   'DH_770214' đều khớp 'DH 770214'); không khớp khi dính liền
                                                   chữ/số phía trước hoặc số phía sau ('DH 7702145')

Với mỗi mã đơn:
  1. Lấy chi tiết đơn → GCN hiện hành (bỏ isLastest = False).
  2. Mỗi GCN: tìm file có tên chứa số phát hành. Nhiều file cùng khớp → luôn chọn 1: tên có chữ 'GCN' >
     đã ký số > đường dẫn xếp trước. Không có file → bỏ qua GCN đó. Tên chứa 2 số GCN của đơn → gán cho cả 2.
     Đơn CHƯA CÓ GCN hiện hành → không tìm file: chọn 1 thành phần có sẵn trong HSQ của đơn (đã là Đơn đăng ký >
     mô tả có chữ 'đơn' > 'giấy tờ' > thành phần đầu) đổi loại Đơn đăng ký, thành phần khác của HSQ đó →
     Giấy tờ (không đổi file). Đơn chưa có HSQ → bỏ qua.
  3. Gán file cho GCN như day_hsq_mota_gan_gcn (viec_cho_gcn): thành phần đã gắn GCN → thay; thành phần có
     mô tả trùng tên file → thay + ghi nhận GCN; không có → thêm thành phần 'Giấy chứng nhận <số>'; đơn chưa
     có HSQ → tạo HSQ mới.
  4. HSQ có đẩy GCN → mọi thành phần khác chuyển loại Giấy tờ (không đổi file). Gửi + kiểm tra lại như
     day_hsq_mota_gan_gcn (giảm thành phần → DỪNG TOÀN BỘ).
"""

from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path
from typing import Any, Callable

import bo_sung_gcn_hsq as bs
import cap_nhat_don_hsq as core
import nhom1_mplis as nhom1
import day_hsq_ky_so as day
import day_hsq_mota_gan_gcn as gg
import day_hsq_theo_mota as dm
import tai_hsq_don as thd

log = thd.log


# ============================ TÌM FILE THEO SỐ GCN ============================

SO_GCN_TOI_THIEU = 5  # số phát hành chuẩn hoá ngắn hơn → không khớp kiểu "chứa trong tên" (dễ nhầm)


def ten_chuan(ten: str) -> str:
    """Bỏ dấu, chữ thường; giữ dấu cách/_/- để còn nhận ra ranh giới khi khớp số GCN."""
    s = core.remove_accents(str(ten or "")).lower().replace("đ", "d").replace("ð", "d")
    return s.strip()


def mau_so_gcn(so: str) -> re.Pattern | None:
    """
    'DH 770214' → khớp 'dh770214', 'dh 770214', 'dh_770214', 'dh-770214' nằm TRONG tên file, nhưng không khớp
    khi dính liền chữ/số ở trước ('xdh770214') hoặc số ở sau ('dh7702145').
    """
    loi = re.sub(r"[^a-z0-9]+", "", ten_chuan(so))
    if len(loi) < SO_GCN_TOI_THIEU:
        return None
    giua = r"[\s_.\-]*".join(re.escape(c) for c in loi)
    return re.compile(rf"(?<![a-z0-9]){giua}(?![0-9])")


def quet_pdf(goc: str) -> list[tuple[str, str]]:
    """Mọi file PDF trong `goc` (kể cả folder con) → [(đường dẫn, tên file chuẩn hoá, bỏ đuôi .pdf)]."""
    kq: list[tuple[str, str]] = []
    for thu_muc, _, files in os.walk(goc):
        for f in files:
            if f.lower().endswith(".pdf"):
                kq.append((os.path.join(thu_muc, f), ten_chuan(Path(f).stem)))
    return kq


def tim_file_gcn(ds_gcn: list[dict], ds_pdf: list[tuple[str, str]],
                 ) -> tuple[dict[int, tuple[str, str, int]], list[tuple[str, dict | None, str]]]:
    """
    Tên file CHỨA số phát hành GCN → ({gcnId: (file, tên file, 0)} cho gg.viec_cho_gcn, [(file, None, lý do)]).
    Tên chứa 2 số GCN của đơn → gán cho cả 2.
    """
    gan: dict[int, tuple[str, str, int]] = {}
    bo_qua: list[tuple[str, dict | None, str]] = []
    for gcn in ds_gcn:
        so = str(gcn.get("soPhatHanh") or "").strip()
        mau = mau_so_gcn(so)
        if mau is None:
            if so:
                bo_qua.append(("", None, f"GCN {so}: số phát hành quá ngắn, không tìm theo tên file"))
            continue
        files = [p for p, ten in ds_pdf if mau.search(ten)]
        if not files:
            continue
        p = chon_1_file(files)
        if len(files) > 1:
            log(f"   ℹ GCN {so}: {len(files)} file cùng khớp → chọn {p}")
        gan[core.safe_int(gcn.get("giayChungNhanId"))] = (p, Path(p).stem, 0)
        bo_qua += [(f, None, f"Đã dùng {p} cho GCN {so}") for f in files if f != p]
    return gan, bo_qua


def chon_1_file(files: list[str]) -> str:
    """Nhiều file cùng khớp 1 GCN → luôn chọn 1: tên có chữ 'GCN' > đã ký số > đường dẫn xếp trước."""
    return min(files, key=lambda p: ("gcn" not in ten_chuan(Path(p).stem), not day.pdf_da_ky_so(p), p.lower()))


# ============================ ĐƠN CHƯA CÓ GCN → ĐƠN ĐĂNG KÝ ============================

LOAI_GIAY_TO, LOAI_DON_DANG_KY = gg.LOAI_GIAY_TO, gg.LOAI_DON_DANG_KY
# Ưu tiên chọn làm Đơn đăng ký: thành phần có mô tả chứa từ "đơn" (không ăn nhầm "đồng"), rồi "giấy tờ"
UU_TIEN_DON = (re.compile(r"(?<![a-z0-9])don(?![a-z0-9])"), re.compile(r"(?<![a-z0-9])giay[\s_.\-]*to(?![a-z0-9])"))


def chon_thanh_phan_don(ds_tp: list[dict[str, Any]]) -> dict[str, Any]:
    """Đã là Đơn đăng ký > mô tả có chữ 'đơn' > mô tả có chữ 'giấy tờ' > thành phần đầu tiên."""
    def hang(tp: dict[str, Any]) -> tuple[int, int]:
        node = tp["node"]
        if core.safe_int(node.get("loaiHoSoQuet"), -1) == LOAI_DON_DANG_KY:
            return 0, tp["stt"]
        ten = ten_chuan(f"{node.get('moTa') or ''} {node.get('tenGiayTo') or ''}")
        for i, mau in enumerate(UU_TIEN_DON, 1):
            if mau.search(ten):
                return i, tp["stt"]
        return len(UU_TIEN_DON) + 1, tp["stt"]
    return min(ds_tp, key=hang)


def _dong_tp(madon: int, tp: dict[str, Any], viec: str, **kw) -> dict[str, Any]:
    node = tp["node"]
    loai = core.safe_int(node.get("loaiHoSoQuet"), -1)
    return {"madon": madon, "hosoquetid": tp["hoso"].get("hoSoQuetId", ""), "vitri": tp["vi_tri"] + 1,
            "mota": node.get("moTa") or "", "loaihosoquet": core.LOAI_HO_SO_QUET_OPTIONS.get(loai, ""),
            "viec": viec, **kw}


def xu_ly_don_khong_gcn(client: core.MplisClient, madon: int, detail: dict[str, Any],
                        chay_thu: bool) -> list[dict[str, Any]]:
    """
    Đơn chưa có GCN hiện hành: chọn 1 thành phần có sẵn trong HSQ của đơn → loại Đơn đăng ký (không đẩy file),
    mọi thành phần khác của HSQ đó → Giấy tờ. Đơn chưa có HSQ → bỏ qua.
    """
    ds_tp = dm.thanh_phan_cua_don(detail)
    if not ds_tp:
        log("   ⚠ Đơn chưa có GCN hiện hành và chưa có HSQ → bỏ qua")
        return [gg._dong(madon, trang_thai="Bỏ qua", ghi_chu="Đơn chưa có GCN hiện hành và chưa có HSQ")]

    chon = chon_thanh_phan_don(ds_tp)
    hoso = chon["hoso"]
    tp_hsq = [t for t in ds_tp if t["hoso"] is hoso]
    sang_giay_to = [t for t in tp_hsq if t is not chon and not (
        core.safe_int(t["node"].get("loaiHoSoQuet"), -1) == LOAI_GIAY_TO and not t["node"].get("laGiayChungNhan"))]
    doi_don = core.safe_int(chon["node"].get("loaiHoSoQuet"), -1) != LOAI_DON_DANG_KY

    log(f"   Đơn chưa có GCN hiện hành → HSQ {hoso.get('hoSoQuetId')}:")
    log(f"     {'Đổi loại Đơn đăng ký' if doi_don else 'Giữ Đơn đăng ký'}  —  {chon['node'].get('moTa')}")
    for t in sang_giay_to:
        log(f"     Chuyển Giấy tờ  —  {t['node'].get('moTa')}")
    if not doi_don and not sang_giay_to:
        log("     (Đã đúng loại — không cần gửi)")
        return [_dong_tp(madon, chon, "Giữ Đơn đăng ký", trang_thai="Giữ nguyên", ghi_chu="Đã đúng loại")]

    viec_don = "Đổi loại Đơn đăng ký" if doi_don else "Giữ Đơn đăng ký"
    if chay_thu:
        return ([_dong_tp(madon, chon, viec_don, trang_thai="Chạy thử", ghi_chu=f"Sẽ {viec_don.lower()}")]
                + [_dong_tp(madon, t, "Chuyển Giấy tờ", trang_thai="Chạy thử", ghi_chu="Sẽ chuyển loại Giấy tờ")
                   for t in sang_giay_to])

    # ---- gửi: ĐỦ thành phần của HSQ, không đính kèm file (files=null → giữ file cũ)
    vi_tri_giay_to = {t["vi_tri"] for t in sang_giay_to}
    items: list[tuple[dict, str | None]] = []
    for i, node in enumerate(day.cac_node(hoso)):
        info = day.info_giu_nguyen(node)
        info["_id"] = i + 1  # như MPLIS web: số thứ tự thành phần
        if i == chon["vi_tri"]:
            gg.dat_loai_don_dang_ky(info)
        elif i in vi_tri_giay_to:
            gg.dat_loai_giay_to(info)
        items.append((info, None))
    try:
        bs.gui_hsq(client, hoso, items)
    except core.HetPhien:
        raise
    except Exception as exc:
        ghi_chu = core.rut_gon_text(exc, 300)
        log(f"   ❌ {ghi_chu}")
        return [_dong_tp(madon, t, v, trang_thai="Lỗi", ghi_chu=ghi_chu)
                for t, v in [(chon, viec_don)] + [(t, "Chuyển Giấy tờ") for t in sang_giay_to]]

    # ---- xác nhận
    hsq_id = core.safe_int(hoso.get("hoSoQuetId"))
    detail_moi = client.lay_chi_tiet_don(madon) or {}
    hoso_moi = next((h for h in detail_moi.get("ListHoSoQuet") or []
                     if isinstance(h, dict) and core.safe_int(h.get("hoSoQuetId")) == hsq_id), None)
    nodes_moi = day.cac_node(hoso_moi) if hoso_moi else []
    if len(nodes_moi) < len(day.cac_node(hoso)):
        raise day.DungKhanCap(f"HSQ {hsq_id} (đơn {madon}) giảm từ {len(day.cac_node(hoso))} còn {len(nodes_moi)} "
                              "thành phần sau khi đẩy → DỪNG TOÀN BỘ.")

    def loai_moi(t: dict[str, Any]) -> int:
        tp_id = core.safe_int(t["node"].get("thanhPhanHoSoQuetId"))
        n = next((n for n in nodes_moi if core.safe_int(n.get("thanhPhanHoSoQuetId")) == tp_id), None) or {}
        return core.safe_int(n.get("loaiHoSoQuet"), -1)

    rows = []
    for t, viec, loai_can in [(chon, viec_don, LOAI_DON_DANG_KY)] + [(t, "Chuyển Giấy tờ", LOAI_GIAY_TO)
                                                                    for t in sang_giay_to]:
        if loai_moi(t) == loai_can:
            rows.append(_dong_tp(madon, t, viec, trang_thai="Đã đẩy", ghi_chu=time.strftime("%d/%m/%Y %H:%M:%S")))
            log(f"   ✅ {viec}: {t['node'].get('moTa')}")
        else:
            rows.append(_dong_tp(madon, t, viec, trang_thai="Lỗi",
                                 ghi_chu="Server báo OK nhưng loại chưa đổi (kiểm tra trên MPLIS)"))
            log(f"   ❌ {viec}: {t['node'].get('moTa')} — server báo OK nhưng loại chưa đổi")
    return rows


# ============================ XỬ LÝ 1 ĐƠN ============================

def xu_ly_don(client: core.MplisClient, madon: int, ds_pdf: list[tuple[str, str]],
              chay_thu: bool, bo_qua_nhom_1: bool = True) -> list[dict[str, Any]]:
    detail = client.lay_chi_tiet_don(madon)
    if detail is None:
        return [gg._dong(madon, trang_thai="Lỗi", ghi_chu=dm.LOI_CHI_TIET)]
    if bo_qua_nhom_1 and nhom1.da_dat(client, madon, detail):
        return [gg._dong(madon, trang_thai="Bỏ qua", ghi_chu=nhom1.LY_DO_DA_DAT)]
    ds_gcn = bs.gcn_hien_hanh(detail)
    if not ds_gcn:
        return xu_ly_don_khong_gcn(client, madon, detail, chay_thu)

    gan, bo_qua = tim_file_gcn(ds_gcn, ds_pdf)
    so_ds = "; ".join(str(g.get("soPhatHanh") or g.get("giayChungNhanId")) for g in ds_gcn)
    if not gan:
        if not bo_qua:
            ly_do = f"Không có file nào tên chứa số GCN ({so_ds}) trong folder"
            log(f"   ⚠ {ly_do}")
            return [gg._dong(madon, trang_thai="Bỏ qua", ghi_chu=ly_do)]
        for p, _, ld in bo_qua:
            log(f"   ⚠ {os.path.basename(p) + ': ' if p else ''}{ld}")
        return [gg._dong(madon, p, trang_thai="Bỏ qua", ghi_chu=ld) for p, _, ld in bo_qua]

    viec, bq_gcn, _ = gg.viec_cho_gcn(detail, gan, ly_do_thieu="không có file nào tên chứa số GCN này trong folder")
    return gg.thuc_hien_ke_hoach(client, madon, detail, viec, bo_qua + bq_gcn, chay_thu, giay_to=True)


# ============================ TIẾN ĐỘ (chạy tiếp) ============================

FILE_TIEN_DO = "tien_do_day_gcn_theo_excel.json"


def _file_tien_do(excel_path: str) -> str:
    return os.path.join(os.path.dirname(excel_path), FILE_TIEN_DO)


def doc_tien_do(excel_path: str) -> dict[str, Any]:
    try:
        with open(_file_tien_do(excel_path), encoding="utf-8") as fp:
            data = json.load(fp)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        log(f"⚠ Không đọc được {FILE_TIEN_DO} ({exc}) — coi như chưa đơn nào xong.")
        return {}


def ghi_tien_do(excel_path: str, tien_do: dict[str, Any]) -> None:
    path = _file_tien_do(excel_path)
    tam = path + ".tmp"
    with open(tam, "w", encoding="utf-8") as fp:
        json.dump(tien_do, fp, ensure_ascii=False, indent=1)
    os.replace(tam, path)  # ghi file tạm rồi thay: tắt ngang không hỏng file tiến độ


def dau_van_tay(paths: list[str]) -> list[list]:
    """[[đường dẫn, kích thước, giờ sửa]] — file đổi / bị xoá thì khác → đơn chạy lại."""
    kq = []
    for p in paths:
        st = os.stat(p)
        kq.append([os.path.normcase(os.path.abspath(p)), st.st_size, int(st.st_mtime)])
    return sorted(kq)


def don_da_xong(tien_do: dict[str, Any], madon: int) -> bool:
    muc = tien_do.get(str(madon))
    if not isinstance(muc, dict):
        return False
    try:
        return dau_van_tay([f[0] for f in muc.get("files") or []]) == muc.get("files")
    except OSError:  # file đã đẩy bị xoá / đổi chỗ
        return False


def ket_qua_xong(rows: list[dict[str, Any]]) -> list[str] | None:
    """
    Đơn xong hẳn → [file đã đẩy] để ghi tiến độ; còn Lỗi / GCN thiếu file / chưa làm gì → None (lần sau chạy lại).
    Dòng 'Bỏ qua' vì file trùng đã có file khác dùng thay ('Đã dùng ...') không tính là chưa xong.
    """
    if any(r.get("trang_thai") == "Lỗi" for r in rows):
        return None
    if any(r.get("trang_thai") == "Bỏ qua" and not str(r.get("ghi_chu") or "").startswith("Đã dùng") for r in rows):
        return None
    if not any(r.get("trang_thai") in ("Đã đẩy", "Giữ nguyên") for r in rows):
        return None
    return [os.path.join(r["thu_muc"], r["file"]) for r in rows
            if r.get("trang_thai") == "Đã đẩy" and r.get("file") and r.get("thu_muc")]


# ============================ CHẠY HÀNG LOẠT ============================

def chay(
    client: core.MplisClient,
    excel_path: str,
    goc: str,
    chay_thu: bool = True,
    nen_dung: Callable[[], bool] = lambda: False,
    bao_tien_do: Callable[[int, int, str], None] = lambda done, total, label: None,
    bo_qua_da_xong: bool = True,
) -> str:
    """Chạy cho mọi mã đơn trong Excel; trả về đường dẫn Excel kết quả (ghi cạnh Excel đầu vào).

    Đơn đẩy xong không lỗi được ghi vào FILE_TIEN_DO (cạnh Excel mã đơn); bo_qua_da_xong → lần sau bỏ qua
    (trừ khi file GCN đã đẩy bị sửa / xoá).
    """
    ids = thd.tach_danh_sach_id(thd.doc_id_tu_excel(excel_path))
    if not ids:
        raise RuntimeError(f"Không đọc được mã đơn nào trong {excel_path}")
    tien_do = doc_tien_do(excel_path)
    if bo_qua_da_xong and tien_do:
        da_xong = [m for m in ids if don_da_xong(tien_do, m)]
        if da_xong:
            log(f"⏭ Bỏ qua {len(da_xong)} đơn đã đẩy xong ({FILE_TIEN_DO}); còn {len(ids) - len(da_xong)} đơn.")
            ids = [m for m in ids if m not in set(da_xong)]
        if not ids:
            log("✅ Mọi đơn đều đã đẩy xong — không còn gì để chạy. (Bỏ tích 'Bỏ qua đơn đã xong' để chạy lại.)")
            return ""
    ds_pdf = quet_pdf(goc)
    if not ds_pdf:
        raise RuntimeError(f"Không có file PDF nào trong {goc} (kể cả folder con)")
    excel = os.path.join(os.path.dirname(excel_path), f"ket_qua_day_gcn_theo_excel_{'thu_' if chay_thu else ''}"
                                                      f"{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    log(f"{'[CHẠY THỬ] ' if chay_thu else ''}{len(ids)} đơn | {len(ds_pdf)} file PDF "
        "trong folder | HSQ có đẩy GCN: thành phần khác → Giấy tờ")

    rows: list[dict[str, Any]] = []
    loi_lien_tiep = 0
    try:
        for n, madon in enumerate(ids, 1):
            if nen_dung():
                log("⏹ Đã dừng theo yêu cầu.")
                break
            bao_tien_do(n - 1, len(ids), f"Đơn {madon} ({n}/{len(ids)})")
            log(f"--- [{n}/{len(ids)}] Đơn {madon} ---")
            try:
                dong = xu_ly_don(client, madon, ds_pdf, chay_thu)
                rows += dong
                files_xong = None if chay_thu else ket_qua_xong(dong)
                if files_xong is not None:
                    tien_do[str(madon)] = {"luc": time.strftime("%d/%m/%Y %H:%M:%S"),
                                           "files": dau_van_tay(files_xong)}
                    ghi_tien_do(excel_path, tien_do)
                loi_lien_tiep = loi_lien_tiep + 1 if dong and dong[0].get("ghi_chu") == dm.LOI_CHI_TIET else 0
                if loi_lien_tiep >= dm.DUNG_SAU_SO_DON_LOI:
                    log(f"❌ {loi_lien_tiep} đơn liên tiếp không đọc được từ MPLIS — có thể đã hết phiên đăng nhập. "
                        "ĐÃ DỪNG: lấy lại session (hoặc dán token + cookie mới) rồi chạy lại.")
                    break
            except day.DungKhanCap:
                rows.append(gg._dong(madon, trang_thai="Lỗi", ghi_chu="MPLIS giảm thành phần sau khi đẩy — ĐÃ DỪNG"))
                raise
            except Exception as exc:
                rows.append(gg._dong(madon, trang_thai="Lỗi", ghi_chu=core.rut_gon_text(exc, 300)))
                log(f"   ❌ {core.rut_gon_text(exc, 300)}")
                if isinstance(exc, core.HetPhien):
                    log(f"❌ {exc} — ĐÃ DỪNG: lấy lại session (hoặc dán token + cookie mới) rồi chạy lại.")
                    break
            excel = gg._ghi_an_toan(excel, rows)
            core.nghi_giua_don()
    finally:
        if rows:
            excel = gg._ghi_an_toan(excel, rows)

    dem: dict[str, int] = {}
    for r in rows:
        if r.get("viec") or r.get("file"):
            dem[r["trang_thai"]] = dem.get(r["trang_thai"], 0) + 1
    bao_tien_do(len(ids), len(ids), "Hoàn tất")
    log("=" * 50)
    log(("[CHẠY THỬ — chưa gửi gì] " if chay_thu else "✅ ") + " | ".join(f"{k}: {v}" for k, v in dem.items()))
    log(f"Excel kết quả: {excel}")
    return excel
