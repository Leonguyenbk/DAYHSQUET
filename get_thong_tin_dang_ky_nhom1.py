# -*- coding: utf-8 -*-
"""
Tool tra cứu thông tin đăng ký nhóm 1 từ MPLIS
Endpoint: https://dla.mplis.gov.vn/dc/LamSachDuLieuAjax/GetThongTinDangKyNhom1

Chức năng:
1. Gửi request POST tới GetThongTinDangKyNhom1 với tinhHinhDangKyId
2. Lấy response JSON và chuyển đổi trường "ngayVaoSo": "/Date(...)/" thành định dạng "DD/MM/YYYY" (múi giờ GMT+7)
3. Hỗ trợ chạy cả Giao diện (GUI Tkinter) lẫn Dòng lệnh (CLI)
4. Tích hợp sẵn chế độ Mock/Sample test dữ liệu người dùng cung cấp
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import requests

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False


# ============================ CẤU HÌNH API ============================

API_URL = "https://dla.mplis.gov.vn/dc/LamSachDuLieuAjax/GetThongTinDangKyNhom1"
REFERER_URL = "https://dla.mplis.gov.vn/dc/lamsachdulieu"
ORIGIN_URL = "https://dla.mplis.gov.vn"
DEFAULT_TIMEOUT = 30
TIMEZONE_VN = timezone(timedelta(hours=7))

HEADERS_DEFAULT = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Content-Type": "application/json; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": REFERER_URL,
    "Origin": ORIGIN_URL,
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}


# ============================ XỬ LÝ NGÀY THÁNG ============================

def parse_dotnet_date(date_val: Any) -> Optional[datetime]:
    """
    Parse chuỗi ngày định dạng .NET WCF /Date(1528995600000)/ sang datetime theo giờ VN (GMT+7).
    Hỗ trợ timestamp âm hoặc dương.
    """
    if not isinstance(date_val, str):
        return None

    match = re.search(r"/Date\((-?\d+)(?:[+-]\d{4})?\)/", date_val)
    if not match:
        return None

    ms = int(match.group(1))
    try:
        dt = datetime.fromtimestamp(ms / 1000, tz=timezone.utc)
        return dt.astimezone(TIMEZONE_VN)
    except (ValueError, OverflowError, OSError):
        # Phòng trường hợp timestamp quá xa (ví dụ -62135596800000 cho năm 0001)
        base = datetime(1970, 1, 1, tzinfo=timezone.utc)
        dt = base + timedelta(milliseconds=ms)
        return dt.astimezone(TIMEZONE_VN)


def dotnet_date_to_dmy(date_val: Any) -> str:
    """
    Chuyển /Date(1528995600000)/ thành chuỗi DD/MM/YYYY.
    Ví dụ: /Date(1528995600000)/ -> '15/06/2018'
    """
    dt = parse_dotnet_date(date_val)
    if dt is not None:
        return dt.strftime("%d/%m/%Y")
    return str(date_val) if date_val is not None else ""


def dotnet_date_to_dmy_hms(date_val: Any) -> str:
    """Chuyển /Date(...)/ thành DD/MM/YYYY HH:MM:SS."""
    dt = parse_dotnet_date(date_val)
    if dt is not None:
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    return str(date_val) if date_val is not None else ""


def convert_ngay_vao_so(data: Any, convert_all_dates: bool = False) -> Any:
    """
    Duyệt đệ quy dict/list để chuyển:
    - Mặc định: chuyển tất cả trường 'ngayVaoSo' có dạng /Date(...)/ thành 'DD/MM/YYYY'.
    - Nếu convert_all_dates=True: chuyển cả các trường thoiDiemDangKy, ngayCap, CreatedDate,...
    """
    if isinstance(data, dict):
        new_dict = {}
        for k, v in data.items():
            if k == "ngayVaoSo" and isinstance(v, str) and "/Date(" in v:
                new_dict[k] = dotnet_date_to_dmy(v)
            elif convert_all_dates and isinstance(v, str) and "/Date(" in v:
                if k in ("thoiDiemDangKy", "thoiDiemDangKyLanDau", "CreatedDate", "ModifiedDate"):
                    new_dict[k] = dotnet_date_to_dmy_hms(v)
                else:
                    new_dict[k] = dotnet_date_to_dmy(v)
            else:
                new_dict[k] = convert_ngay_vao_so(v, convert_all_dates=convert_all_dates)
        return new_dict
    elif isinstance(data, list):
        return [convert_ngay_vao_so(item, convert_all_dates=convert_all_dates) for item in data]
    return data


# ============================ GỌI API MPLIS ============================

def call_api_get_thong_tin_nhom1(
    tinh_hinh_dang_ky_id: int,
    cookie_str: str = "",
    headers_extra: Optional[Dict[str, str]] = None,
    timeout: int = DEFAULT_TIMEOUT,
) -> Dict[str, Any]:
    """
    Gửi HTTP POST tới https://dla.mplis.gov.vn/dc/LamSachDuLieuAjax/GetThongTinDangKyNhom1
    """
    headers = dict(HEADERS_DEFAULT)
    if headers_extra:
        headers.update(headers_extra)

    session = requests.Session()
    session.headers.update(headers)

    if cookie_str.strip():
        # Phân tách cookie chuỗi dạng "name=val; name2=val2"
        for part in cookie_str.split(";"):
            part = part.strip()
            if "=" in part:
                k, v = part.split("=", 1)
                session.cookies.set(k.strip(), v.strip())

    payload = {"tinhHinhDangKyId": int(tinh_hinh_dang_ky_id)}

    try:
        response = session.post(
            API_URL,
            json=payload,
            timeout=timeout,
        )
        if not response.ok:
            return {
                "success": False,
                "error": f"Lỗi HTTP {response.status_code}: {response.text[:500]}",
                "status_code": response.status_code,
            }
        try:
            js = response.json()
            return js
        except Exception as e:
            return {
                "success": False,
                "error": f"Response không phải JSON hợp lệ: {e}. Raw: {response.text[:500]}",
                "raw_text": response.text,
            }
    except Exception as e:
        return {
            "success": False,
            "error": f"Lỗi kết nối tới API: {e}",
        }


# ============================ TRÍCH XUẤT THÔNG TIN TÓM TẮT ============================

def extract_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Trích xuất các thông tin quan trọng từ response để người dùng xem nhanh:
    - Số phát hành GCN, Số vào sổ, Ngày vào sổ (đã chuyển đổi DD/MM/YYYY)
    - Chủ sở hữu
    - Thửa đất (số tờ, thửa, diện tích, địa chỉ, mục đích)
    - Hồ sơ quét
    """
    summary = {
        "success": data.get("success", False),
        "tinhHinhDangKyId": None,
        "thoiDiemDangKy": None,
        "soPhatHanh": None,
        "soVaoSo": None,
        "ngayVaoSo_goc": None,
        "ngayVaoSo_dmy": None,
        "chuSoHuu": [],
        "thuaDat": [],
        "hoSoQuet": [],
    }

    # 1. Tìm trong value
    values = data.get("value") or []
    if values and isinstance(values, list):
        val0 = values[0]
        summary["tinhHinhDangKyId"] = val0.get("tinhHinhDangKyId")
        summary["thoiDiemDangKy"] = dotnet_date_to_dmy(val0.get("thoiDiemDangKy"))

        list_gcn = val0.get("ListGiayChungNhan") or []
        if list_gcn and isinstance(list_gcn, list):
            gcn0 = list_gcn[0]
            summary["soPhatHanh"] = gcn0.get("soPhatHanh")
            summary["soVaoSo"] = gcn0.get("soVaoSo")
            summary["ngayVaoSo_goc"] = gcn0.get("ngayVaoSo")
            summary["ngayVaoSo_dmy"] = dotnet_date_to_dmy(gcn0.get("ngayVaoSo"))

            # Chủ sở hữu
            csh = gcn0.get("ChuSoHuu") or {}
            # Cá nhân
            for cn in csh.get("CaNhans") or []:
                summary["chuSoHuu"].append({
                    "loai": "Cá nhân",
                    "hoTen": cn.get("hoTen"),
                    "namSinh": cn.get("namSinh"),
                    "cccd": cn.get("maSoDinhDanh"),
                    "diaChi": cn.get("diaChi"),
                })
            # Vợ chồng
            for vc in csh.get("VoChongs") or []:
                chong = vc.get("Chong") or {}
                vo = vc.get("Vo") or {}
                chong_gttt = [g.get("soGiayTo") for g in (chong.get("GiayToTuyThans") or []) if g.get("soGiayTo")]
                vo_gttt = [g.get("soGiayTo") for g in (vo.get("GiayToTuyThans") or []) if g.get("soGiayTo")]
                summary["chuSoHuu"].append({
                    "loai": "Vợ chồng",
                    "chong": f"{chong.get('hoTen')} ({chong.get('namSinh')}) - Số GTTT: {', '.join(chong_gttt)}",
                    "vo": f"{vo.get('hoTen')} ({vo.get('namSinh')}) - Số GTTT: {', '.join(vo_gttt)}",
                    "diaChi": chong.get("diaChi") or vo.get("diaChi"),
                })

            # Thửa đất
            taisan = gcn0.get("TaiSan") or {}
            for td in taisan.get("ThuaDats") or []:
                mdsd_list = []
                for m in td.get("MucDichSuDungs") or []:
                    mdsd_list.append(f"{m.get('loaiMucDichSuDungId')} ({m.get('dienTich')}m², {m.get('thoiHanSuDung')})")
                summary["thuaDat"].append({
                    "thuaDatId": td.get("thuaDatId"),
                    "soTo": td.get("soHieuToBanDo"),
                    "soThua": td.get("soThuTuThua"),
                    "dienTich": td.get("dienTich"),
                    "diaChi": td.get("diaChi"),
                    "mucDichSuDung": ", ".join(mdsd_list),
                })

    # 2. Tìm thêm hồ sơ quét trong thongTinDangKys
    ttdk_list = data.get("thongTinDangKys") or []
    if ttdk_list and isinstance(ttdk_list, list):
        ttdk0 = ttdk_list[0]
        if not summary["tinhHinhDangKyId"]:
            thdk = ttdk0.get("TinhHinhDangKy") or {}
            summary["tinhHinhDangKyId"] = thdk.get("tinhHinhDangKyId")
            summary["thoiDiemDangKy"] = dotnet_date_to_dmy(thdk.get("thoiDiemDangKy"))

        for hsq in ttdk0.get("ListHoSoQuet") or []:
            hsq_id = hsq.get("hoSoQuetId")
            files = []
            lf = (hsq.get("ListFileHoSoQuet") or {}).get("ListFileHoSoQuet") or []
            for f in lf:
                files.append(f"{f.get('moTa')} (ID file: {f.get('thanhPhanHoSoQuetId')})")
            summary["hoSoQuet"].append({
                "hoSoQuetId": hsq_id,
                "files": files,
            })

    return summary


# ============================ GIAO DIỆN TKINTER ============================

class ToolNhom1GUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Công cụ tra cứu & chuyển đổi Ngày Vào Sổ - GetThongTinDangKyNhom1")
        self.root.geometry("980, 720")
        self.root.minsize(800, 600)

        self.current_data: Optional[Dict[str, Any]] = None
        self.converted_data: Optional[Dict[str, Any]] = None

        self._setup_ui()

    def _setup_ui(self):
        # 1. Khung cấu hình / Input
        frame_top = ttk.LabelFrame(self.root, text=" Cấu hình Tra Cứu / Input ", padding=10)
        frame_top.pack(fill="x", padx=10, pady=5)

        # Hàng 1: Endpoint & tinhHinhDangKyId
        lbl_id = ttk.Label(frame_top, text="tinhHinhDangKyId:", font=("Segoe UI", 9, "bold"))
        lbl_id.grid(row=0, column=0, sticky="w", padx=5, pady=4)

        self.entry_id = ttk.Entry(frame_top, width=25, font=("Segoe UI", 10))
        self.entry_id.insert(0, "12921026")
        self.entry_id.grid(row=0, column=1, sticky="w", padx=5, pady=4)

        lbl_ep = ttk.Label(frame_top, text="Endpoint:", font=("Segoe UI", 9))
        lbl_ep.grid(row=0, column=2, sticky="w", padx=10, pady=4)

        lbl_url = ttk.Label(frame_top, text=API_URL, foreground="blue", font=("Segoe UI", 9, "italic"))
        lbl_url.grid(row=0, column=3, sticky="w", padx=5, pady=4)

        # Hàng 2: Cookie (nếu gọi MPLIS thật)
        lbl_ck = ttk.Label(frame_top, text="Cookie MPLIS:")
        lbl_ck.grid(row=1, column=0, sticky="w", padx=5, pady=4)

        self.entry_cookie = ttk.Entry(frame_top, width=60)
        self.entry_cookie.grid(row=1, column=1, columnspan=3, sticky="we", padx=5, pady=4)
        frame_top.columnconfigure(3, weight=1)

        # Hàng 3: Tùy chọn chuyển ngày
        self.var_convert_all = tk.BooleanVar(value=False)
        chk_all = ttk.Checkbutton(
            frame_top,
            text="Chuyển đổi toàn bộ ngày /Date(...)/ trong JSON (thay vì chỉ 'ngayVaoSo')",
            variable=self.var_convert_all,
            command=self._on_toggle_convert_all,
        )
        chk_all.grid(row=2, column=1, columnspan=3, sticky="w", padx=5, pady=2)

        # Hàng 4: Các nút chức năng
        frame_buttons = ttk.Frame(frame_top)
        frame_buttons.grid(row=3, column=0, columnspan=4, sticky="we", pady=6)

        btn_sample = ttk.Button(frame_buttons, text="🧪 1. Chạy Dữ Liệu Mẫu (Sample)", command=self.load_sample_data)
        btn_sample.pack(side="left", padx=5)

        btn_open_file = ttk.Button(frame_buttons, text="📂 2. Mở File JSON Khác", command=self.open_json_file)
        btn_open_file.pack(side="left", padx=5)

        btn_call_api = ttk.Button(frame_buttons, text="🌐 3. Gọi API MPLIS Thật", command=self.call_api)
        btn_call_api.pack(side="left", padx=5)

        btn_save = ttk.Button(frame_buttons, text="💾 Lưu JSON Đã Chuyển Đổi", command=self.save_converted_json)
        btn_save.pack(side="right", padx=5)

        btn_copy = ttk.Button(frame_buttons, text="📋 Sao chép JSON", command=self.copy_json_to_clipboard)
        btn_copy.pack(side="right", padx=5)

        # 2. Khung Notebook (Tabs kết quả)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # Tab 1: Tóm tắt thông tin quan trọng
        tab_summary = ttk.Frame(notebook, padding=10)
        notebook.add(tab_summary, text=" 📋 Tóm Tắt & Kết Quả Chuyển Đổi ")

        self.txt_summary = tk.Text(tab_summary, wrap="word", font=("Consolas", 10), padx=8, pady=8)
        self.txt_summary.pack(fill="both", expand=True, side="left")
        scroll_sum = ttk.Scrollbar(tab_summary, command=self.txt_summary.yview)
        scroll_sum.pack(side="right", fill="y")
        self.txt_summary.config(yscrollcommand=scroll_sum.set)

        # Tab 2: JSON sau khi chuyển đổi
        tab_json = ttk.Frame(notebook, padding=10)
        notebook.add(tab_json, text=" 📜 JSON Hoàn Chỉnh Đã Chuyển Đổi ")

        self.txt_json = tk.Text(tab_json, wrap="none", font=("Consolas", 9), padx=8, pady=8)
        self.txt_json.pack(fill="both", expand=True, side="left")

        scroll_json_y = ttk.Scrollbar(tab_json, command=self.txt_json.yview)
        scroll_json_y.pack(side="right", fill="y")
        scroll_json_x = ttk.Scrollbar(tab_json, orient="horizontal", command=self.txt_json.xview)
        scroll_json_x.pack(side="bottom", fill="x")
        self.txt_json.config(yscrollcommand=scroll_json_y.set, xscrollcommand=scroll_json_x.set)

        # 3. Thanh trạng thái dưới cùng
        self.status_bar = ttk.Label(self.root, text="Sẵn sàng. Nhấn '🧪 1. Chạy Dữ Liệu Mẫu' để kiểm tra kết quả ngay.", relief="sunken", anchor="w", padding=4)
        self.status_bar.pack(fill="x", side="bottom")

    def _set_status(self, msg: str):
        self.status_bar.config(text=msg)

    def load_sample_data(self):
        """Tải dữ liệu mẫu từ sample_response_nhom1.json trong thư mục hiện tại."""
        sample_path = Path(__file__).parent / "sample_response_nhom1.json"
        if not sample_path.exists():
            messagebox.showerror("Lỗi", f"Không tìm thấy file mẫu: {sample_path}")
            return

        try:
            with open(sample_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.current_data = data
            self._process_and_display_data(f"Dữ liệu mẫu từ {sample_path.name}")
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", str(e))

    def open_json_file(self):
        """Mở file JSON bất kỳ từ đĩa."""
        file_path = filedialog.askopenfilename(
            title="Chọn file response JSON",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.current_data = data
            self._process_and_display_data(f"Đọc từ file: {Path(file_path).name}")
        except Exception as e:
            messagebox.showerror("Lỗi mở file", f"Không thể đọc JSON: {e}")

    def call_api(self):
        """Gửi request tới API thật của MPLIS."""
        id_str = self.entry_id.get().strip()
        if not id_str.isdigit():
            messagebox.showwarning("Thiếu ID", "Vui lòng nhập tinhHinhDangKyId hợp lệ (chỉ gồm số)!")
            return

        cookie_str = self.entry_cookie.get().strip()
        if not cookie_str:
            resp = messagebox.askyesno(
                "Chưa nhập Cookie",
                "Bạn chưa nhập Cookie đăng nhập MPLIS. Request có thể bị lỗi 401/403/Redirect nếu không có cookie.\n\nBạn vẫn muốn tiếp tục gửi request?",
            )
            if not resp:
                return

        self._set_status(f"Đang gửi request GetThongTinDangKyNhom1 với ID={id_str}...")
        self.root.update_idletasks()

        res = call_api_get_thong_tin_nhom1(int(id_str), cookie_str=cookie_str)
        if not res.get("success", False) and "error" in res:
            messagebox.showerror("Lỗi API", f"{res.get('error')}")
            self._set_status(f"Lỗi gọi API: {res.get('error')[:100]}")
            return

        self.current_data = res
        self._process_and_display_data(f"Kết quả API MPLIS cho ID {id_str}")

    def _on_toggle_convert_all(self):
        if self.current_data:
            self._process_and_display_data("Cập nhật tùy chọn chuyển ngày")

    def _process_and_display_data(self, source_desc: str):
        if not self.current_data:
            return

        convert_all = self.var_convert_all.get()
        # Chuyển đổi dữ liệu
        self.converted_data = convert_ngay_vao_so(copy.deepcopy(self.current_data), convert_all_dates=convert_all)

        # Trích xuất tóm tắt
        summary = extract_summary(self.current_data)

        # Hiển thị Tóm Tắt
        sum_text = []
        sum_text.append("=" * 75)
        sum_text.append(f"  KẾT QUẢ TRA CỨU & CHUYỂN ĐỔI NGÀY VÀO SỔ ({source_desc})")
        sum_text.append("=" * 75)
        sum_text.append("")
        sum_text.append(f"📌 Tình hình đăng ký ID : {summary.get('tinhHinhDangKyId')}")
        sum_text.append(f"📌 Thời điểm đăng ký   : {summary.get('thoiDiemDangKy')}")
        sum_text.append("")
        sum_text.append("-" * 75)
        sum_text.append("  THÔNG TIN GIẤY CHỨNG NHẬN (GCN):")
        sum_text.append("-" * 75)
        sum_text.append(f"  • Số phát hành GCN   : {summary.get('soPhatHanh')}")
        sum_text.append(f"  • Số vào sổ          : {summary.get('soVaoSo')}")
        sum_text.append(f"  • Ngày vào sổ (GỐC)  : {summary.get('ngayVaoSo_goc')}")
        sum_text.append(f"  ⭐ NGÀY VÀO SỔ (DMY)  : {summary.get('ngayVaoSo_dmy')}   <--- ĐÃ CHUYỂN ĐỔI DD/MM/YYYY")
        sum_text.append("")
        sum_text.append("-" * 75)
        sum_text.append("  THÔNG TIN CHỦ SỞ HỮU:")
        sum_text.append("-" * 75)
        for i, c in enumerate(summary.get("chuSoHuu") or [], 1):
            if c.get("loai") == "Vợ chồng":
                sum_text.append(f"  {i}. [Vợ chồng]:")
                sum_text.append(f"     - Chồng: {c.get('chong')}")
                sum_text.append(f"     - Vợ   : {c.get('vo')}")
                sum_text.append(f"     - Địa chỉ: {c.get('diaChi')}")
            else:
                sum_text.append(f"  {i}. [{c.get('loai')}]: {c.get('hoTen')} ({c.get('namSinh')}) - CCCD/CMND: {c.get('cccd')}")
                sum_text.append(f"     - Địa chỉ: {c.get('diaChi')}")

        sum_text.append("")
        sum_text.append("-" * 75)
        sum_text.append("  THÔNG TIN THỬA ĐẤT:")
        sum_text.append("-" * 75)
        for i, td in enumerate(summary.get("thuaDat") or [], 1):
            sum_text.append(f"  {i}. Thửa {td.get('soThua')}, Tờ bản đồ {td.get('soTo')}, Diện tích: {td.get('dienTich')} m²")
            sum_text.append(f"     - Mục đích sử dụng : {td.get('mucDichSuDung')}")
            sum_text.append(f"     - Địa chỉ          : {td.get('diaChi')}")

        sum_text.append("")
        sum_text.append("-" * 75)
        sum_text.append("  HỒ SƠ QUÉT ĐÍNH KÈM:")
        sum_text.append("-" * 75)
        for i, hsq in enumerate(summary.get("hoSoQuet") or [], 1):
            sum_text.append(f"  {i}. Hồ sơ quét ID: {hsq.get('hoSoQuetId')}")
            for f in hsq.get("files") or []:
                sum_text.append(f"     + {f}")

        sum_text.append("\n" + "=" * 75)

        self.txt_summary.delete("1.0", "end")
        self.txt_summary.insert("1.0", "\n".join(sum_text))

        # Hiển thị JSON
        json_str = json.dumps(self.converted_data, ensure_ascii=False, indent=2)
        self.txt_json.delete("1.0", "end")
        self.txt_json.insert("1.0", json_str)

        self._set_status(f"Đã xử lý thành công {source_desc}! 'ngayVaoSo' đã được chuyển thành: {summary.get('ngayVaoSo_dmy')}")

    def save_converted_json(self):
        if not self.converted_data:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng tải hoặc lấy dữ liệu trước khi lưu!")
            return

        file_path = filedialog.asksaveasfilename(
            title="Lưu file JSON đã chuyển đổi",
            defaultextension=".json",
            initialfile="ket_qua_nhom1_converted.json",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.converted_data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Thành công", f"Đã lưu kết quả thành công vào:\n{file_path}")
            self._set_status(f"Đã lưu: {file_path}")
        except Exception as e:
            messagebox.showerror("Lỗi lưu file", str(e))

    def copy_json_to_clipboard(self):
        if not self.converted_data:
            messagebox.showwarning("Chưa có dữ liệu", "Chưa có dữ liệu JSON để sao chép!")
            return

        json_str = json.dumps(self.converted_data, ensure_ascii=False, indent=2)
        self.root.clipboard_clear()
        self.root.clipboard_append(json_str)
        messagebox.showinfo("Đã sao chép", "Đã sao chép toàn bộ nội dung JSON đã chuyển đổi vào Clipboard!")


# ============================ XỬ LÝ CLI ============================

def run_cli(args: argparse.Namespace):
    """Chạy chế độ dòng lệnh."""
    data = None
    source_name = ""

    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"❌ Không tìm thấy file: {file_path}")
            sys.exit(1)
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        source_name = f"File {file_path.name}"
    elif args.sample:
        sample_path = Path(__file__).parent / "sample_response_nhom1.json"
        if not sample_path.exists():
            print(f"❌ Không tìm thấy file mẫu: {sample_path}")
            sys.exit(1)
        with open(sample_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        source_name = "Dữ liệu mẫu (sample_response_nhom1.json)"
    elif args.id:
        print(f"🌐 Đang gọi API GetThongTinDangKyNhom1 với ID={args.id}...")
        res = call_api_get_thong_tin_nhom1(args.id, cookie_str=args.cookie or "")
        if not res.get("success", False) and "error" in res:
            print(f"❌ Lỗi API: {res.get('error')}")
            sys.exit(1)
        data = res
        source_name = f"API MPLIS (ID={args.id})"
    else:
        print("💡 Không có tham số đầu vào. Sử dụng --sample để kiểm tra với dữ liệu mẫu.")
        sample_path = Path(__file__).parent / "sample_response_nhom1.json"
        if sample_path.exists():
            with open(sample_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            source_name = "Dữ liệu mẫu mặc định"
        else:
            print("❌ Vui lòng truyền --file, --sample hoặc --id")
            sys.exit(1)

    # Chuyển đổi ngày
    converted_data = convert_ngay_vao_so(copy.deepcopy(data), convert_all_dates=args.convert_all)
    summary = extract_summary(data)

    print("\n" + "=" * 70)
    print(f"  KẾT QUẢ TRA CỨU & CHUYỂN ĐỔI NGÀY VÀO SỔ ({source_name})")
    print("=" * 70)
    print(f"  • ID Tình hình đăng ký : {summary.get('tinhHinhDangKyId')}")
    print(f"  • Số phát hành GCN     : {summary.get('soPhatHanh')}")
    print(f"  • Số vào sổ            : {summary.get('soVaoSo')}")
    print(f"  • Ngày vào sổ (GỐC)    : {summary.get('ngayVaoSo_goc')}")
    print(f"  ⭐ NGÀY VÀO SỔ (DMY)    : {summary.get('ngayVaoSo_dmy')}   <--- ĐÃ CHUYỂN THÀNH DD/MM/YYYY")
    print("-" * 70)
    for c in summary.get("chuSoHuu") or []:
        if c.get("loai") == "Vợ chồng":
            print(f"  • Chủ sở hữu: [Vợ chồng] {c.get('chong')} & {c.get('vo')}")
        else:
            print(f"  • Chủ sở hữu: {c.get('hoTen')} ({c.get('namSinh')}) - CCCD: {c.get('cccd')}")
    for td in summary.get("thuaDat") or []:
        print(f"  • Thửa đất: Thửa {td.get('soThua')}, Tờ {td.get('soTo')}, DT {td.get('dienTich')}m², {td.get('mucDichSuDung')}")
    print("=" * 70)

    # Lưu file nếu chỉ định hoặc mặc định
    out_file = args.out or "ket_qua_nhom1_converted.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(converted_data, f, ensure_ascii=False, indent=2)
    print(f"✅ Đã lưu file JSON đã chuyển đổi vào: {out_file}\n")


# ============================ MAIN ENTRY ============================

def main():
    parser = argparse.ArgumentParser(description="Tool tra cứu GetThongTinDangKyNhom1 và chuyển đổi ngayVaoSo sang DD/MM/YYYY")
    parser.add_argument("--id", type=int, help="tinhHinhDangKyId cần tra cứu từ API")
    parser.add_argument("--cookie", type=str, default="", help="Cookie đăng nhập MPLIS")
    parser.add_argument("--file", type=str, help="Đọc và chuyển đổi từ file JSON")
    parser.add_argument("--sample", action="store_true", help="Chạy thử với dữ liệu mẫu")
    parser.add_argument("--convert-all", action="store_true", help="Chuyển đổi tất cả trường /Date(...)/ trong JSON")
    parser.add_argument("--out", type=str, help="Đường dẫn file lưu kết quả JSON")
    parser.add_argument("--cli", action="store_true", help="Bắt buộc chạy chế độ dòng lệnh (CLI)")

    args = parser.parse_args()

    # Nếu có tham số CLI hoặc không có GUI Tkinter
    if args.cli or args.id or args.file or args.sample or not HAS_TKINTER:
        run_cli(args)
    else:
        root = tk.Tk()
        app = ToolNhom1GUI(root)
        # Tự động load dữ liệu mẫu khi mở giao diện
        app.load_sample_data()
        root.mainloop()


if __name__ == "__main__":
    main()
