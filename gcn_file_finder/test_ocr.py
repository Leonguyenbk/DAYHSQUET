"""
Công cụ test kết quả OCR từ pipeline của locfilescan_2.py (Có hỗ trợ xử lý nền đỏ thành trắng).

Mục đích:
- Kiểm tra sau khi Tesseract OCR đọc xong trang PDF / ảnh thì nó trả ra chuỗi text gì (Raw OCR).
- Kiểm tra kết quả sau khi xóa dấu cách và viết hoa (Clean text).
- Giải quyết vấn đề chữ/mã đen nằm trên nền đỏ/hồng bị biến mất bằng kỹ thuật:
  + Tách kênh Đỏ (Red Channel): Toàn bộ nền đỏ/hồng và dấu mộc đỏ biến thành TRẮNG TINH, chữ đen giữ nguyên ĐEN TUYỀN.
  + Hoặc nhận diện pixel đỏ/hồng và tô trắng (Whiten Red Background).
- So sánh các phương pháp xử lý ảnh để chọn phương pháp tối ưu nhất.

Cách dùng:
1. Giao diện GUI:
   .venv\\Scripts\\python.exe test_ocr.py
2. Dòng lệnh CLI:
   .venv\\Scripts\\python.exe test_ocr.py "D:\\duong_dan_file.pdf" --method red_channel
"""

import os
import sys
import re
import argparse

# Thiết lập UTF-8 cho terminal Windows tránh lỗi UnicodeEncodeError
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np
import fitz  # PyMuPDF
import pytesseract

# Cấu hình đường dẫn Tesseract OCR (giống locfilescan_2.py)
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(TESSERACT_EXE):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
else:
    TESSERACT_EXE_X86 = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    if os.path.exists(TESSERACT_EXE_X86):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE_X86

# ===== Các hàm trích xuất text =====

def extract_so_hieu(text):
    text_clean = text.replace(" ", "").upper()
    matches = re.findall(r'[A-ZĐ]{1,2}\d{6}', text_clean)
    return matches[0] if matches else None

def extract_so_hieu_full(text):
    text_clean = text.replace(" ", "").upper()
    matches = re.findall(r'[A-ZĐ]{1,2}\d{6}\b', text_clean)
    return matches

def extract_so_hieu_no_boundary(text):
    text_clean = text.replace(" ", "").upper()
    matches = re.findall(r'[A-ZĐ]{1,2}\d{6}', text_clean)
    return matches

def co_quoc_hieu(text):
    return "QUYEN SU DUNG DAT" in text.upper()

def kiem_tra_quoc_hieu_chi_tiet(text):
    """Kiểm tra các biến thể của Quốc hiệu / Tiêu đề GCN"""
    text_upper = text.upper()
    ket_qua = {
        "QUYEN SU DUNG DAT (Không dấu - locfilescan_2)": "QUYEN SU DUNG DAT" in text_upper,
        "QUYỀN SỬ DỤNG ĐẤT (Có dấu)": "QUYỀN SỬ DỤNG ĐẤT" in text_upper,
        "GIẤY CHỨNG NHẬN / GIAY CHUNG NHAN": ("GIẤY CHỨNG NHẬN" in text_upper) or ("GIAY CHUNG NHAN" in text_upper),
        "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM": ("CỘNG HÒA XÃ HỘI" in text_upper) or ("CONG HOA XA HOI" in text_upper),
    }
    return ket_qua


# ===== Các phương pháp xử lý ảnh =====

def xu_ly_tuong_phan_cu(img):
    """Phương pháp cũ trong locfilescan_2.py (dễ làm đen nền đỏ)"""
    img_gray = img.convert("L")
    brightness_enhancer = ImageEnhance.Brightness(img_gray)
    img_bright = brightness_enhancer.enhance(1.0)
    contrast_enhancer = ImageEnhance.Contrast(img_bright)
    img_contrast = contrast_enhancer.enhance(1.6)
    return img_contrast.filter(ImageFilter.SHARPEN)

def xu_ly_tach_kenh_do(img):
    """
    [KHUYÊN DÙNG] Tách kênh Đỏ (Red Channel):
    - Màu đỏ/hồng có chỉ số R rất cao (~200-255) -> Tự động thành MÀU TRẮNG TINH.
    - Dấu mộc đỏ -> Tự động biến mất hoặc thành màu trắng, không đè lên chữ.
    - Chữ đen có chỉ số R rất thấp (~0-40) -> Giữ nguyên MÀU ĐEN TUYỀN.
    - Độ tương phản đạt cực đại tuyệt đối mà không bị lem hay vỡ nét chữ.
    """
    if img.mode != "RGB":
        img = img.convert("RGB")
    # Tách kênh R (kênh 0 trong RGB)
    r_channel = img.split()[0]
    # Tăng tương phản vừa phải để chữ sắc nét
    contrast = ImageEnhance.Contrast(r_channel)
    img_contrast = contrast.enhance(1.3)
    return img_contrast.filter(ImageFilter.SHARPEN)

def xu_ly_to_trang_mau_do(img):
    """
    Tô trắng các pixel màu đỏ/hồng (Whiten Red Background):
    Tìm các pixel có màu đỏ (R > G và R > B) và gán thành trắng (255, 255, 255).
    """
    if img.mode != "RGB":
        img = img.convert("RGB")
    arr = np.array(img, dtype=np.int32)
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]

    # Điều kiện màu đỏ/hồng: R lớn hơn hẳn G và B, đồng thời R > 70 (để không nhầm chữ đen)
    is_red = (r > 70) & (r > g + 20) & (r > b + 20)

    arr_out = np.array(img, dtype=np.uint8)
    arr_out[is_red] = [255, 255, 255]

    img_whitened = Image.fromarray(arr_out).convert("L")
    contrast = ImageEnhance.Contrast(img_whitened)
    img_contrast = contrast.enhance(1.3)
    return img_contrast.filter(ImageFilter.SHARPEN)

def ap_dung_tien_xu_ly(img, method="red_channel"):
    """Áp dụng phương pháp xử lý tương ứng"""
    if method == "red_channel":
        return xu_ly_tach_kenh_do(img)
    elif method == "whiten_red":
        return xu_ly_to_trang_mau_do(img)
    elif method == "legacy":
        return xu_ly_tuong_phan_cu(img)
    elif method == "raw":
        return img.convert("L")
    else:
        return xu_ly_tach_kenh_do(img)


# ===== Hàm xử lý OCR & Phân tích =====

def phan_tich_trang_ocr(img, page_num=1, file_name="", method="red_channel", lang="vie+eng", debug_dir=None):
    ket_qua = {
        "page_num": page_num,
        "img_size": img.size,
        "method": method,
        "raw_text": "",
        "clean_text": "",
        "co_quoc_hieu": False,
        "quoc_hieu_chi_tiet": {},
        "so_hieu_full": [],
        "so_hieu_no_boundary": [],
        "so_hieu_tiem_nang": [],
        "debug_img_path": None
    }

    img_xu_ly = ap_dung_tien_xu_ly(img, method=method)

    # Lưu ảnh debug nếu yêu cầu
    if debug_dir:
        os.makedirs(debug_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(file_name))[0]
        img_filename = f"{base}_p{page_num}_{method}.png"
        save_path = os.path.join(debug_dir, img_filename)
        img_xu_ly.save(save_path)
        ket_qua["debug_img_path"] = save_path

    # Chạy Tesseract OCR
    raw_text = pytesseract.image_to_string(img_xu_ly, lang=lang)
    clean_text = raw_text.replace(" ", "").upper()

    ket_qua["raw_text"] = raw_text
    ket_qua["clean_text"] = clean_text
    ket_qua["co_quoc_hieu"] = co_quoc_hieu(raw_text)
    ket_qua["quoc_hieu_chi_tiet"] = kiem_tra_quoc_hieu_chi_tiet(raw_text)
    ket_qua["so_hieu_full"] = extract_so_hieu_full(raw_text)
    ket_qua["so_hieu_no_boundary"] = extract_so_hieu_no_boundary(raw_text)

    # Tìm các mã tiềm năng (gần giống số hiệu: 1-2 chữ cái + 5-7 chữ số)
    tiem_nang = re.findall(r'[A-ZĐ]{1,2}\d{5,7}', clean_text)
    ket_qua["so_hieu_tiem_nang"] = list(set(tiem_nang) - set(ket_qua["so_hieu_no_boundary"]))

    return ket_qua

def doc_va_ocr_file(file_path, dpi=150, method="red_channel", lang="vie+eng", debug_dir=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    file_name = os.path.basename(file_path)
    ext = os.path.splitext(file_name)[1].lower()
    ket_qua_cac_trang = []

    if ext == ".pdf":
        doc = fitz.open(file_path)
        total_pages = len(doc)
        for idx, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            res = phan_tich_trang_ocr(
                img,
                page_num=idx,
                file_name=file_name,
                method=method,
                lang=lang,
                debug_dir=debug_dir
            )
            res["total_pages"] = total_pages
            ket_qua_cac_trang.append(res)
        doc.close()
    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"]:
        img = Image.open(file_path)
        res = phan_tich_trang_ocr(
            img,
            page_num=1,
            file_name=file_name,
            method=method,
            lang=lang,
            debug_dir=debug_dir
        )
        res["total_pages"] = 1
        ket_qua_cac_trang.append(res)
    else:
        raise ValueError(f"Định dạng file '{ext}' không được hỗ trợ.")

    return ket_qua_cac_trang

def tao_bao_cao_text(file_path, ket_qua_cac_trang, method_desc=""):
    lines = []
    lines.append("=" * 80)
    lines.append(f" BÁO CÁO KẾT QUẢ OCR CHO FILE: {os.path.basename(file_path)}")
    lines.append(f" Phương pháp xử lý ảnh: {method_desc}")
    lines.append(f" Đường dẫn file: {file_path}")
    lines.append(f" Tổng số trang: {len(ket_qua_cac_trang)}")
    lines.append("=" * 80)
    lines.append("")

    tat_ca_so_hieu = []

    for res in ket_qua_cac_trang:
        p = res["page_num"]
        lines.append(f"┌{'─' * 78}┐")
        lines.append(f"│ TRANG {p}/{res['total_pages']} (Kích thước: {res['img_size'][0]}x{res['img_size'][1]} px | Chế độ: {res['method']})")
        lines.append(f"└{'─' * 78}┘")

        # 1. Kiểm tra số hiệu
        matches = res["so_hieu_full"]
        matches_no_b = res["so_hieu_no_boundary"]
        tat_ca_so_hieu.extend(matches)

        lines.append(f"  🔍 [Số hiệu tìm thấy (extract_so_hieu_full)]: {matches if matches else 'KHÔNG TÌM THẤY'}")
        if matches != matches_no_b:
            lines.append(f"  ℹ️ [Số hiệu không kèm \\b]: {matches_no_b}")
        if res.get("so_hieu_tiem_nang"):
            lines.append(f"  💡 [Mã tiềm năng 5-7 số gần giống]: {res['so_hieu_tiem_nang']}")

        # 2. Kiểm tra quốc hiệu
        qh = res["co_quoc_hieu"]
        lines.append(f"  🟡 [Có 'QUYEN SU DUNG DAT' (co_quoc_hieu)]: {'CÓ ✅' if qh else 'KHÔNG ❌'}")
        for k, v in res["quoc_hieu_chi_tiet"].items():
            lines.append(f"     - {k}: {'CÓ' if v else 'Không'}")

        if res.get("debug_img_path"):
            lines.append(f"  🖼️ [Ảnh debug đã lưu tại]: {res['debug_img_path']}")

        lines.append("")
        lines.append("  ────────────────── [RAW OCR TEXT (Văn bản thô đọc được)] ──────")
        raw = res["raw_text"].strip()
        if raw:
            for line in raw.splitlines():
                lines.append(f"  | {line}")
        else:
            lines.append("  | (Trống - OCR không nhận diện được chữ nào)")
        lines.append("  ─────────────────────────────────────────────────────────────────")

        lines.append("")
        lines.append("  ────────────────── [CLEAN TEXT (Bỏ khoảng trắng, viết hoa)] ─────")
        clean = res["clean_text"].strip()
        if clean:
            preview = clean if len(clean) <= 500 else clean[:500] + " ...[CẮT BỚT]"
            lines.append(f"  {preview}")
        else:
            lines.append("  (Trống)")
        lines.append("  ─────────────────────────────────────────────────────────────────")
        lines.append("")

    lines.append("=" * 80)
    lines.append(f" TỔNG KẾT FILE: {os.path.basename(file_path)}")
    lines.append(f" Số hiệu tổng hợp: {tat_ca_so_hieu if tat_ca_so_hieu else 'KHÔNG TÌM THẤY'}")
    if tat_ca_so_hieu:
        ten_goc_ko_ext, ext = os.path.splitext(os.path.basename(file_path))
        ten_moi = f"{ten_goc_ko_ext}_{'_'.join(tat_ca_so_hieu)}{ext}"
        lines.append(f" ✅ Tên file nếu locfilescan đổi: {ten_moi}")
    else:
        lines.append(" ❌ locfilescan sẽ: Bỏ qua (không đổi tên vì không bắt được mã)")
    lines.append("=" * 80)

    return "\n".join(lines)


# ===== Giao diện Tkinter =====

PHUONG_PHAP_DICT = {
    "1. Tách kênh Đỏ (Red Channel - Biến nền đỏ thành trắng tinh) [KHUYÊN DÙNG]": "red_channel",
    "2. Tô trắng pixel màu đỏ/hồng (Whiten Red Background)": "whiten_red",
    "3. Gốc cũ (locfilescan_2 cũ: Grayscale + Tương phản 1.6 - Bị đen nền đỏ)": "legacy",
    "4. Ảnh gốc (Không xử lý, chỉ chuyển ảnh xám)": "raw"
}

class OCRTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Công cụ Test OCR - Pipeline locfilescan (Khử nền đỏ)")
        self.root.geometry("900x750")
        self.root.minsize(750, 550)

        self.file_path_var = tk.StringVar()
        self.dpi_var = tk.IntVar(value=150)
        self.method_var = tk.StringVar(value=list(PHUONG_PHAP_DICT.keys())[0])
        self.save_debug_img_var = tk.BooleanVar(value=True)
        self.lang_var = tk.StringVar(value="vie+eng")

        self.last_report = ""
        self._build_ui()

    def _build_ui(self):
        # Frame chọn file
        frame_top = tk.LabelFrame(self.root, text=" 1. Chọn file cần kiểm tra ", padx=10, pady=8)
        frame_top.pack(fill="x", padx=10, pady=5)

        entry_file = tk.Entry(frame_top, textvariable=self.file_path_var, font=("Segoe UI", 10))
        entry_file.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_browse = tk.Button(frame_top, text="Chọn file (PDF / Ảnh)...", command=self.chon_file, padx=8)
        btn_browse.pack(side="left")

        # Frame cấu hình
        frame_config = tk.LabelFrame(self.root, text=" 2. Cấu hình xử lý ảnh & OCR ", padx=10, pady=8)
        frame_config.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_config, text="Xử lý ảnh (Khử nền đỏ):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=4)
        cb_method = ttk.Combobox(frame_config, textvariable=self.method_var, values=list(PHUONG_PHAP_DICT.keys()), width=65, state="readonly")
        cb_method.grid(row=0, column=1, columnspan=3, sticky="w", padx=5, pady=4)

        tk.Label(frame_config, text="DPI render PDF:").grid(row=1, column=0, sticky="w", padx=5, pady=4)
        cb_dpi = ttk.Combobox(frame_config, textvariable=self.dpi_var, values=[100, 150, 200, 300], width=8, state="readonly")
        cb_dpi.grid(row=1, column=1, sticky="w", padx=5, pady=4)

        tk.Label(frame_config, text="Ngôn ngữ:").grid(row=1, column=2, sticky="w", padx=(20, 5), pady=4)
        cb_lang = ttk.Combobox(frame_config, textvariable=self.lang_var, values=["vie+eng", "vie", "eng"], width=10, state="readonly")
        cb_lang.grid(row=1, column=3, sticky="w", padx=5, pady=4)

        chk_debug_img = tk.Checkbutton(
            frame_config,
            text="Lưu ảnh sau khi xử lý vào thư mục 'debug_ocr_images' để xem trực tiếp ảnh kết quả",
            variable=self.save_debug_img_var
        )
        chk_debug_img.grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=4)

        # Frame điều khiển
        frame_actions = tk.Frame(self.root, padx=10, pady=5)
        frame_actions.pack(fill="x", padx=10)

        self.btn_run = tk.Button(
            frame_actions,
            text="🚀 BẮT ĐẦU CHẠY OCR VÀ XEM KẾT QUẢ",
            command=self.chay_ocr,
            bg="#007acc",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=6
        )
        self.btn_run.pack(side="left", padx=(0, 10))

        self.btn_save_txt = tk.Button(
            frame_actions,
            text="💾 Lưu kết quả ra file .txt",
            command=self.luu_file_txt,
            state="disabled",
            padx=10,
            pady=6
        )
        self.btn_save_txt.pack(side="left", padx=5)

        self.btn_clear = tk.Button(
            frame_actions,
            text="🗑️ Xóa màn hình",
            command=self.xoa_man_hinh,
            padx=10,
            pady=6
        )
        self.btn_clear.pack(side="right")

        # Frame kết quả
        frame_result = tk.LabelFrame(self.root, text=" 3. Kết quả chi tiết sau khi OCR ", padx=10, pady=8)
        frame_result.pack(fill="both", expand=True, padx=10, pady=5)

        self.txt_output = ScrolledText(frame_result, wrap="none", font=("Consolas", 10))
        self.txt_output.pack(fill="both", expand=True)

        # Thanh trạng thái dưới cùng
        self.lbl_status = tk.Label(self.root, text="Sẵn sàng.", bd=1, relief="sunken", anchor="w", padx=10)
        self.lbl_status.pack(fill="x", side="bottom")

    def chon_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file PDF hoặc ảnh để test OCR",
            filetypes=[
                ("File PDF & Ảnh", "*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp"),
                ("File PDF (*.pdf)", "*.pdf"),
                ("File ảnh (*.png;*.jpg;...)", "*.png;*.jpg;*.jpeg;*.bmp;*.tif"),
                ("Tất cả file (*.*)", "*.*")
            ]
        )
        if path:
            self.file_path_var.set(path)

    def set_status(self, text):
        self.lbl_status.config(text=text)
        self.root.update_idletasks()

    def xoa_man_hinh(self):
        self.txt_output.delete("1.0", tk.END)
        self.last_report = ""
        self.btn_save_txt.config(state="disabled")
        self.set_status("Đã xóa màn hình.")

    def chay_ocr(self):
        file_path = self.file_path_var.get().strip()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Lỗi", "Vui lòng chọn một file hợp lệ trước khi bấm Chạy!")
            return

        method_key = self.method_var.get()
        method_code = PHUONG_PHAP_DICT.get(method_key, "red_channel")

        self.btn_run.config(state="disabled")
        self.set_status(f"Đang xử lý OCR cho: {os.path.basename(file_path)} với phương pháp '{method_code}'...")
        self.txt_output.delete("1.0", tk.END)
        self.txt_output.insert(tk.END, f"⏳ Đang thực hiện OCR cho: {file_path}...\n")
        self.root.update()

        debug_dir = None
        if self.save_debug_img_var.get():
            debug_dir = os.path.join(os.path.dirname(file_path), "debug_ocr_images")

        try:
            dpi = self.dpi_var.get()
            lang = self.lang_var.get()

            ket_qua_cac_trang = doc_va_ocr_file(
                file_path=file_path,
                dpi=dpi,
                method=method_code,
                lang=lang,
                debug_dir=debug_dir
            )

            report = tao_bao_cao_text(file_path, ket_qua_cac_trang, method_desc=method_key)
            self.last_report = report

            self.txt_output.delete("1.0", tk.END)
            self.txt_output.insert(tk.END, report)
            self.btn_save_txt.config(state="normal")
            self.set_status(f"✅ Hoàn tất! Đã OCR {len(ket_qua_cac_trang)} trang bằng phương pháp '{method_code}'.")

            print("\n" + report)

        except Exception as e:
            err_msg = f"❌ Lỗi trong quá trình OCR: {str(e)}"
            self.txt_output.insert(tk.END, f"\n\n{err_msg}\n")
            messagebox.showerror("Lỗi", err_msg)
            self.set_status("Xảy ra lỗi!")
        finally:
            self.btn_run.config(state="normal")

    def luu_file_txt(self):
        if not self.last_report:
            return
        default_name = "ket_qua_ocr_test.txt"
        if self.file_path_var.get():
            base = os.path.splitext(os.path.basename(self.file_path_var.get()))[0]
            default_name = f"{base}_ocr_result.txt"

        save_path = filedialog.asksaveasfilename(
            title="Lưu kết quả OCR ra file TXT",
            initialfile=default_name,
            defaultextension=".txt",
            filetypes=[("Text file (*.txt)", "*.txt"), ("All files (*.*)", "*.*")]
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(self.last_report)
            messagebox.showinfo("Thành công", f"Đã lưu kết quả tại:\n{save_path}")


# ===== Entry point chính =====

def main():
    parser = argparse.ArgumentParser(description="Kiểm tra kết quả Tesseract OCR từ pipeline locfilescan (Khử nền đỏ)")
    parser.add_argument("file_path", nargs="?", help="Đường dẫn tới file PDF hoặc ảnh cần test")
    parser.add_argument("--dpi", type=int, default=150, help="DPI khi render PDF (mặc định: 150)")
    parser.add_argument(
        "--method",
        choices=["red_channel", "whiten_red", "legacy", "raw"],
        default="red_channel",
        help="Phương pháp xử lý: red_channel (mặc định, tối ưu nền đỏ), whiten_red, legacy, raw"
    )
    parser.add_argument("--lang", default="vie+eng", help="Ngôn ngữ Tesseract (mặc định: vie+eng)")
    parser.add_argument("--save-debug-img", action="store_true", help="Lưu ảnh debug sau xử lý tương phản")
    parser.add_argument("--out", help="Đường dẫn file .txt lưu kết quả")

    args = parser.parse_args()

    if args.file_path:
        file_path = os.path.abspath(args.file_path)
        print(f"📂 Đang mở file: {file_path}")
        debug_dir = os.path.join(os.path.dirname(file_path), "debug_ocr_images") if args.save_debug_img else None

        try:
            ket_qua = doc_va_ocr_file(
                file_path=file_path,
                dpi=args.dpi,
                method=args.method,
                lang=args.lang,
                debug_dir=debug_dir
            )
            report = tao_bao_cao_text(file_path, ket_qua, method_desc=args.method)
            print("\n" + report)

            out_path = args.out
            if not out_path:
                base = os.path.splitext(file_path)[0]
                out_path = f"{base}_ocr_test.txt"

            with open(out_path, "w", encoding="utf-8") as f:
                f.write(report)
            print(f"\n💾 Đã tự động lưu kết quả vào: {out_path}")

        except Exception as e:
            print(f"❌ Lỗi: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        root = tk.Tk()
        app = OCRTesterApp(root)
        root.mainloop()

if __name__ == "__main__":
    main()
