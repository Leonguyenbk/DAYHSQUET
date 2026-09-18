import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import os
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import re

# Đường dẫn tới Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ===== Các hàm xử lý =====

def extract_so_hieu(text):
    text_clean = text.replace(" ", "").upper()
    matches = re.findall(r'[A-ZĐ]{1,2}\d{6}', text_clean)
    return matches[0] if matches else None

def extract_so_hieu_full(text):
    text_clean = text.replace(" ", "").upper()
    matches = re.findall(r'[A-ZĐ]{1,2}\d{6}\b', text_clean)  # Tìm tất cả các chuỗi khớp
    return matches  # Trả về danh sách các chuỗi khớp

def co_quoc_hieu(text):
    return "QUYEN SU DUNG DAT" in text.upper()

def xu_ly_tuong_phan(img):
    # Chuyển sang grayscale
    img = img.convert("L")  

    # Tăng độ sáng nhẹ (nếu ảnh tối)
    brightness_enhancer = ImageEnhance.Brightness(img)
    img = brightness_enhancer.enhance(1.0)  # Tăng sáng 20%

    # Tăng độ tương phản (vừa phải)
    contrast_enhancer = ImageEnhance.Contrast(img)
    img = contrast_enhancer.enhance(1.6)  # Tăng tương phản 50%

    # Làm nét ảnh nhẹ (giúp viền chữ rõ hơn)
    img = img.filter(ImageFilter.SHARPEN)

    # Cân nhắc nhị phân hóa (binarization) nếu ảnh nền không đồng đều:
    # img = img.point(lambda x: 0 if x < 128 else 255, '1')  # Binarization

    return img

def doi_ten_file(file_path, so_hieu):
    folder, ten_goc = os.path.split(file_path)
    ten_goc_ko_ext, ext = os.path.splitext(ten_goc)
    ten_moi = f"{ten_goc_ko_ext}_{so_hieu}{ext}"
    duong_dan_moi = os.path.join(folder, ten_moi)

    if not os.path.exists(duong_dan_moi):
        os.rename(file_path, duong_dan_moi)
        print(f"✅ Đổi tên: {ten_moi}")
    else:
        print(f"⚠️ Đã tồn tại: {ten_moi}")

def doi_ten_file_full(file_path, so_hieu_list):
    folder, ten_goc = os.path.split(file_path)
    ten_goc_ko_ext, ext = os.path.splitext(ten_goc)
    # Nối các số hiệu bằng dấu _
    so_hieu = "_".join(so_hieu_list)
    ten_moi = f"{ten_goc_ko_ext}_{so_hieu}{ext}"
    duong_dan_moi = os.path.join(folder, ten_moi)

    if not os.path.exists(duong_dan_moi):
        os.rename(file_path, duong_dan_moi)
        print(f"✅ Đổi tên: {ten_moi}")
    else:
        print(f"⚠️ Đã tồn tại: {ten_moi}")

def doi_ten_file(file_path, so_hieu):

    print("⚠️ Hàm này đã bị loại bỏ. Vui lòng sử dụng doi_ten_file_full.")
    return

def xu_ly_pdf(file_path):
    try:
        doc = fitz.open(file_path)
        so_hieu = None
        so_hieu_list = []  # Danh sách để lưu tất cả số hiệu tìm thấy
        for page_num, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=150)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            #img = img.resize((int(img.width / 2), int(img.height / 2)), Image.Resampling.LANCZOS)
            img = xu_ly_tuong_phan(img)
            text = pytesseract.image_to_string(img, lang='vie+eng')

            if co_quoc_hieu(text):
                print(f"🟡 Quốc hiệu ở trang {page_num}: {file_path}")
            matches = extract_so_hieu_full(text)
            if matches:
                print(f"🔍 Các chuỗi khớp ở trang {page_num}: {matches}")
                so_hieu_list.extend(matches)  # Thêm các chuỗi khớp vào danh sách

        doc.close()

        if so_hieu_list:
            print(f"✅ Tìm số hiệu {so_hieu_list} → {file_path}")
            doi_ten_file_full(file_path, so_hieu_list)
            return so_hieu_list

        print(f"❌ Không tìm thấy số hiệu: {file_path}")
        return None
    except Exception as e:
        print(f"❌ Lỗi xử lý {file_path}: {e}")
        return None

# ===== Giao diện =====

def chon_folder():
    folder = filedialog.askdirectory()
    if folder:
        entry_folder.delete(0, tk.END)
        entry_folder.insert(0, folder)

def bat_dau_xu_ly():
    folder_path = entry_folder.get()
    if not os.path.isdir(folder_path):
        messagebox.showerror("Lỗi", "Thư mục không hợp lệ.")
        return

    # Kiểm tra nếu thư mục không chứa file PDF nào
    if not any(file.lower().endswith('.pdf') for _, _, files in os.walk(folder_path) for file in files):
        messagebox.showinfo("Thông báo", "Thư mục không chứa file PDF nào.")
        return

    tong = 0
    bo_qua = 0

    # Đếm tổng số file PDF để thiết lập thanh tiến trình
    total_files = sum(1 for _, _, files in os.walk(folder_path) for file in files if file.lower().endswith('.pdf'))
    progress["maximum"] = total_files
    progress["value"] = 0

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                # Chuẩn hóa tên file: bỏ tất cả ký tự không phải chữ hoặc số
                filename_clean = re.sub(r'\s+', '', file.strip().upper().replace("Ð", "Đ"))
                # Tách chuỗi bằng dấu _ hoặc -
                parts = re.split(r'[_-]', filename_clean)
                # Loại bỏ các phần dài hơn 15 ký tự
                filtered_parts = [part for part in parts if len(part) <= 15]
                # Ghép lại tên file sau khi loại bỏ
                filename_clean = '_'.join(filtered_parts)

                # Bỏ qua file đã có số hiệu hoặc đã đổi tên
                if re.search(r'[A-ZĐ]{1,2}\d{6}\b', filename_clean) or "_SOHIEU" in file.upper():
                    print(f"⏭️ Bỏ qua (đã có số hiệu hoặc đã đổi tên): {file}")
                    bo_qua += 1
                    continue

                file_path = os.path.join(root, file)
                print(f"\n📂 Đang xử lý: {file_path}")
                xu_ly_pdf(file_path)
                tong += 1

                # Cập nhật thanh tiến trình
                progress["value"] += 1
                root_window.update_idletasks()

    messagebox.showinfo("Hoàn tất", f"✅ Đã xử lý: {tong} file PDF\n⏭️ Bỏ qua: {bo_qua} file")
    os.startfile(folder_path)

# ===== Giao diện Tkinter =====

root_window = tk.Tk()
root_window.title("Đổi tên PDF theo số hiệu (tất cả thư mục con)")
root_window.geometry("550x250")

label_folder = tk.Label(root_window, text="Chọn thư mục chứa PDF:")
label_folder.pack(pady=5)

entry_folder = tk.Entry(root_window, width=60)
entry_folder.pack(pady=5)

btn_browse = tk.Button(root_window, text="Chọn thư mục", command=chon_folder)
btn_browse.pack(pady=5)

btn_run = tk.Button(root_window, text="Bắt đầu xử lý", command=bat_dau_xu_ly, bg="green", fg="white")
btn_run.pack(pady=10)

# Thêm thanh tiến trình
progress = Progressbar(root_window, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

root_window.mainloop()