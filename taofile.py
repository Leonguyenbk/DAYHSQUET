import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

def chon_file():
    path = filedialog.askopenfilename(
        title="Chọn file gốc",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, path)

def tao_file():
    file_goc = entry_file.get().strip()
    so_luong_text = entry_so_luong.get().strip()

    if not file_goc or not os.path.isfile(file_goc):
        messagebox.showerror("Lỗi", "Chưa chọn file gốc hợp lệ.")
        return

    if not so_luong_text.isdigit() or int(so_luong_text) < 1:
        messagebox.showerror("Lỗi", "Số lượng phải là số nguyên dương.")
        return

    so_luong = int(so_luong_text)

    folder = os.path.dirname(file_goc)
    ten_file = os.path.basename(file_goc)
    name, ext = os.path.splitext(ten_file)

    if "-GT" not in name:
        messagebox.showerror("Lỗi", "Tên file phải có dạng ..._1-GT.pdf")
        return

    try:
        phan_truoc_gt = name.replace("-GT", "")
        prefix = phan_truoc_gt.rsplit("_", 1)[0]
        so_bat_dau = int(phan_truoc_gt.rsplit("_", 1)[1])
    except:
        messagebox.showerror("Lỗi", "Không lấy được số thứ tự trong tên file.")
        return

    dem_tao = 0
    dem_bo_qua = 0

    for i in range(so_bat_dau + 1, so_luong + 1):
        ten_moi = f"{prefix}_{i}-GT{ext}"
        path_moi = os.path.join(folder, ten_moi)

        if os.path.exists(path_moi):
            dem_bo_qua += 1
            continue

        shutil.copy2(file_goc, path_moi)
        dem_tao += 1

    messagebox.showinfo(
        "Hoàn thành",
        f"Đã tạo: {dem_tao} file\nBỏ qua do đã tồn tại: {dem_bo_qua} file"
    )

root = tk.Tk()
root.title("Copy file tịnh tiến")
root.geometry("620x180")

tk.Label(root, text="File gốc:").pack(anchor="w", padx=10, pady=(10, 0))

frame_file = tk.Frame(root)
frame_file.pack(fill="x", padx=10)

entry_file = tk.Entry(frame_file)
entry_file.pack(side="left", fill="x", expand=True)

tk.Button(frame_file, text="Chọn file", command=chon_file).pack(side="left", padx=5)

tk.Label(root, text="Số lượng file cần có:").pack(anchor="w", padx=10, pady=(10, 0))

entry_so_luong = tk.Entry(root)
entry_so_luong.pack(fill="x", padx=10)
entry_so_luong.insert(0, "49")

tk.Button(root, text="Tạo file", command=tao_file, height=2).pack(pady=15)

root.mainloop()