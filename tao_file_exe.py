import os
import re
import sys
import queue
import shutil
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


DEFAULT_SPEC_TEMPLATE = """# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    [{script!r}],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name={name!r},
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console={console},
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[{icon!r}],
)
"""


def tao_spec_moi(script_path, icon_path, exe_name, console):
    return DEFAULT_SPEC_TEMPLATE.format(
        script=os.path.abspath(script_path),
        name=exe_name,
        console=bool(console),
        icon=os.path.abspath(icon_path),
    )


def va_spec_co_san(spec_text, script_path, icon_path, exe_name):
    """Giữ nguyên toàn bộ nội dung spec, chỉ thay script/icon/name."""
    script_repr = repr(os.path.abspath(script_path))
    icon_repr   = repr(os.path.abspath(icon_path))
    name_repr   = repr(exe_name)

    spec_text, n = re.subn(
        r"(Analysis\(\s*\[\s*)['\"][^'\"]*['\"](\s*\])",
        lambda m: m.group(1) + script_repr + m.group(2),
        spec_text,
        count=1,
    )
    if n == 0:
        raise ValueError("Không tìm thấy Analysis([...]) trong file spec để thay script .py.")

    if re.search(r"\bicon\s*=", spec_text):
        spec_text = re.sub(
            r"icon\s*=\s*\[[^\]]*\]",
            lambda m: f"icon=[{icon_repr}]",
            spec_text,
            count=1,
        )
    else:
        stripped = spec_text.rstrip()
        if not stripped.endswith(")"):
            raise ValueError("File spec không kết thúc bằng ')' — không tự thêm icon được, cần thêm icon=[...] thủ công.")
        body = stripped[:-1].rstrip()
        if not body.endswith(","):
            body += ","
        spec_text = body + f"\n    icon=[{icon_repr}],\n)\n"

    if re.search(r"\bname\s*=\s*['\"]", spec_text):
        spec_text = re.sub(
            r"name\s*=\s*['\"][^'\"]*['\"]",
            lambda m: f"name={name_repr}",
            spec_text,
            count=1,
        )

    return spec_text


def build_exe(py_path, icon_path, spec_path, exe_name, console, log_queue):
    def log(msg):
        log_queue.put(str(msg))

    try:
        if not os.path.isfile(py_path):
            log("❌ Không tìm thấy file .py.")
            return
        if not os.path.isfile(icon_path):
            log("❌ Không tìm thấy file icon.")
            return

        work_dir = os.path.dirname(os.path.abspath(py_path)) or "."
        out_spec_path = os.path.join(work_dir, f"{exe_name}.spec")

        if spec_path:
            if not os.path.isfile(spec_path):
                log("❌ Không tìm thấy file .spec đã chọn.")
                return
            with open(spec_path, "r", encoding="utf-8") as f:
                spec_text = f.read()
            spec_text = va_spec_co_san(spec_text, py_path, icon_path, exe_name)
            log(f"📄 Dùng spec có sẵn: {spec_path}")
        else:
            spec_text = tao_spec_moi(py_path, icon_path, exe_name, console)
            log("📄 Tạo spec mặc định (không có hiddenimports/datas tuỳ biến).")

        with open(out_spec_path, "w", encoding="utf-8") as f:
            f.write(spec_text)
        log(f"✅ Đã ghi spec: {out_spec_path}")

        cmd = [sys.executable, "-m", "PyInstaller", out_spec_path, "--noconfirm"]
        log("▶ Chạy: " + " ".join(cmd))

        proc = subprocess.Popen(
            cmd,
            cwd=work_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        for line in proc.stdout:
            log(line.rstrip("\n"))
        proc.wait()

        if proc.returncode != 0:
            log(f"❌ Build lỗi (exit code {proc.returncode}).")
            return

        exe_path = os.path.join(work_dir, "dist", exe_name + ".exe")
        if os.path.isfile(exe_path):
            log("🎯 XONG. File exe: " + exe_path)
            log_queue.put({"type": "done_ok", "exe_path": exe_path})
        else:
            log("⚠️ Build chạy xong nhưng không thấy file exe ở: " + exe_path)

    except Exception as e:
        log("❌ Lỗi: " + str(e))

    finally:
        log_queue.put("__DONE__")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tạo file EXE kèm Icon")
        self.geometry("820x560")

        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.last_exe_path = None

        self.var_py     = tk.StringVar()
        self.var_icon   = tk.StringVar()
        self.var_spec   = tk.StringVar()
        self.var_name   = tk.StringVar()
        self.var_console = tk.BooleanVar(value=True)

        self.create_widgets()
        self.after(200, self.process_log_queue)

    def create_widgets(self):
        frame = ttk.LabelFrame(self, text="Thông tin build")
        frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame, text="File .py").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame, textvariable=self.var_py, width=70).grid(row=0, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(frame, text="Duyệt", command=self.browse_py).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame, text="File icon (.ico)").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame, textvariable=self.var_icon, width=70).grid(row=1, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(frame, text="Duyệt", command=self.browse_icon).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(frame, text="File .spec (tuỳ chọn)").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame, textvariable=self.var_spec, width=70).grid(row=2, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(frame, text="Duyệt", command=self.browse_spec).grid(row=2, column=2, padx=5, pady=5)
        ttk.Label(frame, text="Để trống = tự tạo spec mặc định", foreground="gray").grid(
            row=3, column=1, padx=5, pady=0, sticky="w"
        )

        ttk.Label(frame, text="Tên file exe").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame, textvariable=self.var_name, width=30).grid(row=4, column=1, padx=5, pady=5, sticky="w")

        ttk.Checkbutton(
            frame, text="Hiện console (chỉ áp dụng khi tự tạo spec mặc định)",
            variable=self.var_console
        ).grid(row=5, column=1, padx=5, pady=5, sticky="w")

        self.btn_build = ttk.Button(frame, text="▶  BUILD EXE", command=self.start_build)
        self.btn_build.grid(row=6, column=1, padx=5, pady=10, sticky="w")

        self.btn_open_dist = ttk.Button(frame, text="Mở thư mục dist", command=self.open_dist, state="disabled")
        self.btn_open_dist.grid(row=6, column=2, padx=5, pady=10, sticky="w")

        frame.columnconfigure(1, weight=1)

        frame_log = ttk.LabelFrame(self, text="Log build")
        frame_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.txt_log = tk.Text(frame_log, wrap="word", font=("Consolas", 9))
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(frame_log, command=self.txt_log.yview)
        scrollbar.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=scrollbar.set)

    def browse_py(self):
        path = filedialog.askopenfilename(title="Chọn file .py", filetypes=[("Python files", "*.py")])
        if path:
            self.var_py.set(path)
            if not self.var_name.get().strip():
                self.var_name.set(os.path.splitext(os.path.basename(path))[0])

    def browse_icon(self):
        path = filedialog.askopenfilename(title="Chọn file icon", filetypes=[("Icon files", "*.ico")])
        if path:
            self.var_icon.set(path)

    def browse_spec(self):
        path = filedialog.askopenfilename(title="Chọn file .spec", filetypes=[("Spec files", "*.spec"), ("All files", "*.*")])
        if path:
            self.var_spec.set(path)

    def open_dist(self):
        if self.last_exe_path and os.path.isfile(self.last_exe_path):
            os.startfile(os.path.dirname(self.last_exe_path))

    def start_build(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Đang chạy", "Đang build, vui lòng đợi.")
            return

        py_path   = self.var_py.get().strip()
        icon_path = self.var_icon.get().strip()
        spec_path = self.var_spec.get().strip() or None
        exe_name  = self.var_name.get().strip()

        if not os.path.isfile(py_path):
            messagebox.showerror("Thiếu thông tin", "Chưa chọn file .py hợp lệ.")
            return
        if not os.path.isfile(icon_path):
            messagebox.showerror("Thiếu thông tin", "Chưa chọn file icon hợp lệ.")
            return
        if not exe_name:
            messagebox.showerror("Thiếu thông tin", "Chưa nhập tên file exe.")
            return

        self.txt_log.delete("1.0", tk.END)
        self.btn_build.config(state="disabled")
        self.btn_open_dist.config(state="disabled")
        self.last_exe_path = None

        self.worker_thread = threading.Thread(
            target=build_exe,
            args=(py_path, icon_path, spec_path, exe_name, self.var_console.get(), self.log_queue),
            daemon=True
        )
        self.worker_thread.start()

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()

                if isinstance(msg, dict) and msg.get("type") == "done_ok":
                    self.last_exe_path = msg["exe_path"]
                    self.btn_open_dist.config(state="normal")
                    continue

                if msg == "__DONE__":
                    self.btn_build.config(state="normal")
                    continue

                self.txt_log.insert(tk.END, msg + "\n")
                self.txt_log.see(tk.END)

        except queue.Empty:
            pass

        self.after(200, self.process_log_queue)


if __name__ == "__main__":
    app = App()
    app.mainloop()
