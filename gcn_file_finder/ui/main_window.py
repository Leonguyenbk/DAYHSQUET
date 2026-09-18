"""Cửa sổ CustomTkinter tiếng Việt của ứng dụng."""

from __future__ import annotations

import logging
import os
import queue
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from gcn_file_finder.core.config import configured_model_dir, load_config
from gcn_file_finder.core.document_processor import ProcessingOptions
from gcn_file_finder.core.excel_reader import list_columns, list_sheets, read_gcn_excel
from gcn_file_finder.core.file_scanner import scan_source_files, validate_source_destination
from gcn_file_finder.core.search_service import SearchOutcome, SearchService


@dataclass(frozen=True)
class UserInputs:
    excel: Path
    sheet: str
    column: str
    source: Path
    destination: Path
    recursive: bool
    verify_content: bool
    include_red: bool
    find_all: bool
    thorough_ocr: bool
    use_gpu: bool
    dpi: int
    workers: int


class QueueLogHandler(logging.Handler):
    """Chuyển log từ worker về queue UI."""

    def __init__(self, events: queue.Queue) -> None:
        super().__init__()
        self.events = events

    def emit(self, record: logging.LogRecord) -> None:
        self.events.put(("log", self.format(record)))


class MainWindow(ctk.CTk):
    """Giao diện chính; mọi cập nhật widget chạy qua ``after``."""

    def __init__(self) -> None:
        super().__init__()
        self.title("CÔNG CỤ TÌM FILE THEO SỐ GCN")
        self.geometry("1120x820")
        self.minsize(980, 720)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.config_data = load_config()
        self.model_dir = configured_model_dir(self.config_data)
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.service = SearchService()
        self.files: list[Path] = []
        self.excel_data = None
        self.is_running = False
        self.close_when_finished = False
        self.processed = 0
        self.found_keys: set[str] = set()
        self._build_ui()
        self._configure_logging()
        self.after(100, self._drain_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(12, weight=1)
        title = ctk.CTkLabel(self, text="CÔNG CỤ TÌM FILE THEO SỐ GCN", font=ctk.CTkFont(size=22, weight="bold"))
        title.grid(row=0, column=0, columnspan=3, padx=16, pady=(14, 10))

        self.excel_var = ctk.StringVar()
        self.source_var = ctk.StringVar()
        self.destination_var = ctk.StringVar()
        self.sheet_var = ctk.StringVar()
        self.column_var = ctk.StringVar(value="GCN")
        self.recursive_var = ctk.BooleanVar(value=True)
        self.verify_var = ctk.BooleanVar(value=False)
        self.red_var = ctk.BooleanVar(value=False)
        self.all_var = ctk.BooleanVar(value=True)
        self.thorough_var = ctk.BooleanVar(value=False)
        self.gpu_var = ctk.BooleanVar(value=False)
        self.dpi_var = ctk.StringVar(value=str(self.config_data.get("default_dpi", 150)))
        self.worker_var = ctk.StringVar(value=str(self.config_data.get("default_workers", 2)))

        self._path_row(1, "File Excel", self.excel_var, self._choose_excel, "Chọn Excel")
        ctk.CTkLabel(self, text="Sheet Excel", anchor="w").grid(row=2, column=0, padx=(16, 8), pady=5, sticky="ew")
        self.sheet_combo = ctk.CTkComboBox(self, variable=self.sheet_var, values=[], command=self._sheet_changed)
        self.sheet_combo.grid(row=2, column=1, padx=8, pady=5, sticky="ew")
        ctk.CTkLabel(self, text="Cột GCN").grid(row=2, column=2, padx=(8, 16), pady=5)
        self.column_combo = ctk.CTkComboBox(self, variable=self.column_var, values=["GCN"])
        self.column_combo.grid(row=2, column=3, padx=(0, 16), pady=5, sticky="ew")
        self._path_row(3, "Thư mục nguồn", self.source_var, self._choose_source, "Chọn nguồn")
        self._path_row(4, "Thư mục kết quả", self.destination_var, self._choose_destination, "Chọn kết quả")

        options = ctk.CTkFrame(self)
        options.grid(row=5, column=0, columnspan=4, padx=16, pady=8, sticky="ew")
        for index in range(4):
            options.grid_columnconfigure(index, weight=1)
        checks = [
            ("Quét cả thư mục con", self.recursive_var),
            ("Kiểm tra nội dung dù tên file đã có GCN", self.verify_var),
            ("Bật OCR tách màu đỏ", self.red_var),
            ("Tìm tất cả GCN trong cùng một file", self.all_var),
            ("Sử dụng GPU nếu khả dụng", self.gpu_var),
            ("OCR kỹ nhiều hướng (chậm)", self.thorough_var),
        ]
        check_positions = [(0, 0), (0, 2), (1, 0), (1, 2), (2, 0), (3, 0)]
        for (text, variable), (row, column) in zip(checks, check_positions):
            ctk.CTkCheckBox(options, text=text, variable=variable).grid(
                row=row, column=column, columnspan=2, padx=12, pady=7, sticky="w"
            )
        ctk.CTkLabel(options, text="DPI PDF").grid(row=2, column=2, padx=8, pady=7, sticky="e")
        ctk.CTkComboBox(options, variable=self.dpi_var, values=["150", "200", "300", "350"], width=90).grid(row=2, column=3, padx=8, pady=7, sticky="w")
        ctk.CTkLabel(options, text="Số luồng").grid(row=3, column=2, padx=8, pady=7, sticky="e")
        ctk.CTkComboBox(options, variable=self.worker_var, values=[str(i) for i in range(1, 9)], width=90).grid(row=3, column=3, padx=8, pady=7, sticky="w")

        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.grid(row=6, column=0, columnspan=4, padx=16, pady=4, sticky="ew")
        for index in range(6):
            buttons.grid_columnconfigure(index, weight=1)
        self.check_button = ctk.CTkButton(buttons, text="KIỂM TRA DỮ LIỆU", command=self._validate_clicked)
        self.start_button = ctk.CTkButton(buttons, text="BẮT ĐẦU TÌM", command=self._start_clicked)
        self.pause_button = ctk.CTkButton(buttons, text="TẠM DỪNG", command=self._pause, state="disabled")
        self.resume_button = ctk.CTkButton(buttons, text="TIẾP TỤC", command=self._resume, state="disabled")
        self.stop_button = ctk.CTkButton(buttons, text="DỪNG", command=self._stop, state="disabled", fg_color="#B33A3A")
        self.open_button = ctk.CTkButton(buttons, text="MỞ THƯ MỤC KẾT QUẢ", command=self._open_destination)
        for index, button in enumerate((self.check_button, self.start_button, self.pause_button, self.resume_button, self.stop_button, self.open_button)):
            button.grid(row=0, column=index, padx=4, pady=4, sticky="ew")

        self.progress = ctk.CTkProgressBar(self)
        self.progress.set(0)
        self.progress.grid(row=7, column=0, columnspan=4, padx=16, pady=(8, 4), sticky="ew")
        self.current_label = ctk.CTkLabel(self, text="Chưa xử lý file", anchor="w")
        self.current_label.grid(row=8, column=0, columnspan=4, padx=16, pady=2, sticky="ew")
        self.stats_label = ctk.CTkLabel(self, text="Tổng file: 0 | Đã xử lý: 0 | Tổng GCN: 0 | Đã tìm: 0 | Chưa tìm: 0", anchor="w")
        self.stats_label.grid(row=9, column=0, columnspan=4, padx=16, pady=2, sticky="ew")
        self.validation_label = ctk.CTkLabel(self, text="Chưa kiểm tra dữ liệu", anchor="w", justify="left")
        self.validation_label.grid(row=10, column=0, columnspan=4, padx=16, pady=4, sticky="ew")
        ctk.CTkLabel(self, text="Nhật ký thời gian thực", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=11, column=0, columnspan=4, padx=16, pady=(6, 2), sticky="ew")
        self.log_box = ctk.CTkTextbox(self, wrap="word")
        self.log_box.grid(row=12, column=0, columnspan=4, padx=16, pady=(0, 14), sticky="nsew")
        self.log_box.configure(state="disabled")

    def _path_row(self, row: int, label: str, variable, command, button_text: str) -> None:
        ctk.CTkLabel(self, text=label, anchor="w").grid(row=row, column=0, padx=(16, 8), pady=5, sticky="ew")
        ctk.CTkEntry(self, textvariable=variable).grid(row=row, column=1, columnspan=2, padx=8, pady=5, sticky="ew")
        ctk.CTkButton(self, text=button_text, width=130, command=command).grid(row=row, column=3, padx=(8, 16), pady=5)

    def _configure_logging(self) -> None:
        root = logging.getLogger()
        root.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S")
        queue_handler = QueueLogHandler(self.events)
        queue_handler.setFormatter(formatter)
        root.addHandler(queue_handler)
        try:
            file_handler = logging.FileHandler(Path.cwd() / str(self.config_data.get("log_file", "gcn_file_finder.log")), encoding="utf-8")
            file_handler.setFormatter(formatter)
            root.addHandler(file_handler)
        except OSError:
            logging.warning("Không tạo được file log; nhật ký vẫn hiển thị trên giao diện.")

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        self.excel_var.set(path)
        try:
            sheets = list_sheets(path)
            self.sheet_combo.configure(values=sheets)
            if sheets:
                self.sheet_var.set(sheets[0])
                self._sheet_changed(sheets[0])
        except Exception as exc:
            messagebox.showerror("Lỗi Excel", str(exc))

    def _sheet_changed(self, sheet: str) -> None:
        try:
            columns = list_columns(self.excel_var.get(), sheet)
            self.column_combo.configure(values=columns)
            preferred = next((column for column in columns if "".join(column.split(" - ", 1)[-1].upper().split()) == "GCN"), None)
            self.column_var.set(preferred or (columns[0] if columns else "GCN"))
        except Exception as exc:
            logging.error("Không đọc được danh sách cột: %s", exc)

    def _choose_source(self) -> None:
        path = filedialog.askdirectory(title="Chọn thư mục nguồn")
        if path:
            self.source_var.set(path)

    def _choose_destination(self) -> None:
        path = filedialog.askdirectory(title="Chọn thư mục kết quả")
        if path:
            self.destination_var.set(path)

    def _inputs(self) -> UserInputs:
        try:
            dpi, workers = int(self.dpi_var.get()), int(self.worker_var.get())
        except ValueError as exc:
            raise ValueError("DPI và số luồng phải là số nguyên.") from exc
        return UserInputs(
            Path(self.excel_var.get()), self.sheet_var.get().strip(), self.column_var.get().strip(),
            Path(self.source_var.get()), Path(self.destination_var.get()), self.recursive_var.get(),
            self.verify_var.get(), self.red_var.get(), self.all_var.get(), self.thorough_var.get(),
            self.gpu_var.get(), dpi, workers,
        )

    def _validate_clicked(self) -> None:
        self._begin_validation(False)

    def _start_clicked(self) -> None:
        self._begin_validation(True)

    def _begin_validation(self, start_after: bool) -> None:
        if self.is_running:
            return
        try:
            inputs = self._inputs()
        except Exception as exc:
            messagebox.showerror("Thiếu dữ liệu", str(exc))
            return
        self.check_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.validation_label.configure(text="Đang kiểm tra dữ liệu và khởi tạo PaddleOCR...")
        threading.Thread(target=self._validation_job, args=(inputs, start_after), daemon=True).start()

    def _validation_job(self, inputs: UserInputs, start_after: bool) -> None:
        try:
            if not inputs.excel.is_file():
                raise ValueError("File Excel không tồn tại.")
            if not inputs.sheet:
                raise ValueError("Chưa chọn sheet Excel.")
            source, destination = validate_source_destination(inputs.source, inputs.destination)
            excel_data = read_gcn_excel(inputs.excel, inputs.sheet, inputs.column)
            if not excel_data.valid_gcn_keys:
                raise ValueError("Excel không có GCN hợp lệ nào.")
            files = scan_source_files(source, destination, inputs.recursive)
            self.service.prepare_engine(inputs.use_gpu, self.model_dir)
            self.events.put(("validated", (inputs, excel_data, files, start_after)))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _launch_search(self, inputs: UserInputs) -> None:
        self.is_running = True
        self.processed = 0
        self.found_keys.clear()
        self.pause_button.configure(state="normal")
        self.stop_button.configure(state="normal")
        options = ProcessingOptions(
            dpi=inputs.dpi,
            verify_content_after_filename=inputs.verify_content,
            include_red=inputs.include_red,
            find_all=inputs.find_all,
            thorough_ocr=inputs.thorough_ocr,
        )
        threading.Thread(
            target=self._run_search_job,
            args=(self.files, self.excel_data, inputs.destination, options, inputs.workers, inputs.use_gpu, self.model_dir, self._service_event),
            daemon=True,
        ).start()

    def _run_search_job(self, *args) -> None:
        try:
            self.service.run(*args)
        except Exception as exc:
            logging.exception("Phiên tìm kiếm thất bại")
            self.events.put(("search_error", str(exc)))

    def _service_event(self, name: str, payload: object) -> None:
        self.events.put((name, payload))

    def _pause(self) -> None:
        self.service.pause()
        self.pause_button.configure(state="disabled")
        self.resume_button.configure(state="normal")
        logging.info("Đã tạm dừng nhận file mới; các tác vụ đang chạy sẽ hoàn thành an toàn.")

    def _resume(self) -> None:
        self.service.resume()
        self.pause_button.configure(state="normal")
        self.resume_button.configure(state="disabled")
        logging.info("Đã tiếp tục.")

    def _stop(self) -> None:
        self.service.stop()
        self.stop_button.configure(state="disabled")
        logging.info("Đã yêu cầu dừng; đang hoàn thành tác vụ hiện tại và lưu báo cáo/cache.")

    def _open_destination(self) -> None:
        path = Path(self.destination_var.get())
        if path.is_dir():
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            messagebox.showwarning("Thư mục", "Thư mục kết quả chưa tồn tại.")

    def _update_stats(self) -> None:
        total_gcn = len(self.excel_data.valid_gcn_keys) if self.excel_data else 0
        self.stats_label.configure(
            text=f"Tổng file: {len(self.files)} | Đã xử lý: {self.processed} | Tổng GCN: {total_gcn} | "
                 f"Đã tìm: {len(self.found_keys)} | Chưa tìm: {max(0, total_gcn - len(self.found_keys))}"
        )
        self.progress.set(self.processed / len(self.files) if self.files else 0)

    def _append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _drain_events(self) -> None:
        try:
            while True:
                name, payload = self.events.get_nowait()
                if name == "log":
                    self._append_log(str(payload))
                elif name == "error":
                    self.check_button.configure(state="normal")
                    self.start_button.configure(state="normal")
                    self.validation_label.configure(text=f"Lỗi: {payload}")
                    messagebox.showerror("Không thể tiếp tục", str(payload))
                elif name == "search_error":
                    self.is_running = False
                    self.check_button.configure(state="normal")
                    self.start_button.configure(state="normal")
                    self.pause_button.configure(state="disabled")
                    self.resume_button.configure(state="disabled")
                    self.stop_button.configure(state="disabled")
                    self.current_label.configure(text=f"Lỗi phiên chạy: {payload}")
                    messagebox.showerror("Lỗi phiên chạy", str(payload))
                elif name == "validated":
                    inputs, self.excel_data, self.files, start_after = payload  # type: ignore[misc]
                    self.check_button.configure(state="normal")
                    self.start_button.configure(state="normal")
                    data = self.excel_data
                    self.validation_label.configure(
                        text=f"Tổng dòng Excel: {data.total_data_rows} | GCN hợp lệ: {len(data.valid_gcn_keys)} | "
                             f"Không hợp lệ: {data.invalid_count} | Trùng: {data.duplicate_count} | File chuẩn bị quét: {len(self.files)}"
                    )
                    self._update_stats()
                    logging.info("Kiểm tra dữ liệu thành công. Model OCR offline: %s", self.model_dir)
                    if start_after:
                        self._launch_search(inputs)
                elif name == "current":
                    self.current_label.configure(text=f"Đang xử lý: {payload}")
                elif name == "file_done":
                    self.processed += 1
                    result = payload
                    self.found_keys.update(match.key for match in result.matches if match.key)
                    logging.info("[%d/%d] %s - %s", self.processed, len(self.files), result.source_path.name, result.status)
                    self._update_stats()
                elif name == "finished":
                    outcome: SearchOutcome = payload  # type: ignore[assignment]
                    self.is_running = False
                    self.check_button.configure(state="normal")
                    self.start_button.configure(state="normal")
                    self.pause_button.configure(state="disabled")
                    self.resume_button.configure(state="disabled")
                    self.stop_button.configure(state="disabled")
                    message = "Đã dừng an toàn" if outcome.stopped else "Đã hoàn thành"
                    self.current_label.configure(text=f"{message}. Báo cáo: {outcome.report_path}")
                    logging.info("%s. Đã xử lý %d file, báo cáo: %s", message, len(outcome.file_results), outcome.report_path)
                    if self.close_when_finished:
                        self.service.shutdown()
                        self.destroy()
                        return
        except queue.Empty:
            pass
        finally:
            self.after(100, self._drain_events)

    def _on_close(self) -> None:
        if self.is_running:
            if not messagebox.askyesno("Đang xử lý", "Dừng an toàn và đóng sau khi lưu báo cáo?"):
                return
            self.close_when_finished = True
            self.service.stop()
            self.current_label.configure(text="Đang dừng an toàn và lưu báo cáo trước khi đóng...")
            return
        self.service.shutdown()
        self.destroy()
