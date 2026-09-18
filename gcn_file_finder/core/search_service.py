"""Điều phối worker nền, tạm dừng/dừng, sao chép và lưu báo cáo."""

from __future__ import annotations

import logging
import os
import threading
import time
from concurrent.futures import FIRST_COMPLETED, Future, ProcessPoolExecutor, wait
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from gcn_file_finder.core.cache_manager import CacheManager
from gcn_file_finder.core.document_processor import (
    ProcessingOptions,
    deserialize_result,
    process_document,
    serialize_result,
)
from gcn_file_finder.core.file_copier import copy_for_gcn
from gcn_file_finder.core.ocr_engine import OCREngine
from gcn_file_finder.core.report_writer import CopiedFile, write_report
from gcn_file_finder.models import ExcelData, FileScanResult

LOGGER = logging.getLogger(__name__)

# Biến toàn cục riêng của mỗi tiến trình con: PaddleOCR không thể chia sẻ qua
# tiến trình (không pickle được), nên mỗi worker tự nạp và tái sử dụng đúng
# một model, tương tự cách self._engine từng được tái sử dụng khi còn chạy
# bằng thread. Việc dùng tiến trình thay vì thread bỏ được lock tuần tự hoá
# trong OCREngine.recognize, cho phép nhiều ảnh được OCR thật sự song song
# trên nhiều lõi CPU.
_WORKER_ENGINE: OCREngine | None = None


def _init_worker(use_gpu: bool, model_dir: str | None, cpu_threads: int | None) -> None:
    """Chạy đúng một lần khi tiến trình con khởi động.

    Phải giới hạn số luồng CPU trước khi PaddleOCR/OpenCV được import: mỗi
    tiến trình mặc định tự chiếm hết số lõi máy cho BLAS/oneDNN, nên chạy
    nhiều tiến trình cùng lúc mà không giới hạn sẽ làm tất cả chậm đi vì
    tranh chấp lõi CPU, thay vì nhanh hơn như kỳ vọng của multiprocessing.
    """

    global _WORKER_ENGINE
    if cpu_threads and cpu_threads > 0:
        for var in ("OMP_NUM_THREADS", "MKL_NUM_THREADS", "OPENBLAS_NUM_THREADS"):
            os.environ.setdefault(var, str(cpu_threads))
        try:
            import cv2

            cv2.setNumThreads(cpu_threads)
        except ImportError:
            pass
    _WORKER_ENGINE = OCREngine(
        use_gpu=use_gpu, model_dir=Path(model_dir) if model_dir else None, cpu_threads=cpu_threads,
    )


def _process_in_worker(path: Path, valid_keys: set[str], options: ProcessingOptions) -> FileScanResult:
    """Hàm thực thi trong tiến trình con; cache được đọc/ghi tập trung ở tiến trình chính."""

    if _WORKER_ENGINE is None:  # pragma: no cover - chỉ xảy ra nếu initializer lỗi
        raise RuntimeError("OCREngine chưa được khởi tạo trong tiến trình worker.")
    return process_document(path, valid_keys, _WORKER_ENGINE, options, cache=None)


@dataclass
class SearchOutcome:
    file_results: list[FileScanResult] = field(default_factory=list)
    copied_files: list[CopiedFile] = field(default_factory=list)
    stopped: bool = False
    report_path: Path | None = None


class SearchService:
    """Chạy pipeline trong thread riêng; UI nhận sự kiện qua callback/queue."""

    def __init__(self) -> None:
        self.stop_event = threading.Event()
        self.resume_event = threading.Event()
        self.resume_event.set()
        self._engine: OCREngine | None = None
        self._engine_settings: tuple[bool, str] | None = None
        self._executor: ProcessPoolExecutor | None = None
        self._executor_settings: tuple[int, bool, str] | None = None

    def prepare_engine(self, use_gpu: bool, model_dir: Path | None) -> OCREngine:
        """Khởi tạo hoặc tái sử dụng đúng một model trong tiến trình chính.

        Chỉ dùng để kiểm tra nhanh (bước KIỂM TRA DỮ LIỆU) rằng model nạp
        được trước khi chạy quét thật; việc quét thật dùng tiến trình con
        riêng, xem ``_get_executor``.
        """

        settings = (use_gpu, str(model_dir.resolve()) if model_dir else "")
        if self._engine is None or self._engine_settings != settings:
            self._engine = OCREngine(use_gpu=use_gpu, model_dir=model_dir)
            self._engine_settings = settings
        return self._engine

    def _get_executor(self, worker_count: int, use_gpu: bool, model_dir: Path | None) -> ProcessPoolExecutor:
        """Tái sử dụng pool tiến trình khi cấu hình không đổi giữa các lần chạy.

        Mỗi tiến trình con tự nạp PaddleOCR (xem ``_init_worker``) nên tạo pool
        mới luôn kéo theo nạp lại model; tái sử dụng giúp các lần bấm
        "BẮT ĐẦU TÌM" liên tiếp với cùng cấu hình không mất thời gian nạp lại.
        """

        model_dir_str = str(model_dir.resolve()) if model_dir else ""
        settings = (max(1, worker_count), use_gpu, model_dir_str)
        if self._executor is None or self._executor_settings != settings:
            self.shutdown()
            # Đo thực tế cho thấy tăng cpu_threads không rút ngắn thời gian
            # một lệnh OCR (Paddle không song song hoá tốt theo luồng cho
            # workload này), trong khi Paddle tự cảnh báo OMP_NUM_THREADS > 1
            # kết hợp nhiều tiến trình ("data parallel") gây tranh chấp lõi
            # CPU và làm chậm đi. Để mỗi tiến trình dùng đúng 1 luồng CPU;
            # độ song song đến từ số tiến trình, không phải số luồng/tiến trình.
            cpu_threads = 1
            self._executor = ProcessPoolExecutor(
                max_workers=settings[0],
                initializer=_init_worker,
                initargs=(use_gpu, model_dir_str or None, cpu_threads),
            )
            self._executor_settings = settings
        return self._executor

    def shutdown(self) -> None:
        """Dừng hẳn pool tiến trình con, gọi khi đổi cấu hình hoặc đóng ứng dụng."""

        if self._executor is not None:
            self._executor.shutdown(wait=True, cancel_futures=True)
            self._executor = None
            self._executor_settings = None

    def pause(self) -> None:
        self.resume_event.clear()

    def resume(self) -> None:
        self.resume_event.set()

    def stop(self) -> None:
        self.stop_event.set()
        self.resume_event.set()

    def reset(self) -> None:
        self.stop_event.clear()
        self.resume_event.set()

    def run(
        self,
        files: list[Path],
        excel_data: ExcelData,
        destination: Path,
        options: ProcessingOptions,
        worker_count: int,
        use_gpu: bool,
        model_dir: Path | None,
        on_event: Callable[[str, object], None] | None = None,
    ) -> SearchOutcome:
        """Quét tài liệu, chỉ nhận thêm task khi chưa dừng và lưu báo cáo cuối."""

        self.reset()
        started = time.perf_counter()
        outcome = SearchOutcome()
        cache = CacheManager(destination / ".gcn_scan_cache.json")
        settings_key = options.cache_key(excel_data.valid_gcn_keys)
        emit = on_event or (lambda _name, _payload: None)
        pending: dict[Future, Path] = {}
        file_iterator = iter(files)

        def record_result(result: FileScanResult, from_cache: bool) -> None:
            outcome.file_results.append(result)
            emit("file_done", result)
            # Lỗi tạm thời (IO/model) không nên bị nhớ, để lần chạy sau thử lại
            # thay vì luôn trả về LỖI XỬ LÝ từ cache.
            if not from_cache and result.status != "LỖI XỬ LÝ":
                cache.put(result.source_path, serialize_result(result), settings_key)
            if len(outcome.file_results) % 10 == 0:
                try:
                    cache.save()
                    write_report(
                        destination / "KET_QUA_TIM_GCN.xlsx", excel_data,
                        sorted(outcome.file_results, key=lambda item: str(item.source_path).casefold()), [],
                        time.perf_counter() - started,
                    )
                    emit("log", f"Đã tự lưu báo cáo tạm sau {len(outcome.file_results)} file.")
                except OSError as exc:
                    emit("log", f"Chưa lưu được báo cáo tạm (có thể file đang mở): {exc}")

        def submit_next(executor: ProcessPoolExecutor) -> bool:
            while True:
                if self.stop_event.is_set():
                    return False
                self.resume_event.wait()
                if self.stop_event.is_set():
                    return False
                try:
                    path = next(file_iterator)
                except StopIteration:
                    return False
                # Cache được đọc/ghi chỉ trong tiến trình chính: tránh chia sẻ
                # CacheManager (có threading.Lock, không pickle được) qua các
                # tiến trình con, và tránh OCR lại khi trùng file/cấu hình.
                cached = cache.get(path, settings_key)
                if cached and cached.get("status") != "LỖI XỬ LÝ":
                    record_result(deserialize_result(path, cached), from_cache=True)
                    continue
                emit("current", path)
                future = executor.submit(_process_in_worker, path, excel_data.valid_gcn_keys, options)
                pending[future] = path
                return True

        executor = self._get_executor(worker_count, use_gpu, model_dir)
        for _ in range(max(1, worker_count)):
            if not submit_next(executor):
                break
        while pending:
            completed, _ = wait(pending, return_when=FIRST_COMPLETED)
            for future in completed:
                path = pending.pop(future)
                try:
                    result = future.result()
                except Exception as exc:
                    result = FileScanResult(path, path.suffix.lstrip(".").upper(), status="LỖI XỬ LÝ", error=str(exc))
                record_result(result, from_cache=False)
                submit_next(executor)

        outcome.stopped = self.stop_event.is_set()
        outcome.file_results.sort(key=lambda item: str(item.source_path).casefold())
        # Thứ tự nguồn đã ổn định, nhờ đó hậu tố kết quả cũng ổn định.
        for result in outcome.file_results:
            keys = list(dict.fromkeys(match.key for match in result.matches if match.key))
            for key in keys:
                if self.stop_event.is_set() and not outcome.stopped:
                    break
                try:
                    copied = copy_for_gcn(result.source_path, destination, key)
                    outcome.copied_files.append(CopiedFile(key, result.source_path, copied.destination, copied.status))
                except Exception as exc:
                    result.status = "LỖI XỬ LÝ"
                    result.error = f"Lỗi sao chép: {exc}"
                    LOGGER.exception("Không sao chép được %s", result.source_path)
        try:
            cache.save()
        except OSError as exc:
            emit("log", f"Không lưu được cache: {exc}")
        try:
            outcome.report_path = write_report(
                destination / "KET_QUA_TIM_GCN.xlsx", excel_data, outcome.file_results, outcome.copied_files,
                time.perf_counter() - started,
            )
        except OSError as exc:
            fallback = destination / f"KET_QUA_TIM_GCN_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            emit("log", f"Báo cáo chuẩn đang bị khóa ({exc}); lưu sang {fallback.name}.")
            outcome.report_path = write_report(
                fallback, excel_data, outcome.file_results, outcome.copied_files, time.perf_counter() - started
            )
        emit("finished", outcome)
        return outcome
