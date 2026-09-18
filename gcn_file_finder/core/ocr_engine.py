"""Bộ máy PaddleOCR dùng chung, an toàn khi gọi từ nhiều worker."""

from __future__ import annotations

import logging
import os
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import cv2
import numpy as np

LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class OCRLine:
    """Một dòng văn bản và độ tin cậy do OCR trả về."""

    text: str
    confidence: float


class OCREngine:
    """Khởi tạo PaddleOCR đúng một lần và tuần tự hóa suy luận model."""

    def __init__(
        self,
        use_gpu: bool = False,
        model_dir: str | Path | None = None,
        cpu_threads: int | None = None,
    ) -> None:
        self.use_gpu = use_gpu
        self.model_dir = Path(model_dir).resolve() if model_dir else None
        # Giới hạn số luồng tính toán CPU của riêng model này. Bắt buộc phải
        # đặt khi chạy nhiều tiến trình OCR song song (xem search_service.py):
        # mỗi tiến trình mặc định tự dùng hết số lõi máy cho các phép toán
        # BLAS/oneDNN, nên N tiến trình chạy cùng lúc mà không giới hạn sẽ
        # tranh chấp lõi CPU và làm MỌI tiến trình chậm đi thay vì nhanh hơn.
        self.cpu_threads = cpu_threads
        self._lock = threading.Lock()
        self._ocr = self._create_engine()

    def _create_engine(self):
        if self.model_dir:
            missing = [name for name in ("det", "rec", "cls") if not (self.model_dir / name).is_dir() or not any((self.model_dir / name).iterdir())]
            if missing:
                raise RuntimeError(
                    f"Thiếu model OCR offline ({', '.join(missing)}) trong {self.model_dir}. "
                    "Hãy chạy download_models.ps1 một lần khi có Internet."
                )
            # PaddleX không kiểm tra các host model khi ta đã cung cấp đủ model local.
            os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "1")
        try:
            from paddleocr import PaddleOCR
        except ImportError as exc:
            raise RuntimeError("Chưa cài PaddleOCR. Hãy chạy run.ps1 khi có Internet để cài lần đầu.") from exc

        # PaddleOCR 3.x trước, sau đó tương thích 2.x.
        modern: dict[str, Any] = {
            "lang": "vi",
            "ocr_version": "PP-OCRv5",
            "device": "gpu:0" if self.use_gpu else "cpu",
            # PaddlePaddle 3.3.0 on Windows can fail in the oneDNN/PIR path with
            # ConvertPirAttribute2RuntimeAttribute. Disable only that CPU
            # optimization; regular Paddle inference remains enabled.
            "enable_mkldnn": False,
            # Hướng cả trang đã được thử 0/90/180/270 ở image_processor.
            "use_doc_orientation_classify": False,
            "use_doc_unwarping": False,
            # Trang đã được đặt đúng chiều ở chế độ nhanh, hoặc xoay thủ công
            # ở chế độ kỹ; bỏ model phân loại từng dòng để giảm thời gian CPU.
            "use_textline_orientation": False,
        }
        legacy: dict[str, Any] = {
            "lang": "vi",
            "use_angle_cls": True,
            "use_gpu": self.use_gpu,
            "show_log": False,
        }
        if self.cpu_threads and self.cpu_threads > 0:
            modern["cpu_threads"] = self.cpu_threads
            legacy["cpu_threads"] = self.cpu_threads
        if self.model_dir:
            # Thư mục này chứa các thư mục con det/rec/cls đã tải một lần.
            # Với PaddleOCR 3.x, tên/path model đã xác định đầy đủ nên không
            # truyền lang/ocr_version để tránh cảnh báo chúng bị bỏ qua.
            modern.pop("lang", None)
            modern.pop("ocr_version", None)
            modern.update({
                "text_detection_model_name": "PP-OCRv5_mobile_det",
                "text_detection_model_dir": str(self.model_dir / "det"),
                "text_recognition_model_name": "latin_PP-OCRv5_mobile_rec",
                "text_recognition_model_dir": str(self.model_dir / "rec"),
                "textline_orientation_model_name": "PP-LCNet_x1_0_textline_ori",
                "textline_orientation_model_dir": str(self.model_dir / "cls"),
            })
            legacy.update({
                "det_model_dir": str(self.model_dir / "det"),
                "rec_model_dir": str(self.model_dir / "rec"),
                "cls_model_dir": str(self.model_dir / "cls"),
            })
        try:
            return PaddleOCR(**modern)
        except TypeError as modern_error:
            LOGGER.info("Thử API PaddleOCR 2.x do API 3.x không tương thích: %s", modern_error)
            try:
                return PaddleOCR(**legacy)
            except Exception as exc:
                raise RuntimeError(f"Không khởi tạo được PaddleOCR: {exc}") from exc
        except Exception as exc:
            raise RuntimeError(f"Không khởi tạo được PaddleOCR: {exc}") from exc

    def recognize(self, image: np.ndarray) -> list[OCRLine]:
        """OCR một ảnh OpenCV và chuẩn hóa kết quả của PaddleOCR 2.x/3.x."""

        # PaddleOCR/PaddleX giả định ảnh luôn có 3 kênh màu; các biến thể tiền
        # xử lý ảnh xám/nhị phân (ẢNH XÁM, THRESHOLD, LÀM NÉT, TÁCH MÀU ĐỎ,
        # NỀN ĐỎ THÀNH TRẮNG) chỉ có 1 kênh (mảng 2 chiều) và làm predict()
        # crash ngay ở bước resize nội bộ ("not enough values to unpack").
        # Chuyển lại 3 kênh (giá trị giống nhau ở cả 3) trước khi gọi model.
        if image.ndim == 2:
            image = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)

        try:
            with self._lock:
                if hasattr(self._ocr, "predict"):
                    result = list(self._ocr.predict(image))
                else:
                    result = self._ocr.ocr(image, cls=True)
        except Exception as exc:
            raise RuntimeError(f"PaddleOCR xử lý ảnh thất bại: {exc}") from exc
        return self._parse_result(result)

    @staticmethod
    def _parse_result(result: Any) -> list[OCRLine]:
        lines: list[OCRLine] = []
        if not result:
            return lines
        for page in result if isinstance(result, list) else [result]:
            payload = getattr(page, "json", None)
            if callable(payload):
                payload = payload()
            if isinstance(payload, dict):
                payload = payload.get("res", payload)
            elif hasattr(page, "res"):
                payload = page.res
            if isinstance(payload, dict) and "rec_texts" in payload:
                scores = payload.get("rec_scores", [0.0] * len(payload["rec_texts"]))
                lines.extend(OCRLine(str(text), float(score)) for text, score in zip(payload["rec_texts"], scores))
                continue
            # Dạng 2.x: [[box, (text, score)], ...]
            rows = page if isinstance(page, list) else []
            for row in rows:
                if isinstance(row, list) and len(row) == 1 and isinstance(row[0], list):
                    row = row[0]
                try:
                    text, score = row[1]
                    lines.append(OCRLine(str(text), float(score)))
                except (TypeError, ValueError, IndexError):
                    continue
        return lines
