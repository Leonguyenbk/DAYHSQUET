from __future__ import annotations

import sys
from types import SimpleNamespace

import numpy as np

from gcn_file_finder.core.document_processor import ProcessingOptions, _ocr_image, process_document
from gcn_file_finder.core.image_processor import limit_image_size, whiten_red_background
from gcn_file_finder.core.ocr_engine import OCREngine


def test_limit_image_size_preserves_ratio_and_does_not_upscale():
    large = np.zeros((41, 58, 3), dtype=np.uint8)
    resized = limit_image_size(large, max_side=40)

    assert resized.shape[:2] == (28, 40)
    small = np.zeros((100, 200, 3), dtype=np.uint8)
    assert limit_image_size(small) is small


def test_modern_cpu_engine_disables_mkldnn(monkeypatch):
    received = {}

    class FakePaddleOCR:
        def __init__(self, **kwargs):
            received.update(kwargs)

    monkeypatch.setitem(sys.modules, "paddleocr", SimpleNamespace(PaddleOCR=FakePaddleOCR))
    OCREngine(use_gpu=False)

    assert received["device"] == "cpu"
    assert received["enable_mkldnn"] is False
    assert received["use_textline_orientation"] is False


def test_fast_ocr_falls_back_to_enhanced_gray_when_original_fails():
    class CountingEngine:
        def __init__(self):
            self.calls = 0

        def recognize(self, _image):
            self.calls += 1
            return []

    engine = CountingEngine()
    image = np.zeros((100, 200, 3), dtype=np.uint8)

    matches = _ocr_image(image, engine, {"AM143443"}, False, 1, True)

    assert matches == []
    # Ảnh gốc + ảnh xám tăng tương phản (rẻ, luôn thử khi ảnh gốc không khớp).
    assert engine.calls == 2


def test_fast_ocr_with_include_red_tries_both_red_variants():
    class CountingEngine:
        def __init__(self):
            self.calls = 0

        def recognize(self, _image):
            self.calls += 1
            return []

    engine = CountingEngine()
    image = np.zeros((100, 200, 3), dtype=np.uint8)

    matches = _ocr_image(image, engine, {"AM143443"}, True, 1, True)

    assert matches == []
    # Ảnh gốc + ảnh xám + tách màu đỏ (chữ đỏ) + nền đỏ thành trắng (nền/dấu mộc đỏ).
    assert engine.calls == 4


def test_whiten_red_background_removes_red_keeps_black_text():
    image = np.full((10, 10, 3), (0, 0, 200), dtype=np.uint8)  # BGR: đỏ đậm
    image[4:6, 4:6] = (0, 0, 0)  # vùng "chữ" màu đen giữa nền đỏ

    result = whiten_red_background(image)

    assert result.shape == (10, 10)
    assert result[0, 0] == 255  # nền đỏ thành trắng
    assert result[4, 4] == 0  # chữ đen giữ nguyên


def test_processing_errors_are_not_cached(tmp_path):
    source = tmp_path / "broken.pdf"
    source.write_bytes(b"not a pdf")

    class RecordingCache:
        def __init__(self):
            self.put_calls = []

        def get(self, _path, _settings_key):
            return {"status": "LỖI XỬ LÝ"}

        def put(self, *args):
            self.put_calls.append(args)

    cache = RecordingCache()
    result = process_document(source, {"AM143443"}, object(), ProcessingOptions(), cache)

    assert result.status == "LỖI XỬ LÝ"
    assert cache.put_calls == []
