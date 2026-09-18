"""Điểm vào của ứng dụng Windows tìm file theo số GCN."""

from __future__ import annotations

import multiprocessing
import sys
from pathlib import Path

# Cho phép cả ``python app.py`` và import/chạy theo package.
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from gcn_file_finder.ui.main_window import MainWindow


def _run_ocr_self_test() -> None:
    """Exercise model creation and inference inside a packaged executable."""

    import numpy as np

    from gcn_file_finder.core.config import configured_model_dir, load_config
    from gcn_file_finder.core.ocr_engine import OCREngine

    engine = OCREngine(use_gpu=False, model_dir=configured_model_dir(load_config()))
    engine.recognize(np.full((200, 300, 3), 255, dtype=np.uint8))


def main() -> None:
    """Khởi chạy giao diện ứng dụng."""

    multiprocessing.freeze_support()
    if "--ocr-self-test" in sys.argv:
        _run_ocr_self_test()
        return
    MainWindow().mainloop()


if __name__ == "__main__":
    main()
