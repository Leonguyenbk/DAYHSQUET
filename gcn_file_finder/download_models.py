"""Tải model PaddleOCR tiếng Việt một lần vào thư mục dự án.

Đây là thao tác mạng chủ động, tách khỏi ứng dụng chính. Ứng dụng chính sẽ từ
chối OCR nếu ba thư mục model offline chưa đầy đủ.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def main() -> None:
    """Tải model PaddleOCR 3.x vào đường dẫn offline rõ ràng của dự án."""

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    project = Path(__file__).resolve().parent
    if str(project.parent) not in sys.path:
        sys.path.insert(0, str(project.parent))
    root = project / "models" / "paddleocr"
    directories = {name: root / name for name in ("det", "rec", "cls")}
    for directory in directories.values():
        directory.mkdir(parents=True, exist_ok=True)
    model_names = {
        "det": "PP-OCRv5_mobile_det",
        "rec": "latin_PP-OCRv5_mobile_rec",
        "cls": "PP-LCNet_x1_0_textline_ori",
    }
    complete = all(
        (directory / "inference.yml").is_file()
        and any(directory.glob("*.pdiparams"))
        for directory in directories.values()
    )
    if complete:
        print(f"Đã có đủ model PaddleOCR trong: {root}. Bỏ qua tải lại.")
    else:
        print(f"Đang tải model PaddleOCR tiếng Việt vào: {root}")
        models_parent = project / "models"
        models_parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="paddlex_download_", dir=models_parent) as temporary:
            os.environ["PADDLE_PDX_CACHE_HOME"] = temporary
            os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "BOS")
            from paddleocr import PaddleOCR

            PaddleOCR(
                device="cpu",
                use_doc_orientation_classify=False,
                use_doc_unwarping=False,
                use_textline_orientation=True,
                text_detection_model_name=model_names["det"],
                text_recognition_model_name=model_names["rec"],
                textline_orientation_model_name=model_names["cls"],
            )
            cache_root = Path(temporary)
            for role, model_name in model_names.items():
                candidates = [path for path in cache_root.rglob(model_name) if path.is_dir()]
                if not candidates:
                    raise RuntimeError(f"Không xác định được model vừa tải: {model_name}")
                shutil.copytree(candidates[0], directories[role], dirs_exist_ok=True)
    missing = [
        name for name, directory in directories.items()
        if not (directory / "inference.yml").is_file() or not any(directory.glob("*.pdiparams"))
    ]
    if missing:
        raise RuntimeError(f"Tải model chưa hoàn tất: {', '.join(missing)}")
    # Xác minh trong process sạch bằng chính cấu hình offline của ứng dụng.
    check_code = (
        "from pathlib import Path; "
        "from gcn_file_finder.core.ocr_engine import OCREngine; "
        f"OCREngine(use_gpu=False, model_dir=Path({str(root)!r}))"
    )
    check_env = dict(os.environ)
    check_env.pop("PADDLE_PDX_CACHE_HOME", None)
    check_env["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "1"
    subprocess.run([sys.executable, "-c", check_code], cwd=project.parent, env=check_env, check=True)
    print("Đã tải và kiểm tra đủ model det/rec/cls. Có thể ngắt Internet và chạy ứng dụng.")


if __name__ == "__main__":
    main()
