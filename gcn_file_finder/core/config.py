"""Đọc cấu hình JSON không chứa dữ liệu cá nhân."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


def application_dir() -> Path:
    """Thư mục người dùng đặt mã nguồn hoặc executable PyInstaller."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resource_dir() -> Path:
    """Thư mục tài nguyên do PyInstaller giải nén/đặt trong bundle."""

    bundled = getattr(sys, "_MEIPASS", None)
    return Path(bundled) if bundled else application_dir()


DEFAULT_CONFIG: dict[str, Any] = {
    "model_dir": "models/paddleocr",
    "default_dpi": 150,
    "default_workers": 2,
    "log_file": "gcn_file_finder.log",
}


def load_config() -> dict[str, Any]:
    """Trộn config.json cạnh ứng dụng với mặc định an toàn."""

    config = dict(DEFAULT_CONFIG)
    path = application_dir() / "config.json"
    if path.is_file():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                config.update(loaded)
        except (OSError, json.JSONDecodeError):
            pass
    return config


def configured_model_dir(config: dict[str, Any]) -> Path:
    """Giải đường dẫn model tương đối theo thư mục ứng dụng."""

    path = Path(str(config.get("model_dir", DEFAULT_CONFIG["model_dir"])))
    if path.is_absolute():
        return path
    external = application_dir() / path
    return external if external.exists() else resource_dir() / path
