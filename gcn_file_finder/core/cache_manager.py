"""Cache JSON nguyên tử để tránh OCR lại file không đổi."""

from __future__ import annotations

import json
import os
import tempfile
import threading
from pathlib import Path
from typing import Any


class CacheManager:
    """Lưu cache theo đường dẫn tuyệt đối, kích thước, mtime và cấu hình quét."""

    VERSION = 1

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self._lock = threading.Lock()
        self._data: dict[str, Any] = {"version": self.VERSION, "files": {}}
        self.load()

    def load(self) -> None:
        try:
            loaded = json.loads(self.path.read_text(encoding="utf-8"))
            if loaded.get("version") == self.VERSION:
                self._data = loaded
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            self._data = {"version": self.VERSION, "files": {}}

    @staticmethod
    def signature(path: str | Path, settings_key: str = "") -> str:
        stat = Path(path).stat()
        return f"{Path(path).resolve()}|{stat.st_size}|{stat.st_mtime_ns}|{settings_key}"

    def get(self, path: str | Path, settings_key: str = "") -> dict[str, Any] | None:
        with self._lock:
            return self._data["files"].get(self.signature(path, settings_key))

    def put(self, path: str | Path, value: dict[str, Any], settings_key: str = "") -> None:
        with self._lock:
            absolute = str(Path(path).resolve())
            files = self._data["files"]
            for key in [key for key in files if key.split("|", 1)[0] == absolute]:
                files.pop(key, None)
            files[self.signature(path, settings_key)] = value

    def save(self) -> None:
        """Ghi file tạm rồi replace để cache không hỏng khi dừng đột ngột."""

        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._lock:
            payload = json.dumps(self._data, ensure_ascii=False, indent=2)
        handle, temporary_name = tempfile.mkstemp(prefix="gcn_cache_", suffix=".tmp", dir=self.path.parent)
        try:
            with os.fdopen(handle, "w", encoding="utf-8") as stream:
                stream.write(payload)
            os.replace(temporary_name, self.path)
        finally:
            try:
                Path(temporary_name).unlink(missing_ok=True)
            except OSError:
                pass

