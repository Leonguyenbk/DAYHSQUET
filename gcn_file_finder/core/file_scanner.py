"""Liệt kê file nguồn và loại bỏ file tạm, báo cáo, thư mục kết quả."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable

SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
REPORT_FILENAME = "KET_QUA_TIM_GCN.xlsx"


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def validate_source_destination(source_dir: str | Path, destination_dir: str | Path) -> tuple[Path, Path]:
    """Kiểm tra hai thư mục và cấm thư mục đích nằm trong nguồn."""

    source = Path(source_dir).resolve()
    destination = Path(destination_dir).resolve()
    if not source.is_dir():
        raise ValueError("Thư mục nguồn không tồn tại hoặc không phải thư mục.")
    if destination == source or _is_relative_to(destination, source):
        raise ValueError("Thư mục kết quả không được nằm trong thư mục nguồn.")
    if not os.access(source, os.R_OK):
        raise PermissionError("Không có quyền đọc thư mục nguồn.")
    destination.mkdir(parents=True, exist_ok=True)
    probe = destination / ".gcn_write_test"
    try:
        probe.touch(exist_ok=False)
        probe.unlink()
    except OSError as exc:
        raise PermissionError("Không có quyền ghi thư mục kết quả.") from exc
    return source, destination


def scan_source_files(
    source_dir: str | Path,
    destination_dir: str | Path,
    recursive: bool = True,
) -> list[Path]:
    """Trả về danh sách file hỗ trợ đã sắp xếp ổn định."""

    source = Path(source_dir).resolve()
    destination = Path(destination_dir).resolve()
    iterator: Iterable[Path] = source.rglob("*") if recursive else source.glob("*")
    files: list[Path] = []
    for path in iterator:
        try:
            resolved = path.resolve()
        except OSError:
            continue
        if not path.is_file() or _is_relative_to(resolved, destination):
            continue
        lower_name = path.name.lower()
        if path.name.startswith("~$") or lower_name.endswith((".tmp", ".temp", ".part")):
            continue
        if lower_name == REPORT_FILENAME.lower() or path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        files.append(path)
    return sorted(files, key=lambda item: str(item).casefold())
