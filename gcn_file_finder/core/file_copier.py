"""Sao chép an toàn, đặt tên chuẩn và chống trùng nội dung."""

from __future__ import annotations

import hashlib
import re
import shutil
from dataclasses import dataclass
from pathlib import Path

from gcn_file_finder.core.gcn_normalizer import format_gcn_display


@dataclass(frozen=True)
class CopyResult:
    destination: Path
    status: str


def sha256_file(path: str | Path, chunk_size: int = 1024 * 1024) -> str:
    """Tính SHA-256 theo luồng, không nạp cả file vào RAM."""

    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(chunk_size), b""):
            digest.update(block)
    return digest.hexdigest()


def destination_candidates(destination_dir: str | Path, gcn_key: str, extension: str):
    """Sinh tuần tự tên ``AA 123456.ext``, ``AA 123456_02.ext``..."""

    folder = Path(destination_dir)
    display = format_gcn_display(gcn_key)
    normalized_extension = extension if extension.startswith(".") else f".{extension}"
    index = 1
    while True:
        suffix = "" if index == 1 else f"_{index:02d}"
        yield folder / f"{display}{suffix}{normalized_extension.lower()}"
        index += 1


def copy_for_gcn(source_path: str | Path, destination_dir: str | Path, gcn_key: str) -> CopyResult:
    """Sao chép bằng copy2; không ghi đè và không lặp nếu nội dung đã tồn tại."""

    source = Path(source_path)
    destination = Path(destination_dir)
    destination.mkdir(parents=True, exist_ok=True)
    source_size = source.stat().st_size
    source_hash: str | None = None
    for candidate in destination_candidates(destination, gcn_key, source.suffix):
        if not candidate.exists():
            shutil.copy2(source, candidate)
            return CopyResult(candidate, "ĐÃ TÌM THẤY")
        if candidate.is_file() and candidate.stat().st_size == source_size:
            source_hash = source_hash or sha256_file(source)
            if sha256_file(candidate) == source_hash:
                return CopyResult(candidate, "ĐÃ TỒN TẠI")
    raise RuntimeError("Không thể chọn tên file kết quả.")


def is_valid_result_filename(filename: str, gcn_key: str) -> bool:
    """Kiểm tra tên kết quả có đúng dấu cách và hậu tố hay không."""

    display = re.escape(format_gcn_display(gcn_key))
    return bool(re.fullmatch(rf"{display}(?:_\d{{2,}})?\.[^.]+", filename))
