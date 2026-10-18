"""Ghi báo cáo CSV và quản lý cache Gemini."""

from __future__ import annotations

import csv
import json
import os
from pathlib import Path
from typing import Any, Iterable

from .detect import DetectionResult
from .gemini import GeminiPageResult
from .pdfio import InspectedFile


REPORT_COLUMNS = [
    "File",
    "Trang",
    "Giấy chứng nhận",
    "Cách nhận diện",
    "Điểm",
    "Từ khoá thấy",
    "Ghi chú hoặc tiêu đề và lý do của Gemini",
]
# Tăng phiên bản khi schema đổi để không dùng lại cache từng lẫn mã phôi và số vào sổ.
CACHE_VERSION = 5

CERTIFICATE_COLUMNS = [
    "File nguồn",
    "File GCN",
    "Các trang",
    "Số giấy chứng nhận (mã phôi)",
    "Số vào sổ cấp GCN",
    "Ngày cấp GCN",
    "Tên chủ sử dụng hiện tại",
    "Ghi chú trích xuất",
]


def cache_key(content_hash: str, page: int, model: str) -> str:
    """Tạo khoá cache từ hash file, số trang và model."""

    return f"{content_hash}:{page}:{model}"


def load_cache(path: Path) -> dict[str, dict[str, Any]]:
    """Đọc cache; cache lỗi hoặc khác phiên bản được xem là rỗng."""

    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if payload.get("version") != CACHE_VERSION or not isinstance(payload.get("entries"), dict):
            return {}
        return payload["entries"]
    except (OSError, json.JSONDecodeError, AttributeError):
        return {}


def save_cache(path: Path, entries: dict[str, dict[str, Any]]) -> None:
    """Ghi cache JSON theo cách thay thế nguyên tử."""

    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    payload = {"version": CACHE_VERSION, "entries": entries}
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    os.replace(temporary, path)


def cached_gemini_result(value: dict[str, Any]) -> GeminiPageResult | None:
    """Kiểm tra và chuyển một mục cache thành kết quả Gemini."""

    try:
        index = value["i"]
        certificate = value["is_certificate"]
        title = value["title"]
        reason = value["reason"]
        if (
            not isinstance(index, int)
            or type(certificate) is not bool
            or not isinstance(title, str)
            or not isinstance(reason, str)
        ):
            return None
        return GeminiPageResult(index, certificate, title, reason)
    except (KeyError, TypeError):
        return None


def serialize_gemini_result(result: GeminiPageResult) -> dict[str, Any]:
    """Chuyển kết quả Gemini thành mục JSON."""

    return {
        "i": result.i,
        "is_certificate": result.is_certificate,
        "title": result.title,
        "reason": result.reason,
    }


def _result_row(relative_path: str, result: DetectionResult) -> dict[str, str | int]:
    """Tạo một dòng báo cáo cho trang PDF."""

    return {
        "File": Path(relative_path).as_posix(),
        "Trang": result.page,
        "Giấy chứng nhận": "Có" if result.is_certificate else "",
        "Cách nhận diện": result.method,
        "Điểm": result.score,
        "Từ khoá thấy": "; ".join(result.keywords),
        "Ghi chú hoặc tiêu đề và lý do của Gemini": result.note,
    }


def rows_for_file(inspected: InspectedFile) -> list[dict[str, str | int]]:
    """Chuyển kết quả một file thành các dòng CSV, kể cả lỗi file."""

    if inspected.error:
        return [
            {
                "File": Path(inspected.relative_path).as_posix(),
                "Trang": "",
                "Giấy chứng nhận": "",
                "Cách nhận diện": "",
                "Điểm": "",
                "Từ khoá thấy": "",
                "Ghi chú hoặc tiêu đề và lý do của Gemini": inspected.error,
            }
        ]
    rows = [_result_row(inspected.relative_path, page.result) for page in inspected.pages]
    if inspected.warning and rows:
        note_column = "Ghi chú hoặc tiêu đề và lý do của Gemini"
        current = str(rows[0][note_column])
        if current and inspected.warning.startswith(current):
            rows[0][note_column] = inspected.warning
        elif inspected.warning not in current:
            rows[0][note_column] = "; ".join(
                part for part in (current, inspected.warning) if part
            )
    return rows


def write_csv(path: Path, rows: Iterable[dict[str, str | int]]) -> None:
    """Ghi CSV UTF-8 có BOM để Microsoft Excel đọc đúng tiếng Việt."""

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_certificate_csv(path: Path, files: Iterable[InspectedFile]) -> None:
    """Ghi một dòng thông tin tổng hợp cho mỗi GCN đã gom trang."""

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=CERTIFICATE_COLUMNS)
        writer.writeheader()
        for inspected in files:
            for group in inspected.certificate_groups:
                writer.writerow(
                    {
                        "File nguồn": Path(inspected.relative_path).as_posix(),
                        "File GCN": group.output_file,
                        "Các trang": ", ".join(str(index + 1) for index in group.page_indexes),
                        "Số giấy chứng nhận (mã phôi)": group.certificate_serial,
                        "Số vào sổ cấp GCN": group.registration_number,
                        "Ngày cấp GCN": group.issue_date,
                        "Tên chủ sử dụng hiện tại": group.owner_name,
                        "Ghi chú trích xuất": group.note,
                    }
                )
