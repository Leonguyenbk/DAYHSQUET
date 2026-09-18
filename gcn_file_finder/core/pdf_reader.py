"""Đọc text layer và render PDF trực tiếp trong bộ nhớ."""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

import fitz
import numpy as np


@contextmanager
def open_pdf(path: str | Path):
    """Mở PDF và luôn đóng tài liệu, kể cả khi một trang bị lỗi."""

    document = fitz.open(Path(path))
    try:
        if document.needs_pass:
            raise PermissionError("PDF có mật khẩu và chưa được mở khóa.")
        yield document
    finally:
        document.close()


def extract_page_texts(document) -> Iterator[tuple[int, str]]:
    """Đọc text theo trang, đánh số từ 1."""

    for index in range(document.page_count):
        yield index + 1, document.load_page(index).get_text("text")


def render_page_bgr(document, page_index: int, dpi: int = 350) -> np.ndarray:
    """Render một trang thành mảng BGR, không tạo ảnh tạm."""

    page = document.load_page(page_index)
    pixmap = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False)
    rgb = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, 3)
    return rgb[:, :, ::-1].copy()
