"""Đọc, sửa hướng và tạo biến thể ảnh phục vụ OCR."""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

import cv2
import numpy as np
from PIL import Image, ImageOps

OCR_MAX_SIDE = 4000


def limit_image_size(image: np.ndarray, max_side: int = OCR_MAX_SIDE) -> np.ndarray:
    """Downscale once before OCR so PaddleX does not resize every variant."""

    if max_side <= 0:
        raise ValueError("Kích thước cạnh tối đa phải lớn hơn 0.")
    height, width = image.shape[:2]
    largest_side = max(height, width)
    if largest_side <= max_side:
        return image
    scale = max_side / largest_side
    target = (max(1, round(width * scale)), max(1, round(height * scale)))
    return cv2.resize(image, target, interpolation=cv2.INTER_AREA)


def load_image(path: str | Path) -> np.ndarray:
    """Đọc ảnh có đường dẫn Unicode, áp dụng EXIF orientation và trả về BGR."""

    try:
        with Image.open(Path(path)) as image:
            corrected = ImageOps.exif_transpose(image).convert("RGB")
            rgb = np.asarray(corrected)
    except Exception as exc:
        raise ValueError(f"Ảnh hỏng hoặc không đọc được: {exc}") from exc
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def orientation_variants(image: np.ndarray) -> Iterator[np.ndarray]:
    """Sinh ảnh gốc rồi các hướng 90/180/270 độ để khôi phục ảnh quét xoay."""

    yield image
    yield cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
    yield cv2.rotate(image, cv2.ROTATE_180)
    yield cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)


def _red_mask(image: np.ndarray) -> np.ndarray:
    """Mặt nạ nhị phân các pixel màu đỏ (chữ đỏ, nền đỏ, dấu mộc đỏ...)."""

    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    mask1 = cv2.inRange(hsv, np.array([0, 45, 35]), np.array([12, 255, 255]))
    mask2 = cv2.inRange(hsv, np.array([168, 45, 35]), np.array([180, 255, 255]))
    return cv2.bitwise_or(mask1, mask2)


def whiten_red_background(image: np.ndarray) -> np.ndarray:
    """Biến mọi pixel màu đỏ thành trắng, giữ nguyên chữ đen, rồi chuyển xám.

    Ngược với trường hợp chữ đỏ trên nền trắng (``OCR TÁCH MÀU ĐỎ``): ở đây
    nền/dấu mộc là màu đỏ nhưng chữ vẫn đen, nên chuyển xám thẳng cho tương
    phản thấp vì đỏ vẫn còn đủ sáng. Xoá hẳn màu đỏ về trắng trước khi chuyển
    xám giúp chữ đen nổi bật rõ trên nền trắng sạch.
    """

    whitened = image.copy()
    whitened[_red_mask(image) > 0] = (255, 255, 255)
    return cv2.cvtColor(whitened, cv2.COLOR_BGR2GRAY)


def enhanced_gray(image: np.ndarray) -> np.ndarray:
    """Chuyển xám rồi tăng tương phản cục bộ (CLAHE).

    Rẻ (một phép chuyển đổi), nhưng thường đủ để đọc được chữ đen trên nền
    không đồng màu/hơi tối mà ảnh màu gốc OCR bỏ sót, kể cả khi không phải
    do màu đỏ.
    """

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8)).apply(gray)


def preprocessing_variants(image: np.ndarray, include_red: bool = True) -> Iterator[tuple[str, np.ndarray]]:
    """Sinh lần lượt các biến thể OCR theo yêu cầu nghiệp vụ."""

    yield "OCR ẢNH GỐC", image
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    yield "OCR ẢNH XÁM", gray
    contrast = enhanced_gray(image)
    yield "OCR ẢNH XÁM", contrast
    threshold = cv2.adaptiveThreshold(
        contrast, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11
    )
    yield "OCR THRESHOLD", threshold
    blurred = cv2.GaussianBlur(contrast, (0, 0), 2.0)
    sharpened = cv2.addWeighted(contrast, 1.8, blurred, -0.8, 0)
    yield "OCR LÀM NÉT", sharpened
    if include_red:
        red_text = cv2.bitwise_not(_red_mask(image))
        yield "OCR TÁCH MÀU ĐỎ", red_text
        yield "OCR NỀN ĐỎ THÀNH TRẮNG", whiten_red_background(image)
