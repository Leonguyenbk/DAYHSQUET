# -*- coding: utf-8 -*-
"""Giao diện PySide6 (Qt Widgets): tải hồ sơ quét theo đơn → ký số → đẩy lại thay thế đúng vị trí.

- Chuyển toàn bộ GUI sang Python + PySide6 (không phụ thuộc Flet hay trình duyệt).
- Tích hợp sẵn bộ mô phỏng & giải trực tiếp SliderCaptcha (278 × 155) bằng QPainter ngay trên giao diện:
  + Hiển thị ảnh thật từ GET /dc/SliderCaptcha/Challenge (hoặc tạo cảnh mô phỏng 278 × 155 khi thử nghiệm).
  + Kéo trực tiếp trên ảnh (chuột) hoặc kéo thanh trượt QSlider hoặc bấm '🤖 Tự động kéo & Giải'.
  + Luôn có sẵn nút '✓ Giải Captcha trên Tool & Tiếp tục' VÀ nút '✓ Đã giải xong Captcha (trên Chrome) – Tiếp tục'.
"""

from __future__ import annotations

import base64
import json
import math
import os
import random
import re
import subprocess
import sys
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from PySide6.QtCore import QObject, Qt, QTimer, Signal, Slot
from PySide6.QtGui import (
    QColor,
    QCursor,
    QFont,
    QImage,
    QLinearGradient,
    QMouseEvent,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
    QTextCursor,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSlider,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import tai_hsq_don as thd

APP_DIR = Path(thd.THU_MUC_APP)

LOGICAL_W, LOGICAL_H = 278, 155
SCALE = 2
STAGE_W, STAGE_H = LOGICAL_W * SCALE, LOGICAL_H * SCALE  # 556 x 310
BOX, PAD, CORE = 74, 16, 60
MIN_X, MAX_X = 0, STAGE_W - BOX
MAX_LOGICAL_LEFT = 241
# Verify trả false → tự kéo & giải lại tối đa ngần này lần, sau đó mở Chrome để giải tay
SO_LAN_TU_GIAI_CAPTCHA = 3


def jigsaw_path(x: float, y: float) -> QPainterPath:
    """Đường viền mảnh ghép Jigsaw chuẩn trên hệ toạ độ hiển thị 556x310 (mép trái khớp đúng tại x)."""
    p, s, tab = PAD, CORE, 9
    path = QPainterPath()
    path.moveTo(x, y + p)
    path.lineTo(x + 19, y + p)
    path.cubicTo(x + 17, y + p - tab, x + 23, y + 2, x + 30, y + 2)
    path.cubicTo(x + 37, y + 2, x + 43, y + p - tab, x + 41, y + p)
    path.lineTo(x + s, y + p)
    path.lineTo(x + s, y + p + 19)
    path.cubicTo(x + s + tab, y + p + 17, x + s + 12, y + p + 23, x + s + 12, y + p + 30)
    path.cubicTo(x + s + 12, y + p + 37, x + s + tab, y + p + 43, x + s, y + p + 41)
    path.lineTo(x + s, y + p + s)
    path.lineTo(x + 41, y + p + s)
    path.cubicTo(x + 43, y + p + s - tab, x + 37, y + p + s - 12, x + 30, y + p + s - 12)
    path.cubicTo(x + 23, y + p + s - 12, x + 17, y + p + s - tab, x + 19, y + p + s)
    path.lineTo(x, y + p + s)
    path.lineTo(x, y + p + 41)
    path.cubicTo(x + tab, y + p + 43, x + 9, y + p + 37, x + 9, y + p + 30)
    path.cubicTo(x + 9, y + p + 23, x + tab, y + p + 17, x, y + p + 19)
    path.closeSubpath()
    return path


def decode_b64_to_qimage(b64_str: str) -> QImage | None:
    if not b64_str or not isinstance(b64_str, str):
        return None
    s = b64_str.strip()
    if "," in s and s.startswith("data:"):
        s = s.split(",", 1)[1]
    s = re.sub(r"\s+", "", s)
    if not s:
        return None
    try:
        raw = base64.b64decode(s)
        img = QImage.fromData(raw)
        return img if not img.isNull() else None
    except Exception:
        return None


def detect_piece_left_pad_qimage(piece_img: QImage | None) -> int:
    """Tìm số pixel trong suốt bên mép trái của ảnh mảnh ghép piece (trên hệ toạ độ 1x)."""
    if piece_img is None or piece_img.isNull():
        return 0
    w, h = piece_img.width(), piece_img.height()
    for x in range(min(w, 24)):
        for y in range(h):
            if piece_img.pixelColor(x, y).alpha() > 45:
                return x
    return 0


def detect_cutout_left_qimage(img: QImage | None, piece_top: int, piece_pad_left: int = 0) -> int:
    """Dò cột X (0..241) của mép trái vết khuyết trên ảnh nền 278x155."""
    if img is None or img.isNull():
        return 115
    scaled = img.scaled(LOGICAL_W, LOGICAL_H, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
    y0 = max(2, min(LOGICAL_H - 12, int(piece_top) + 9))
    y1 = max(y0 + 6, min(LOGICAL_H - 2, int(piece_top) + 35))

    def lum(x: int, y: int) -> int:
        c = scaled.pixelColor(x, y)
        return (c.red() * 299 + c.green() * 587 + c.blue() * 114) // 1000

    best_x = 115
    best_score = -1.0
    for x in range(36, LOGICAL_W - 38):
        edge_score = 0.0
        dark_drop = 0.0
        for y in range(y0, y1):
            c = lum(x, y)
            prev1 = lum(x - 1, y)
            next1 = lum(x + 1, y)
            next2 = lum(x + 2, y)
            # Cạnh trái vết khuyết: cột x là viền sáng (c cao) chuyển sang lòng tối (next1/next2 thấp)
            # hoặc từ nền ngoài (prev1) sập xuống lòng tối (c/next1)
            edge_contrast = abs(c - prev1) + max(0, c - next1) * 1.6 + max(0, prev1 - next1) * 1.2
            edge_score += edge_contrast

            out_lum = (lum(max(0, x - 6), y) + lum(max(0, x - 3), y)) * 0.5
            in_lum = (lum(min(LOGICAL_W - 1, x + 5), y) + lum(min(LOGICAL_W - 1, x + 14), y)) * 0.5
            dark_drop += (out_lum - in_lum)

        # Chỉ chọn cạnh trái thực sự (bên phải cạnh phải tối hơn bên trái cạnh)
        if dark_drop <= 0:
            continue
        total_score = edge_score + dark_drop * 1.5
        if total_score > best_score:
            best_score = total_score
            best_x = x
    return max(0, min(MAX_LOGICAL_LEFT, best_x - max(0, int(piece_pad_left))))


def generate_adjacent_walk_datas(rng: random.Random, run_count: int | None = None) -> list[int]:
    """Sinh mảng datas (random walk từ -5..5, mỗi nấc lệch 1 và lặp 2..5 lần)."""
    values: list[int] = []
    current = 0
    runs = run_count or rng.randint(13, 18)
    for _ in range(runs):
        values.extend([current] * rng.randint(2, 5))
        if current <= -5:
            current += 1
        elif current >= 5:
            current -= 1
        else:
            current += rng.choice((-1, 1))
    return values


class CaptchaCanvasWidget(QWidget):
    """Widget vẽ trực tiếp SliderCaptcha (278x155 phóng 2x = 556x310) bằng QPainter + kéo chuột trực tiếp."""

    leftChanged = Signal(int, int)  # (logical_left 0..241, dy)
    dragStarted = Signal()
    dragFinished = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.rng = random.Random()
        self.setFixedSize(STAGE_W, STAGE_H)
        self.setCursor(QCursor(Qt.CursorShape.OpenHandCursor))

        self.is_mock = True
        self.mock_scene = QPixmap(STAGE_W, STAGE_H)
        self.mock_target_x = 260  # display coords (0..STAGE_W-BOX)
        self.top_px = 124         # display coords (pieceTop * 2)

        self.real_bg_pixmap: QPixmap | None = None
        self.real_piece_pixmap: QPixmap | None = None
        self.real_piece_full_strip = False
        self.hint_left = 115
        self.logical_left = 0

        self._dragging = False
        self._drag_start_x = 0.0
        self._drag_start_y = 0.0
        self._drag_start_left = 0

        self.generate_mock_scene()

    def generate_mock_scene(self) -> tuple[int, int]:
        """Sinh cảnh giả lập 278x155 (hiển thị 556x310) và trả về (target_left_logic, piece_top_logic)."""
        palettes = [
            ("#2876b7", "#c8e8f6", "#356a69", "#64a866"),
            ("#7554aa", "#f0afba", "#4d5c79", "#77966c"),
            ("#167a8b", "#cdedd6", "#2a674c", "#69a45b"),
        ]
        top, bottom, far, near = self.rng.choice(palettes)
        self.mock_scene = QPixmap(STAGE_W, STAGE_H)
        painter = QPainter(self.mock_scene)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        gradient = QLinearGradient(0, 0, 0, STAGE_H)
        gradient.setColorAt(0, QColor(top))
        gradient.setColorAt(1, QColor(bottom))
        painter.fillRect(0, 0, STAGE_W, STAGE_H, gradient)

        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(255, 228, 120, 220))
        painter.drawEllipse(self.rng.randint(60, 460), self.rng.randint(30, 75), 64, 64)
        painter.setBrush(QColor(250, 253, 255, 205))
        for _ in range(6):
            x, y = self.rng.randint(-20, 490), self.rng.randint(30, 120)
            painter.drawEllipse(x, y + 10, 75, 26)
            painter.drawEllipse(x + 19, y, 38, 38)

        far_path = QPainterPath()
        far_path.moveTo(0, STAGE_H)
        far_path.lineTo(0, 205)
        for x in range(-40, STAGE_W + 81, 95):
            far_path.lineTo(x, 205)
            far_path.lineTo(x + 48, self.rng.randint(125, 175))
            far_path.lineTo(x + 95, 205)
        far_path.lineTo(STAGE_W, STAGE_H)
        far_path.closeSubpath()
        painter.fillPath(far_path, QColor(far))

        near_path = QPainterPath()
        near_path.moveTo(0, STAGE_H)
        near_path.lineTo(0, 245)
        for x in range(-70, STAGE_W + 131, 130):
            near_path.cubicTo(
                x + 35,
                self.rng.randint(180, 220),
                x + 90,
                self.rng.randint(180, 225),
                x + 130,
                245,
            )
        near_path.lineTo(STAGE_W, STAGE_H)
        near_path.closeSubpath()
        painter.fillPath(near_path, QColor(near))

        tree_colours = [QColor("#176143"), QColor("#277849"), QColor("#557e3e")]
        painter.setPen(QPen(QColor("#49382b"), 3))
        for _ in range(11):
            x, y = self.rng.randint(10, STAGE_W - 10), self.rng.randint(240, STAGE_H)
            painter.drawLine(x, y, x, y - 20)
            radius = self.rng.randint(8, 13)
            painter.setBrush(self.rng.choice(tree_colours))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(x - radius, y - 25 - radius, radius * 2, radius * 2)
            painter.setPen(QPen(QColor("#49382b"), 3))
        painter.end()

        target_left = self.rng.randint(95, 210)
        piece_top = self.rng.randint(28, 96)
        self.is_mock = True
        self.real_bg_pixmap = None
        self.real_piece_pixmap = None
        self.mock_target_x = target_left * SCALE
        self.top_px = piece_top * SCALE
        self.hint_left = target_left
        self.logical_left = 0
        self.update()
        return target_left, piece_top

    def load_challenge_dict(self, ch_data: dict[str, Any]) -> bool:
        """Nạp dữ liệu từ /dc/SliderCaptcha/Challenge. Trả về True nếu là ảnh thật từ server."""
        p_top = core.safe_int(ch_data.get("pieceTop"), 62)
        bg_qimg = decode_b64_to_qimage(str(ch_data.get("bg") or ""))
        piece_qimg = decode_b64_to_qimage(str(ch_data.get("piece") or ""))

        if bg_qimg is None:
            target_left, piece_top = self.generate_mock_scene()
            ch_data["_mock_target_left"] = target_left
            ch_data["pieceTop"] = piece_top
            return False

        self.is_mock = "_mock_target_left" in ch_data
        self.top_px = p_top * SCALE
        self.real_bg_pixmap = QPixmap.fromImage(
            bg_qimg.scaled(STAGE_W, STAGE_H, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
        )

        if piece_qimg is not None:
            pw, ph = piece_qimg.width(), piece_qimg.height()
            self.real_piece_full_strip = abs(ph - LOGICAL_H) <= 6
            disp_pw = max(20, pw * SCALE)
            disp_ph = STAGE_H if self.real_piece_full_strip else max(20, ph * SCALE)
            self.real_piece_pixmap = QPixmap.fromImage(
                piece_qimg.scaled(disp_pw, disp_ph, Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
            )
        else:
            self.real_piece_pixmap = None
            self.real_piece_full_strip = False

        piece_pad_left = detect_piece_left_pad_qimage(piece_qimg)
        if "_mock_target_left" in ch_data:
            self.hint_left = int(ch_data["_mock_target_left"])
        else:
            self.hint_left = detect_cutout_left_qimage(bg_qimg, p_top, piece_pad_left)

        self.logical_left = 0
        self.update()
        return not self.is_mock

    def set_logical_left(self, left_val: int) -> None:
        self.logical_left = max(0, min(MAX_LOGICAL_LEFT, int(left_val)))
        self.update()

    def paintEvent(self, _event) -> None:  # type: ignore[override]
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        piece_x = float(self.logical_left * SCALE)

        if self.real_bg_pixmap is not None:
            painter.drawPixmap(0, 0, self.real_bg_pixmap)
            # Vạch ngang tại pieceTop và vạch xanh gợi ý vết khuyết
            painter.setPen(QPen(QColor(245, 158, 11, 95), 1, Qt.PenStyle.DashLine))
            painter.drawLine(0, self.top_px + 36, STAGE_W, self.top_px + 36)

            hx = self.hint_left * SCALE
            painter.setPen(QPen(QColor(16, 185, 129, 130), 2, Qt.PenStyle.DotLine))
            painter.drawLine(hx, 0, hx, STAGE_H)

            if self.real_piece_pixmap is not None:
                py = 0 if self.real_piece_full_strip else self.top_px
                painter.drawPixmap(round(piece_x), round(py), self.real_piece_pixmap)
            else:
                piece = jigsaw_path(piece_x, float(self.top_px))
                painter.fillPath(piece, QColor(0, 240, 255, 60))
                painter.setPen(QPen(QColor(255, 255, 255, 240), 2.5))
                painter.drawPath(piece)

            painter.setPen(QPen(QColor(0, 240, 255, 195), 2))
            painter.drawLine(round(piece_x), 0, round(piece_x), STAGE_H)
        else:
            painter.drawPixmap(0, 0, self.mock_scene)
            hole = jigsaw_path(float(self.mock_target_x), float(self.top_px))
            painter.fillPath(hole, QColor(3, 7, 18, 185))
            painter.setPen(QPen(QColor(245, 248, 252, 230), 3))
            painter.drawPath(hole)

            piece = jigsaw_path(piece_x, float(self.top_px))
            painter.save()
            painter.setClipPath(piece)
            painter.drawPixmap(round(piece_x - self.mock_target_x), 0, self.mock_scene)
            painter.restore()
            painter.setPen(QPen(QColor("white"), 3))
            painter.drawPath(piece)

        painter.end()

    def mousePressEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self.setCursor(QCursor(Qt.CursorShape.ClosedHandCursor))
            pos = event.position()
            self._drag_start_x = pos.x()
            self._drag_start_y = pos.y()
            clicked_left = int(round(pos.x() / SCALE - 18))
            self._drag_start_left = max(0, min(MAX_LOGICAL_LEFT, clicked_left))
            self.set_logical_left(self._drag_start_left)
            self.dragStarted.emit()
            self.leftChanged.emit(self.logical_left, 0)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        if not self._dragging:
            return
        pos = event.position()
        dx = (pos.x() - self._drag_start_x) / SCALE
        dy = max(-5, min(5, int(round((pos.y() - self._drag_start_y) / 4.0))))
        new_left = max(0, min(MAX_LOGICAL_LEFT, int(round(self._drag_start_left + dx))))
        self.set_logical_left(new_left)
        self.leftChanged.emit(self.logical_left, dy)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:  # type: ignore[override]
        if event.button() == Qt.MouseButton.LeftButton and self._dragging:
            self._dragging = False
            self.setCursor(QCursor(Qt.CursorShape.OpenHandCursor))
            self.dragFinished.emit()


class CaptchaPanelWidget(QFrame):
    """Khung giải SliderCaptcha + xác nhận hoàn thành nhúng trực tiếp trên cửa sổ chính PySide6."""

    def __init__(self, app_window: "TaiHsqPySide6Window") -> None:
        super().__init__(app_window)
        self.app_win = app_window
        self.rng = random.Random()
        self.setObjectName("captchaCard")

        self.session = None
        self.chuc_nang = "HoSoQuetAjax/UpdateHoSoQuetExistFile"
        self.challenge: dict[str, Any] = {}
        self.xong_event: threading.Event | None = None
        self.kq_dict: dict[str, Any] | None = None
        self.verifying = False
        self.so_lan_giai_hong = 0
        self.started_t = time.monotonic()
        self.drag_datas: list[int] = []
        self.duration_ms = 2084

        self.agent_plan: list[int] = []
        self.agent_idx = 0
        self.agent_timer = QTimer(self)
        self.agent_timer.setInterval(30)
        self.agent_timer.timeout.connect(self._agent_tick)

        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        # Dòng tiêu đề nổi bật
        header_row = QHBoxLayout()
        title = QLabel("🧩 CẦN GIẢI CAPTCHA ĐỂ TIẾP TỤC — Kéo mảnh ghép bên dưới hoặc bấm Xác nhận đã hoàn thành")
        title.setObjectName("captchaTitle")
        header_row.addWidget(title, 1)
        layout.addLayout(header_row)

        self.lbl_info = QLabel("API: HoSoQuetAjax/UpdateHoSoQuetExistFile")
        self.lbl_info.setObjectName("monoCyan")
        layout.addWidget(self.lbl_info)

        # Hàng nút hành động chính đặt ngay trên đầu để người dùng luôn thấy ngay lập tức
        action_row = QHBoxLayout()
        action_row.setSpacing(8)

        self.btn_verify = QPushButton("✓ Giải Captcha trên Tool & Tiếp tục")
        self.btn_verify.setObjectName("btnSuccess")
        self.btn_verify.clicked.connect(self.verify_captcha)

        self.btn_auto_agent = QPushButton("🤖 Tự động kéo & Giải (4–6s)")
        self.btn_auto_agent.setObjectName("btnPurple")
        self.btn_auto_agent.clicked.connect(self.start_auto_agent)

        self.btn_open_chrome = QPushButton("🌐 Mở Chrome để giải")
        self.btn_open_chrome.setObjectName("btnBlue")
        self.btn_open_chrome.clicked.connect(lambda: self.app_win._mo_hoac_dua_chrome_len())

        self.btn_confirm_chrome = QPushButton("✓ Đã giải xong Captcha (trên Chrome) – Tiếp tục")
        self.btn_confirm_chrome.setObjectName("btnPrimary")
        self.btn_confirm_chrome.clicked.connect(lambda: self.finish_captcha(True, "▶ Đã xác nhận hoàn thành Captcha — tiếp tục chạy."))

        self.btn_stop_captcha = QPushButton("⏹ Dừng chạy")
        self.btn_stop_captcha.setObjectName("btnDanger")
        self.btn_stop_captcha.clicked.connect(lambda: self.finish_captcha(False, "⏹ Đã dừng tại bước giải Captcha."))

        action_row.addWidget(self.btn_open_chrome)
        action_row.addWidget(self.btn_confirm_chrome)
        action_row.addWidget(self.btn_verify)
        action_row.addWidget(self.btn_auto_agent)
        action_row.addWidget(self.btn_stop_captcha)
        action_row.addStretch()
        layout.addLayout(action_row)

        self.lbl_status = QLabel(
            "Cách 1: Kéo thanh trượt (hoặc bấm '🤖 Tự động kéo & Giải') rồi bấm '✓ Giải Captcha trên Tool & Tiếp tục'.\n"
            "Cách 2: Nếu bạn giải trên trình duyệt Chrome, bấm ngay '✓ Đã giải xong Captcha (trên Chrome) – Tiếp tục'."
        )
        self.lbl_status.setObjectName("statusText")
        self.lbl_status.setWordWrap(True)
        layout.addWidget(self.lbl_status)

        # Khung ảnh Captcha 556x310
        canvas_wrap = QHBoxLayout()
        canvas_wrap.addStretch()
        self.canvas = CaptchaCanvasWidget(self)
        self.canvas.dragStarted.connect(self._on_drag_started)
        self.canvas.leftChanged.connect(self._on_canvas_dragged)
        self.canvas.dragFinished.connect(self._on_drag_finished)
        canvas_wrap.addWidget(self.canvas)
        canvas_wrap.addStretch()
        layout.addLayout(canvas_wrap)

        # Thanh kéo QSlider (0..241 px logic), căn đúng chiều rộng 556px khớp với khung ảnh Captcha
        slider_wrap = QHBoxLayout()
        slider_wrap.addStretch()
        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setFixedWidth(STAGE_W)
        self.slider.setRange(0, MAX_LOGICAL_LEFT)
        self.slider.setSingleStep(1)
        self.slider.sliderPressed.connect(self._on_drag_started)
        self.slider.valueChanged.connect(self._on_slider_changed)
        self.slider.sliderReleased.connect(self._on_drag_finished)
        slider_wrap.addWidget(self.slider)
        slider_wrap.addStretch()
        layout.addLayout(slider_wrap)

        # Hàng nút tinh chỉnh toạ độ
        nudge_row = QHBoxLayout()
        self.lbl_metric = QLabel("left: 0 px  |  gợi ý: ~115 px  |  duration: 0 ms")
        self.lbl_metric.setObjectName("monoAmber")
        nudge_row.addWidget(self.lbl_metric, 1)

        for label, delta in (("-5", -5), ("-1px", -1), ("+1px", 1), ("+5", 5)):
            btn = QPushButton(label)
            btn.setFixedWidth(48)
            btn.clicked.connect(lambda _checked=False, d=delta: self.nudge_left(d))
            nudge_row.addWidget(btn)

        btn_hint = QPushButton("⚡ Đặt vào vết khuyết")
        btn_hint.clicked.connect(self.snap_to_hint)
        nudge_row.addWidget(btn_hint)

        btn_reload = QPushButton("🔄 Đổi hình mới")
        btn_reload.clicked.connect(self.reload_challenge)
        nudge_row.addWidget(btn_reload)
        layout.addLayout(nudge_row)

        # Hàng tuỳ chọn tự động gửi + dán JSON thủ công
        extra_row = QHBoxLayout()
        self.chk_auto_submit = QCheckBox("Tự gửi Verify ngay khi thả thanh trượt")
        self.chk_auto_submit.setChecked(False)
        extra_row.addWidget(self.chk_auto_submit)

        self.inp_json = QLineEdit()
        self.inp_json.setPlaceholderText('Hoặc dán JSON từ /dc/SliderCaptcha/Challenge (tuỳ chọn)...')
        extra_row.addWidget(self.inp_json, 1)

        btn_load_json = QPushButton("Nạp JSON")
        btn_load_json.clicked.connect(self.load_manual_json)
        extra_row.addWidget(btn_load_json)
        layout.addLayout(extra_row)

    def open_for_challenge(
        self,
        session,
        chuc_nang: str,
        challenge: dict[str, Any] | None,
        xong: threading.Event,
        kq: dict[str, Any],
    ) -> None:
        self.agent_timer.stop()
        self.session = session
        self.chuc_nang = chuc_nang or "HoSoQuetAjax/UpdateHoSoQuetExistFile"
        self.xong_event = xong
        self.kq_dict = kq
        self.verifying = False
        self.so_lan_giai_hong = 0
        self.apply_challenge_data(challenge or {})
        self.setVisible(True)

    def apply_challenge_data(self, ch_data: dict[str, Any]) -> None:
        self.challenge = dict(ch_data or {})
        is_real = self.canvas.load_challenge_dict(self.challenge)
        self.started_t = time.monotonic()
        self.duration_ms = self.rng.randint(1850, 2900)
        self.drag_datas = generate_adjacent_walk_datas(self.rng)
        self.slider.blockSignals(True)
        self.slider.setValue(0)
        self.slider.blockSignals(False)

        cid = str(self.challenge.get("challengeId") or "")[:12]
        p_top = core.safe_int(self.challenge.get("pieceTop"), 62)
        mode_txt = "ẢNH THẬT TỪ MPLIS" if is_real else "MÔ PHỎNG OFFLINE 278×155"
        self.lbl_info.setText(
            f"API: {self.chuc_nang}  •  challengeId: {cid or 'local'}…  •  pieceTop: {p_top}px  •  [{mode_txt}]"
        )
        self._update_metric_label()

    def _update_metric_label(self) -> None:
        self.lbl_metric.setText(
            f"left: {self.canvas.logical_left} px  |  gợi ý vết khuyết: ~{self.canvas.hint_left} px  |  "
            f"duration: {self.duration_ms} ms  |  datas: {len(self.drag_datas)} mẫu"
        )

    def _on_drag_started(self) -> None:
        self.agent_timer.stop()
        self.started_t = time.monotonic()
        self.drag_datas = [0]

    def _on_canvas_dragged(self, left_val: int, dy: int) -> None:
        self.duration_ms = max(140, int((time.monotonic() - self.started_t) * 1000))
        self.drag_datas.append(dy)
        self.slider.blockSignals(True)
        self.slider.setValue(left_val)
        self.slider.blockSignals(False)
        self._update_metric_label()

    def _on_slider_changed(self, value: int) -> None:
        self.canvas.set_logical_left(value)
        if not self.agent_timer.isActive():
            self.duration_ms = max(140, int((time.monotonic() - self.started_t) * 1000))
            last_y = self.drag_datas[-1] if self.drag_datas else 0
            if self.rng.random() < 0.3:
                last_y = max(-5, min(5, last_y + self.rng.choice((-1, 1))))
            self.drag_datas.append(last_y)
        self._update_metric_label()

    def _on_drag_finished(self) -> None:
        self.duration_ms = max(260, int((time.monotonic() - self.started_t) * 1000))
        if len(self.drag_datas) < 16 or all(x == 0 for x in self.drag_datas):
            self.drag_datas = generate_adjacent_walk_datas(self.rng)
        self._update_metric_label()
        if self.chk_auto_submit.isChecked():
            self.verify_captcha()

    def nudge_left(self, delta: int) -> None:
        self.agent_timer.stop()
        new_left = max(0, min(MAX_LOGICAL_LEFT, self.canvas.logical_left + delta))
        self.canvas.set_logical_left(new_left)
        self.slider.blockSignals(True)
        self.slider.setValue(new_left)
        self.slider.blockSignals(False)
        if len(self.drag_datas) < 16:
            self.drag_datas = generate_adjacent_walk_datas(self.rng)
        self._update_metric_label()

    def snap_to_hint(self) -> None:
        self.agent_timer.stop()
        self.duration_ms = self.rng.randint(1850, 2950)
        self.drag_datas = generate_adjacent_walk_datas(self.rng)
        self.canvas.set_logical_left(self.canvas.hint_left)
        self.slider.blockSignals(True)
        self.slider.setValue(self.canvas.hint_left)
        self.slider.blockSignals(False)
        self._update_metric_label()
        self.lbl_status.setText(
            f"⚡ Đã đặt mảnh ghép tới vị trí gợi ý left = {self.canvas.logical_left}px. "
            "Bấm '✓ Giải Captcha trên Tool & Tiếp tục'!"
        )

    def start_auto_agent(self) -> None:
        """Mô phỏng kéo ngẫu nhiên không đều 4–6 giây tới đúng vị trí vết khuyết rồi tự động Verify."""
        if self.verifying:
            return
        self.agent_timer.stop()
        start_left = self.canvas.logical_left
        target_left = self.canvas.hint_left
        self.duration_ms = self.rng.randint(4000, 6000)
        self.drag_datas = generate_adjacent_walk_datas(self.rng, self.rng.randint(14, 19))

        frame_count = max(2, math.ceil(self.duration_ms / 30))
        weights: list[float] = []
        segment_factor, frames_left = 1.0, 0
        for index in range(frame_count):
            phase = index / (frame_count - 1)
            if frames_left <= 0:
                if phase < 0.62:
                    segment_factor, frames_left = 0.45 + self.rng.random() * 2.25, self.rng.randint(3, 12)
                elif phase < 0.82:
                    segment_factor, frames_left = 0.18 + self.rng.random() * 1.05, self.rng.randint(3, 10)
                else:
                    segment_factor, frames_left = 0.03 + self.rng.random() * 0.48, self.rng.randint(2, 8)
            frames_left -= 1
            base = 1.45 if phase < 0.55 else 0.72 if phase < 0.78 else 0.22
            weight = base * segment_factor * (0.72 + self.rng.random() * 0.58)
            hold_chance = 0.05 if phase < 0.62 else 0.14 if phase < 0.82 else 0.28
            weights.append(0.0 if self.rng.random() < hold_chance else weight)

        total = sum(weights) or 1.0
        cumulative = 0.0
        travel = target_left - start_left
        phase_offset = self.rng.random() * math.tau
        plan: list[int] = []
        for index, weight in enumerate(weights):
            cumulative += weight
            progress = min(1.0, cumulative / total)
            correction = 0.0
            if 0.76 < progress < 1.0:
                fade = min(1.0, (1.0 - progress) / 0.24)
                correction = math.sin(index * 0.58 + phase_offset) * (1.2 + self.rng.random() * 2.5) * fade
            plan.append(max(0, min(MAX_LOGICAL_LEFT, round(start_left + travel * progress + correction))))
        plan[-1] = max(0, min(MAX_LOGICAL_LEFT, target_left))

        self.agent_plan = plan
        self.agent_idx = 0
        self.started_t = time.monotonic()
        self.lbl_status.setText(f"🤖 Đang kéo mô phỏng tự nhiên ({self.duration_ms / 1000:.1f}s) tới left = {target_left}px…")
        self.agent_timer.start()

    def _agent_tick(self) -> None:
        if self.agent_idx >= len(self.agent_plan):
            self.agent_timer.stop()
            self.canvas.set_logical_left(self.canvas.hint_left)
            self.slider.blockSignals(True)
            self.slider.setValue(self.canvas.hint_left)
            self.slider.blockSignals(False)
            self._update_metric_label()
            self.verify_captcha()
            return
        val = self.agent_plan[self.agent_idx]
        self.agent_idx += 1
        self.canvas.set_logical_left(val)
        self.slider.blockSignals(True)
        self.slider.setValue(val)
        self.slider.blockSignals(False)
        self.duration_ms = max(200, int((time.monotonic() - self.started_t) * 1000))
        self._update_metric_label()

    def reload_challenge(self) -> None:
        if self.verifying:
            return
        self.agent_timer.stop()
        self.lbl_status.setText("🔄 Đang tải hình Captcha mới...")

        def _worker():
            new_ch = None
            if self.session is not None:
                try:
                    new_ch = core.lay_challenge_slider_captcha(self.session, self.chuc_nang)
                except Exception as exc:
                    thd.log(f"   ⚠ Không tải được Challenge mới từ server: {core.rut_gon_text(exc, 120)}")
            self.app_win.bridge.captchaReloaded.emit(new_ch or {})

        threading.Thread(target=_worker, daemon=True).start()

    def load_manual_json(self) -> None:
        raw_txt = self.inp_json.text().strip()
        if not raw_txt:
            return
        try:
            obj = json.loads(raw_txt)
            if isinstance(obj.get("data"), dict):
                obj = obj["data"]
            self.apply_challenge_data(obj)
            self.lbl_status.setText("✅ Đã nạp JSON Challenge thủ công lên khung vẽ!")
        except Exception as exc:
            self.lbl_status.setText(f"❌ JSON không hợp lệ: {core.rut_gon_text(exc, 100)}")

    def verify_captcha(self) -> None:
        if self.verifying:
            return
        self.agent_timer.stop()
        if len(self.drag_datas) < 12 or all(x == 0 for x in self.drag_datas):
            self.drag_datas = generate_adjacent_walk_datas(self.rng)

        payload = {
            "datas": self.drag_datas,
            "left": int(self.canvas.logical_left),
            "duration": int(self.duration_ms or 2084),
            "challengeId": str(self.challenge.get("challengeId") or uuid.uuid4().hex),
            "functionId": str(self.challenge.get("functionId") or self.chuc_nang or "HoSoQuetAjax/UpdateHoSoQuetExistFile"),
        }

        if self.canvas.is_mock or self.session is None:
            diff = abs(payload["left"] - self.canvas.hint_left)
            if diff <= 4:
                self.finish_captcha(
                    True,
                    f"✅ [Mô phỏng PySide6] Giải Captcha khớp chính xác (left={payload['left']}px, duration={payload['duration']}ms, lệch {diff}px)!",
                )
            else:
                direction = "phải" if payload["left"] < self.canvas.hint_left else "trái"
                self.lbl_status.setText(
                    f"⚠ Chưa khớp (left={payload['left']}px, vết khuyết ở {self.canvas.hint_left}px — dịch sang {direction}). "
                    "Hãy kéo chỉnh lại hoặc bấm '✓ Đã giải xong Captcha (trên Chrome) – Tiếp tục'!"
                )
            return

        self.verifying = True
        self.lbl_status.setText(f"⏳ Đang gửi POST /dc/SliderCaptcha/Verify (left={payload['left']}px, duration={payload['duration']}ms)...")

        def _worker():
            try:
                res = core.xac_minh_slider_captcha(self.session, payload)
                if isinstance(res, dict) and res.get("verified"):
                    token = str(res.get("recaptchaResponse") or "")
                    self.app_win.bridge.captchaVerifyResult.emit(
                        True,
                        f"✅ Giải SliderCaptcha thành công ngay trên tool (left={payload['left']}px, duration={payload['duration']}ms, recaptchaResponse={token[:12]}...)",
                        {},
                    )
                    return
                new_ch = core.lay_challenge_slider_captcha(self.session, self.chuc_nang)
                self.app_win.bridge.captchaVerifyResult.emit(
                    False,
                    "⚠ Vị trí chưa khớp — đã tự tải hình Captcha mới, hãy kéo lại hoặc bấm '✓ Đã giải xong Captcha – Tiếp tục'!",
                    new_ch or {},
                )
            except Exception as exc:
                self.app_win.bridge.captchaVerifyResult.emit(
                    False,
                    f"❌ Lỗi gửi Verify ({core.rut_gon_text(exc, 140)}) — bạn có thể giải trên Chrome rồi bấm '✓ Đã giải xong Captcha – Tiếp tục'!",
                    {},
                )

        threading.Thread(target=_worker, daemon=True).start()

    def finish_captcha(self, tiep: bool, log_msg: str = "") -> None:
        self.agent_timer.stop()
        self.verifying = False
        if log_msg:
            self.app_win._append_log(log_msg)
        if self.kq_dict is not None:
            self.kq_dict["tiep"] = tiep
        self.setVisible(False)
        if self.xong_event is not None:
            self.xong_event.set()


class UiBridge(QObject):
    logSignal = Signal(str)
    progressSignal = Signal(int, int, str)
    excelTaiSignal = Signal(str)
    excelKqSignal = Signal(str)
    loginSignal = Signal(str, str, bool)
    captchaSignal = Signal(object, str, object, object, object)
    captchaHideSignal = Signal()
    captchaReloaded = Signal(dict)
    captchaVerifyResult = Signal(bool, str, dict)
    doneSignal = Signal(str)


class TaiHsqPySide6Window(QMainWindow):
    WINDOW_TITLE = "MPLIS • Tải / đẩy hồ sơ quét ký số (PySide6)"

    O_CHO = (
        ("Chờ giữa 2 request (giây)", "DELAY_GIUA_REQUEST"),
        ("Chờ giữa 2 đơn (giây)", "DELAY_GIUA_DON"),
    )

    def __init__(self) -> None:
        super().__init__()
        self.is_running = False
        self.stop_flag = False
        self.login_client: core.MplisClient | None = None
        self.o_cho_inputs: list[tuple[QLineEdit, str]] = []

        self.bridge = UiBridge(self)
        self.bridge.logSignal.connect(self._append_log)
        self.bridge.progressSignal.connect(self._on_progress)
        self.bridge.excelTaiSignal.connect(self._on_excel_tai)
        self.bridge.loginSignal.connect(self._on_login_status)
        self.bridge.captchaSignal.connect(self._on_show_captcha)
        self.bridge.captchaHideSignal.connect(self._on_hide_captcha)
        self.bridge.captchaReloaded.connect(self._on_captcha_reloaded)
        self.bridge.captchaVerifyResult.connect(self._on_captcha_verify_result)
        self.bridge.doneSignal.connect(self._on_done)

        self.setWindowTitle(self.WINDOW_TITLE)
        self.setMinimumSize(1100, 720)
        self.resize(1280, 840)

        self._build_main_ui()
        self._apply_stylesheet()

        thd.dat_ham_log(lambda msg: self.bridge.logSignal.emit(str(msg)))
        core.CHO_GIAI_CAPTCHA = self._cho_giai_captcha
        mau = thd.doc_mau_da_luu()
        self._append_log("Endpoint tải: " + (mau["url"] if mau else thd.MAU_MAC_DINH[0]["url"]))

    def _build_main_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(14)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        root_layout.addWidget(splitter, 1)

        # Cột trái: Phiên đăng nhập + Phần chức năng
        left_widget = QWidget()
        left_widget.setMinimumWidth(450)
        left_widget.setMaximumWidth(560)
        left_col = QVBoxLayout(left_widget)
        left_col.setContentsMargins(0, 0, 0, 0)
        left_col.setSpacing(12)

        left_col.addWidget(self._build_session_card())

        body_card = QFrame()
        body_card.setObjectName("card")
        body_layout = QVBoxLayout(body_card)
        body_layout.setContentsMargins(12, 12, 12, 12)
        body_layout.addWidget(self._build_body(), 1)
        left_col.addWidget(body_card, 1)

        splitter.addWidget(left_widget)
        splitter.addWidget(self._build_log_card())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

    def _build_session_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title = QLabel("🔑 Phiên đăng nhập MPLIS")
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        mode_row = QHBoxLayout()
        self.rad_login = QRadioButton("Đăng nhập (Chrome)")
        self.rad_cookie = QRadioButton("Dán token + cookie")
        self.rad_login.setChecked(True)
        self.rad_login.toggled.connect(self._doi_che_do)
        mode_row.addWidget(self.rad_login)
        mode_row.addWidget(self.rad_cookie)
        mode_row.addStretch()
        layout.addLayout(mode_row)

        # Panel Chrome
        self.panel_login = QWidget()
        login_layout = QVBoxLayout(self.panel_login)
        login_layout.setContentsMargins(0, 0, 0, 0)
        login_layout.setSpacing(6)

        up_row = QHBoxLayout()
        self.inp_username = QLineEdit()
        self.inp_username.setPlaceholderText("Username MPLIS")
        self.inp_password = QLineEdit()
        self.inp_password.setPlaceholderText("Password")
        self.inp_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_password.returnPressed.connect(self._mo_chrome)
        up_row.addWidget(self.inp_username, 1)
        up_row.addWidget(self.inp_password, 1)
        login_layout.addLayout(up_row)

        btn_row = QHBoxLayout()
        self.btn_chrome = QPushButton("🌐 1. Mở Chrome")
        self.btn_chrome.setObjectName("btnBlue")
        self.btn_chrome.clicked.connect(self._mo_chrome)
        self.btn_lay_session = QPushButton("🔑 2. Lấy session")
        self.btn_lay_session.setObjectName("btnSuccess")
        self.btn_lay_session.clicked.connect(self._lay_session_chrome)
        btn_row.addWidget(self.btn_chrome)
        btn_row.addWidget(self.btn_lay_session)
        btn_row.addStretch()
        login_layout.addLayout(btn_row)

        self.lbl_login_status = QLabel("Chưa đăng nhập")
        self.lbl_login_status.setObjectName("mutedText")
        login_layout.addWidget(self.lbl_login_status)
        layout.addWidget(self.panel_login)

        # Panel Cookie
        self.panel_cookie = QWidget()
        self.panel_cookie.setVisible(False)
        cookie_layout = QVBoxLayout(self.panel_cookie)
        cookie_layout.setContentsMargins(0, 0, 0, 0)
        cookie_layout.setSpacing(6)

        self.inp_token = QLineEdit()
        self.inp_token.setPlaceholderText("__RequestVerificationToken (F12 > Network > Request Headers)")
        self.inp_cookie = QPlainTextEdit()
        self.inp_cookie.setPlaceholderText("Dán nguyên dòng Cookie từ Request Headers")
        self.inp_cookie.setFixedHeight(62)
        cookie_layout.addWidget(self.inp_token)
        cookie_layout.addWidget(self.inp_cookie)
        layout.addWidget(self.panel_cookie)

        # Hàng chỉnh thời gian chờ + nút Mở thử Captcha
        delay_row = QHBoxLayout()
        self.o_cho_inputs.clear()
        for label, ten_bien in self.O_CHO:
            col = QVBoxLayout()
            lbl = QLabel(label)
            lbl.setObjectName("mutedText")
            inp = QLineEdit(core.ghi_khoang_giay(getattr(core, ten_bien)))
            inp.textChanged.connect(lambda txt, tb=ten_bien, widget=inp: self._on_delay_changed(txt, tb, widget))
            col.addWidget(lbl)
            col.addWidget(inp)
            delay_row.addLayout(col, 1)
            self.o_cho_inputs.append((inp, ten_bien))

        btn_test_captcha = QPushButton("🧩 Mở thử Captcha")
        btn_test_captcha.setObjectName("btnPurple")
        btn_test_captcha.clicked.connect(self._mo_thu_captcha)
        delay_row.addWidget(btn_test_captcha, 0, Qt.AlignmentFlag.AlignBottom)
        layout.addLayout(delay_row)

        note = QLabel("Khi MPLIS hỏi Captcha, khung kéo giải & nút xác nhận sẽ hiện ngay bên phải trên giao diện này.")
        note.setObjectName("mutedText")
        note.setWordWrap(True)
        layout.addWidget(note)

        return card

    def _on_delay_changed(self, txt: str, ten_bien: str, widget: QLineEdit) -> None:
        try:
            setattr(core, ten_bien, core.doc_khoang_giay(txt))
            widget.setStyleSheet("")
        except ValueError:
            widget.setStyleSheet("border: 1px solid #ef4444;")

    def _build_log_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        top_row = QHBoxLayout()
        title = QLabel("🖥 Nhật ký & Giải Captcha trực tiếp")
        title.setObjectName("cardTitle")
        top_row.addWidget(title)

        self.lbl_status = QLabel("Sẵn sàng")
        self.lbl_status.setObjectName("monoCyan")
        top_row.addWidget(self.lbl_status, 1)

        self.btn_stop = QPushButton("⏹ Dừng")
        self.btn_stop.setObjectName("btnDanger")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop)
        top_row.addWidget(self.btn_stop)

        btn_excel = QPushButton("📊 Mở Excel KQ")
        btn_excel.clicked.connect(self._mo_excel)
        top_row.addWidget(btn_excel)

        btn_folder = QPushButton("📂 Mở thư mục")
        btn_folder.clicked.connect(self._mo_thu_muc)
        top_row.addWidget(btn_folder)

        btn_clear = QPushButton("🧹 Xóa log")
        btn_clear.clicked.connect(lambda: self.log_edit.clear())
        top_row.addWidget(btn_clear)
        layout.addLayout(top_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        layout.addWidget(self.progress)

        # Khung Captcha nhúng trực tiếp ngay phía trên Nhật ký
        self.captcha_panel = CaptchaPanelWidget(self)
        self.captcha_panel.setVisible(False)
        layout.addWidget(self.captcha_panel)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setFont(QFont("Consolas", 9))
        self.log_edit.setObjectName("terminalLog")
        layout.addWidget(self.log_edit, 1)

        return card

    def _build_body(self) -> QWidget:
        """Phần chức năng cho tool Tải / Đẩy HSQ ký số (2 tab)."""
        tabs = QTabWidget()

        # --- Tab 1: Tải HSQ ---
        tab_tai = QWidget()
        tai_layout = QVBoxLayout(tab_tai)
        tai_layout.setContentsMargins(8, 10, 8, 8)
        tai_layout.setSpacing(8)

        row_nap = QHBoxLayout()
        lbl_nap = QLabel("Nhập tinhHinhDangKyId hoặc nạp từ Excel có cột madon:")
        lbl_nap.setObjectName("mutedText")
        btn_nap = QPushButton("📥 Nạp Excel")
        btn_nap.setObjectName("btnPurple")
        btn_nap.clicked.connect(self._nap_excel_ids)
        row_nap.addWidget(lbl_nap, 1)
        row_nap.addWidget(btn_nap)
        tai_layout.addLayout(row_nap)

        self.inp_ids = QPlainTextEdit()
        self.inp_ids.setPlaceholderText("Mỗi dòng 1 tinhHinhDangKyId (hoặc cách nhau dấu phẩy / khoảng trắng)")
        tai_layout.addWidget(self.inp_ids, 1)

        row_out = QHBoxLayout()
        self.inp_out_dir = QLineEdit(str(APP_DIR / "hsq_tai_ve"))
        btn_out = QPushButton("📂 Chọn")
        btn_out.clicked.connect(self._chon_out_dir)
        row_out.addWidget(self.inp_out_dir, 1)
        row_out.addWidget(btn_out)
        tai_layout.addLayout(row_out)

        row_btns = QHBoxLayout()
        self.btn_check = QPushButton("🔍 Kiểm tra")
        self.btn_check.setObjectName("btnBlue")
        self.btn_check.clicked.connect(lambda: self._start_tai(chi_kiem_tra=True))
        self.btn_tai = QPushButton("⬇ Tải toàn bộ")
        self.btn_tai.setObjectName("btnPrimary")
        self.btn_tai.clicked.connect(lambda: self._start_tai(chi_kiem_tra=False))
        row_btns.addWidget(self.btn_check)
        row_btns.addWidget(self.btn_tai)
        row_btns.addStretch()
        tai_layout.addLayout(row_btns)

        self.lbl_excel_tai = QLabel("")
        self.lbl_excel_tai.setObjectName("monoGreen")
        self.lbl_excel_tai.setWordWrap(True)
        tai_layout.addWidget(self.lbl_excel_tai)

        # --- Tab 2: Đẩy HSQ ký số ---
        tab_day = QWidget()
        day_layout = QVBoxLayout(tab_day)
        day_layout.setContentsMargins(8, 10, 8, 8)
        day_layout.setSpacing(8)

        row_ex = QHBoxLayout()
        self.inp_excel_day = QLineEdit()
        self.inp_excel_day.setPlaceholderText("Excel kết quả tải (ket_qua_tai_hsq_*.xlsx)")
        btn_ex = QPushButton("📄 Chọn Excel")
        btn_ex.clicked.connect(self._chon_excel_day)
        row_ex.addWidget(self.inp_excel_day, 1)
        row_ex.addWidget(btn_ex)
        day_layout.addLayout(row_ex)

        row_ks = QHBoxLayout()
        self.inp_ky_so_dir = QLineEdit()
        self.inp_ky_so_dir.setPlaceholderText("Thư mục file ký số (để trống = cạnh file đã tải)")
        btn_ks = QPushButton("📂 Chọn")
        btn_ks.clicked.connect(self._chon_ky_so_dir)
        row_ks.addWidget(self.inp_ky_so_dir, 1)
        row_ks.addWidget(btn_ks)
        day_layout.addLayout(row_ks)

        self.chk_chi_da_ky = QCheckBox("Chỉ đẩy file đã có chữ ký số")
        self.chk_chi_da_ky.setChecked(True)
        day_layout.addWidget(self.chk_chi_da_ky)

        self.chk_chay_thu = QCheckBox("Chạy thử (chỉ in kế hoạch, không gửi lên MPLIS)")
        self.chk_chay_thu.setChecked(True)
        self.chk_chay_thu.setStyleSheet("color: #fbbf24;")
        day_layout.addWidget(self.chk_chay_thu)

        note_day = QLabel(
            "• Mỗi HSQ gửi 1 lần kèm ĐỦ mọi thành phần: thành phần không đổi giữ nguyên, "
            "thành phần có file ký số được thay đúng vị trí (thanhPhanHoSoQuetId), giữ mô tả/loại/GCN.\n"
            "• Sau khi đẩy, nếu MPLIS bị giảm số thành phần → dừng toàn bộ ngay.\n"
            "• Trạng thái ghi ngược vào chính file Excel (đóng Excel trước khi chạy)."
        )
        note_day.setObjectName("mutedText")
        note_day.setWordWrap(True)
        day_layout.addWidget(note_day)

        self.btn_day = QPushButton("☁ Đẩy lên MPLIS")
        self.btn_day.setObjectName("btnSuccess")
        self.btn_day.clicked.connect(self._xac_nhan_day)
        day_layout.addWidget(self.btn_day, 0, Qt.AlignmentFlag.AlignLeft)
        day_layout.addStretch()

        tabs.addTab(tab_tai, "⬇ Tải HSQ")
        tabs.addTab(tab_day, "☁ Đẩy HSQ ký số")
        return tabs

    def _apply_stylesheet(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background-color: #07101d;
                color: #e5edf7;
                font-family: 'Segoe UI';
                font-size: 9.5pt;
            }
            QFrame#card {
                background-color: #0d1626;
                border: 1px solid #24324b;
                border-radius: 10px;
            }
            QFrame#captchaCard {
                background-color: #111c30;
                border: 2px solid #f59e0b;
                border-radius: 10px;
            }
            QLabel#cardTitle {
                color: #f8fafc;
                font-size: 10.5pt;
                font-weight: 700;
            }
            QLabel#captchaTitle {
                color: #fbbf24;
                font-size: 10.5pt;
                font-weight: 700;
            }
            QLabel#mutedText {
                color: #8b9bb4;
                font-size: 8.5pt;
            }
            QLabel#statusText {
                color: #e2e8f0;
                font-size: 9pt;
            }
            QLabel#monoCyan {
                color: #38bdf8;
                font-family: Consolas;
                font-size: 9pt;
                font-weight: 600;
            }
            QLabel#monoAmber {
                color: #fbbf24;
                font-family: Consolas;
                font-size: 9pt;
                font-weight: 700;
            }
            QLabel#monoGreen {
                color: #4ade80;
                font-family: Consolas;
                font-size: 8.5pt;
            }
            QLineEdit, QPlainTextEdit {
                background-color: #060b14;
                color: #f1f5f9;
                border: 1px solid #2e3e5c;
                border-radius: 6px;
                padding: 5px 8px;
            }
            QLineEdit:focus, QPlainTextEdit:focus {
                border: 1px solid #38bdf8;
            }
            QPlainTextEdit#terminalLog {
                background-color: #050911;
                color: #cbd5e1;
                border: 1px solid #1e293b;
                border-radius: 8px;
                padding: 8px;
            }
            QPushButton {
                min-height: 32px;
                padding: 0 12px;
                color: #e2e8f0;
                background-color: #1e293b;
                border: 1px solid #334155;
                border-radius: 7px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #28374f;
            }
            QPushButton:disabled {
                color: #64748b;
                background-color: #0f172a;
                border-color: #1e293b;
            }
            QPushButton#btnPrimary {
                color: #042f2e;
                background-color: #2dd4bf;
                border-color: #2dd4bf;
            }
            QPushButton#btnSuccess {
                color: #052e16;
                background-color: #4ade80;
                border-color: #4ade80;
            }
            QPushButton#btnBlue {
                color: #eff6ff;
                background-color: #2563eb;
                border-color: #3b82f6;
            }
            QPushButton#btnPurple {
                color: #f3e8ff;
                background-color: #4c1d95;
                border-color: #7c3aed;
            }
            QPushButton#btnDanger {
                color: #fef2f2;
                background-color: #991b1b;
                border-color: #ef4444;
            }
            QProgressBar {
                background-color: #1e293b;
                border: none;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #38bdf8;
                border-radius: 4px;
            }
            QSlider::groove:horizontal {
                height: 8px;
                border-radius: 4px;
                background: #334155;
            }
            QSlider::handle:horizontal {
                width: 74px;
                margin: -7px 0;
                border-radius: 7px;
                background: #38bdf8;
                border: 1px solid #bae6fd;
            }
            QTabWidget::pane {
                border: 1px solid #24324b;
                border-radius: 8px;
                background: #0d1626;
            }
            QTabBar::tab {
                background: #111c2e;
                color: #94a3b8;
                padding: 7px 14px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background: #1e293b;
                color: #38bdf8;
                font-weight: 700;
            }
            """
        )

    # ------------------------------------------------------------------ Slots UI
    @Slot(str)
    def _append_log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_edit.appendPlainText(f"[{stamp}] {message}")
        cursor = self.log_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_edit.setTextCursor(cursor)

    @Slot(int, int, str)
    def _on_progress(self, done: int, total: int, label: str) -> None:
        val = int(round((done / total) * 1000)) if total else 0
        self.progress.setValue(max(0, min(1000, val)))
        self.lbl_status.setText(label)

    @Slot(str)
    def _on_excel_tai(self, path: str) -> None:
        if hasattr(self, "lbl_excel_tai"):
            self.lbl_excel_tai.setText(f"Excel kết quả: {path}")
        if hasattr(self, "inp_excel_day"):
            self.inp_excel_day.setText(path)

    @Slot(str, str, bool)
    def _on_login_status(self, text: str, color: str, chrome_mo: bool) -> None:
        self.lbl_login_status.setText(text)
        self.lbl_login_status.setStyleSheet(f"color: {color};")
        self.btn_chrome.setEnabled(not chrome_mo)

    @Slot(object, str, object, object, object)
    def _on_show_captcha(self, session, chuc_nang: str, challenge, xong: threading.Event, kq: dict) -> None:
        self.lbl_status.setText("⏸ Đang chờ giải hoặc xác nhận hoàn thành Captcha...")
        self.captcha_panel.open_for_challenge(session, chuc_nang, challenge, xong, kq)
        # Khung Captcha hiện xong thì tự động bấm nút '🤖 Tự động kéo & Giải'
        QTimer.singleShot(300, self.captcha_panel.btn_auto_agent.click)

    @Slot()
    def _on_hide_captcha(self) -> None:
        self.captcha_panel.setVisible(False)

    @Slot(dict)
    def _on_captcha_reloaded(self, new_ch: dict) -> None:
        self.captcha_panel.apply_challenge_data(new_ch)
        self.captcha_panel.lbl_status.setText("Đã tải hình Captcha mới — hãy kéo mảnh ghép vào vết khuyết.")

    @Slot(bool, str, dict)
    def _on_captcha_verify_result(self, verified: bool, msg: str, new_ch: dict) -> None:
        self.captcha_panel.verifying = False
        panel = self.captcha_panel
        if verified:
            panel.finish_captcha(True, msg)
            return
        if new_ch:
            panel.apply_challenge_data(new_ch)
        panel.lbl_status.setText(msg)
        if not panel.isVisible():
            return
        panel.so_lan_giai_hong += 1
        if new_ch and panel.so_lan_giai_hong < SO_LAN_TU_GIAI_CAPTCHA:
            # Verify trả false → tự kéo & giải lại trên hình mới
            self._append_log(
                f"🔁 Giải Captcha chưa khớp (lần {panel.so_lan_giai_hong}/{SO_LAN_TU_GIAI_CAPTCHA}) — tự giải lại..."
            )
            QTimer.singleShot(800, panel.btn_auto_agent.click)
        else:
            # Hết số lần tự giải / lỗi Verify → mở Chrome để giải tay, xong bấm '✓ Đã giải xong Captcha (trên Chrome)'
            self._append_log("🌐 Tự giải Captcha không được — mở Chrome để bạn giải tay rồi bấm '✓ Đã giải xong Captcha'.")
            QTimer.singleShot(300, panel.btn_open_chrome.click)

    @Slot(str)
    def _on_done(self, status_text: str) -> None:
        self.is_running = False
        self._set_running(False)
        self.lbl_status.setText(status_text)

    # ------------------------------------------------------------------ Captcha Worker Hook
    def _mo_thu_captcha(self) -> None:
        session = None
        if self.rad_login.isChecked() and self.login_client and self.login_client.session:
            session = self.login_client.session
        elif self.rad_cookie.isChecked() and self.inp_token.text().strip() and self.inp_cookie.toPlainText().strip():
            try:
                tmp = core.MplisClient(thd.log)
                tmp.build_session_from_manual(self.inp_token.text().strip(), self.inp_cookie.toPlainText().strip())
                session = tmp.session
            except Exception:
                session = None

        chuc_nang = "HoSoQuetAjax/UpdateHoSoQuetExistFile"

        def _preview_worker():
            challenge = None
            if session is not None:
                try:
                    challenge = core.lay_challenge_slider_captcha(session, chuc_nang)
                    thd.log(
                        f"📥 Đã lấy Challenge thật từ MPLIS "
                        f"(challengeId={str(challenge.get('challengeId'))[:10]}..., pieceTop={challenge.get('pieceTop')}px)"
                    )
                except Exception as exc:
                    thd.log(f"⚠ Không lấy được Challenge thật ({core.rut_gon_text(exc, 120)}) — mở chế độ mô phỏng PySide6.")
            xong = threading.Event()
            kq = {"tiep": False}
            self.bridge.captchaSignal.emit(session, chuc_nang, challenge or {}, xong, kq)

        threading.Thread(target=_preview_worker, daemon=True).start()

    def _cho_giai_captcha(self, session, chuc_nang: str) -> bool:
        if self.stop_flag:
            return False
        xong = threading.Event()
        kq = {"tiep": False}
        challenge = None
        try:
            challenge = core.lay_challenge_slider_captcha(session, chuc_nang)
            thd.log(
                f"⚠ MPLIS yêu cầu SliderCaptcha ({chuc_nang}) — đã hiện khung giải & xác nhận ngay trên giao diện PySide6 "
                f"(challengeId={str(challenge.get('challengeId'))[:10]}..., pieceTop={challenge.get('pieceTop')}px)."
            )
        except Exception as exc:
            thd.log(
                f"⚠ MPLIS yêu cầu Captcha ({chuc_nang}) nhưng chưa tải được ảnh tự động ({core.rut_gon_text(exc, 140)}) "
                "— đã bật khung xác nhận hoàn thành Captcha trên giao diện."
            )

        self.bridge.captchaSignal.emit(session, chuc_nang, challenge or {}, xong, kq)
        while not xong.wait(0.5):
            if self.stop_flag:
                self.bridge.captchaHideSignal.emit()
                return False
        if not kq["tiep"]:
            thd.log("⏹ Đã huỷ / dừng tại bước giải Captcha.")
            return False

        driver = self.login_client.driver if self.rad_login.isChecked() and self.login_client else None
        if driver is not None and not getattr(session, "last_recaptcha_token", ""):
            try:
                for c in driver.get_cookies():
                    session.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))
            except Exception as exc:
                thd.log(f"   ⚠ Không chép lại được cookie từ Chrome: {core.rut_gon_text(exc, 150)}")
        thd.log("▶ Đã hoàn thành Captcha — gửi lại thao tác đang dở.")
        return True

    # ------------------------------------------------------------------ Phiên đăng nhập & chọn file
    def _doi_che_do(self) -> None:
        dang_nhap = self.rad_login.isChecked()
        self.panel_login.setVisible(dang_nhap)
        self.panel_cookie.setVisible(not dang_nhap)

    def _mo_hoac_dua_chrome_len(self) -> None:
        """Tự động bấm mở Chrome hoặc đưa cửa sổ Chrome đang mở lên phía trước để người dùng giải Captcha."""
        if self.login_client and self.login_client.driver:
            def _focus_driver():
                try:
                    drv = self.login_client.driver
                    drv.switch_to.window(drv.current_window_handle)
                    # window.focus() không đưa được cửa sổ lên trước trên Windows → thu nhỏ rồi phóng to lại
                    drv.minimize_window()
                    drv.maximize_window()
                    drv.execute_script("window.focus();")
                    thd.log("🌐 Đã đưa cửa sổ Chrome đang mở lên trước để bạn giải Captcha.")
                except Exception as exc:
                    thd.log(f"⚠ Không focus được Chrome đang mở ({core.rut_gon_text(exc, 120)}), đang mở tab MPLIS...")
                    try:
                        os.startfile(core.REFERER_LOGIN)
                    except Exception:
                        pass

            threading.Thread(target=_focus_driver, daemon=True).start()
            return

        if self.inp_username.text().strip() and self.inp_password.text():
            thd.log("🌐 Tự động bấm '1. Mở Chrome' để bạn đăng nhập / giải Captcha...")
            self._mo_chrome()
            return

        try:
            subprocess.Popen(["cmd", "/c", "start", "chrome", core.REFERER_LOGIN])
            thd.log(f"🌐 Đã mở Chrome tới {core.REFERER_LOGIN} để bạn giải Captcha.")
        except Exception:
            try:
                os.startfile(core.REFERER_LOGIN)
            except Exception as exc:
                thd.log(f"⚠ Không mở được trình duyệt tự động: {core.rut_gon_text(exc, 120)}")

    def _mo_chrome(self) -> None:
        user = self.inp_username.text().strip()
        pwd = self.inp_password.text()
        if not user or not pwd:
            QMessageBox.warning(self, "Thiếu thông tin", "Hãy nhập Username và Password.")
            return
        if self.login_client and self.login_client.driver:
            QMessageBox.information(self, "Chrome đang mở", "Chrome đang mở — đăng nhập xong thì bấm '2. Lấy session'.")
            return
        self.btn_chrome.setEnabled(False)
        self.lbl_login_status.setText("Đang mở Chrome...")
        self.lbl_login_status.setStyleSheet("color: #fbbf24;")

        def work():
            client = core.MplisClient(thd.log)
            try:
                client.open_browser_and_fill_login(user, pwd)
                self.login_client = client
                self.bridge.loginSignal.emit(
                    "Chrome đã mở. Đăng nhập xong (OTP nếu có) thì bấm '2. Lấy session'.",
                    "#fbbf24",
                    True,
                )
            except Exception as exc:
                client.close_browser()
                thd.log(f"❌ Mở Chrome lỗi: {core.rut_gon_text(exc, 300)}")
                self.bridge.loginSignal.emit("Mở Chrome lỗi", "#ef4444", False)

        threading.Thread(target=work, daemon=True).start()

    def _lay_session_chrome(self) -> None:
        client = self.login_client
        if client is None or client.driver is None:
            QMessageBox.warning(self, "Chưa mở Chrome", "Bấm '1. Mở Chrome' và đăng nhập trước.")
            return

        def work():
            try:
                client.build_session_from_browser()
                self.bridge.loginSignal.emit(
                    "✅ Đã có session. Chrome vẫn mở — hết hạn thì bấm lại '2. Lấy session'.",
                    "#4ade80",
                    True,
                )
            except Exception as exc:
                thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
                self.bridge.loginSignal.emit(
                    "Chưa lấy được session — kiểm tra đã đăng nhập xong chưa.",
                    "#ef4444",
                    True,
                )

        threading.Thread(target=work, daemon=True).start()

    def _nguon_session(self) -> dict[str, Any] | None:
        if self.rad_login.isChecked():
            if self.login_client is None or self.login_client.session is None:
                QMessageBox.warning(self, "Chưa có session", "Đăng nhập Chrome rồi bấm '2. Lấy session' trước.")
                return None
            return {"client": self.login_client}
        token = self.inp_token.text().strip()
        cookie = self.inp_cookie.toPlainText().strip()
        if not token or not cookie:
            QMessageBox.warning(self, "Thiếu token/cookie", "Hãy dán đủ __RequestVerificationToken và Cookie trước.")
            return None
        return {"token": token, "cookie": cookie}

    def _tao_client(self, nguon: dict[str, Any]) -> core.MplisClient:
        if "client" in nguon:
            return nguon["client"]
        client = core.MplisClient(thd.log)
        client.build_session_from_manual(nguon["token"], nguon["cookie"])
        return client

    def _nap_excel_ids(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel có cột madon", str(APP_DIR), "Excel (*.xlsx *.xlsm)"
        )
        if not path:
            return
        try:
            ids = thd.tach_danh_sach_id(thd.doc_id_tu_excel(path))
        except Exception as exc:
            QMessageBox.critical(self, "Lỗi đọc Excel", f"Không đọc được Excel: {exc}")
            return
        cu = thd.tach_danh_sach_id(self.inp_ids.toPlainText())
        gop = cu + [i for i in ids if i not in cu]
        self.inp_ids.setPlainText("\n".join(str(i) for i in gop))
        self._append_log(f"Nạp {len(ids)} ID từ {os.path.basename(path)} (tổng {len(gop)}).")

    def _chon_out_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục", self.inp_out_dir.text().strip() or str(APP_DIR))
        if folder:
            self.inp_out_dir.setText(os.path.normpath(folder))

    def _chon_ky_so_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục file ký số", self.inp_ky_so_dir.text().strip() or self.inp_out_dir.text().strip()
        )
        if folder:
            self.inp_ky_so_dir.setText(os.path.normpath(folder))

    def _chon_excel_day(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn Excel kết quả tải", self.inp_out_dir.text().strip() or str(APP_DIR), "Excel (*.xlsx *.xlsm)"
        )
        if path:
            self.inp_excel_day.setText(os.path.normpath(path))

    def _mo_excel(self) -> None:
        path = self.inp_excel_day.text().strip() if hasattr(self, "inp_excel_day") else ""
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Chưa có Excel", "Chưa có file Excel kết quả.")
            return
        os.startfile(path)

    def _mo_thu_muc(self) -> None:
        path = self.inp_out_dir.text().strip()
        if not os.path.isdir(path):
            QMessageBox.warning(self, "Chưa có thư mục", "Thư mục lưu chưa tồn tại.")
            return
        subprocess.Popen(["explorer", os.path.normpath(path)])

    # ------------------------------------------------------------------ Chạy / Dừng
    def _set_running(self, running: bool) -> None:
        for btn in (self.btn_check, self.btn_tai, self.btn_day):
            btn.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _stop(self) -> None:
        if self.is_running:
            self.stop_flag = True
            if self.captcha_panel.isVisible():
                self.captcha_panel.finish_captcha(False)
            self._append_log("⏹ Đã yêu cầu dừng, sẽ dừng sau file đang xử lý...")

    def _run_thread(self, target: Callable, *args) -> None:
        self.is_running = True
        self.stop_flag = False
        self._set_running(True)
        self.progress.setValue(0)
        threading.Thread(target=target, args=args, daemon=True).start()

    def _start_tai(self, chi_kiem_tra: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        ids = thd.tach_danh_sach_id(self.inp_ids.toPlainText())
        if not ids:
            QMessageBox.warning(self, "Thiếu ID", "Chưa có tinhHinhDangKyId hợp lệ.")
            return
        out_dir = self.inp_out_dir.text().strip() or str(APP_DIR / "hsq_tai_ve")
        self._run_thread(self._worker_tai, nguon, ids, out_dir, chi_kiem_tra)

    def _xac_nhan_day(self) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        path = self.inp_excel_day.text().strip()
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Thiếu file Excel", "Chọn Excel kết quả tải trước.")
            return
        chay_thu = self.chk_chay_thu.isChecked()
        if not chay_thu:
            ans = QMessageBox.question(
                self,
                "Xác nhận đẩy HSQ",
                f"Đẩy file ký số thay thế file hiện có trên MPLIS theo:\n{path}\n\n"
                + ("Chỉ đẩy file đã có chữ ký số." if self.chk_chi_da_ky.isChecked() else "⚠ Đẩy cả file CHƯA ký số.")
                + "\nThao tác GHI THẬT lên MPLIS. Tiếp tục?",
            )
            if ans != QMessageBox.StandardButton.Yes:
                return
        self._run_thread(
            self._worker_day,
            nguon,
            path,
            self.inp_ky_so_dir.text().strip(),
            self.chk_chi_da_ky.isChecked(),
            chay_thu,
        )

    def _worker_tai(self, nguon: dict[str, Any], ids: list[int], out_dir: str, chi_kiem_tra: bool) -> None:
        tong: list[dict[str, Any]] = []
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            token = client.session.headers.get("__requestverificationtoken", "")
            downloader = thd.Downloader(None, client.session, token, thd.doc_mau_da_luu(), thd.MAU_MAC_DINH)
            os.makedirs(out_dir, exist_ok=True)
            excel_kq = "" if chi_kiem_tra else thd.ten_excel_ket_qua(out_dir)

            thd.log(f"{'Kiểm tra' if chi_kiem_tra else 'Tải'} {len(ids)} đơn...")
            for n, id_don in enumerate(ids, 1):
                if self.stop_flag:
                    ket_thuc = "Đã dừng"
                    break
                self.bridge.progressSignal.emit(n - 1, len(ids), f"Đơn {id_don} ({n}/{len(ids)})")
                thd.log(f"--- [{n}/{len(ids)}] Đơn {id_don} ---")
                het_phien = False
                try:
                    kq = thd.xu_ly_don(client, downloader, id_don, out_dir, chi_kiem_tra, nen_dung=lambda: self.stop_flag)
                except core.HetPhien as exc:
                    thd.log(f"❌ {exc}")
                    thd.log("⛔ HẾT PHIÊN — đã dừng. Lấy session mới rồi chạy tiếp các đơn chưa tải.")
                    kq = thd.kq_het_phien(id_don, exc)
                    het_phien = True
                    ket_thuc = "HẾT PHIÊN — đã dừng"
                except Exception as exc:
                    thd.log(f"❌ Đơn {id_don}: {core.rut_gon_text(exc, 300)}")
                    kq = {
                        "id_don": id_don,
                        "so_file": 0,
                        "da_tai": 0,
                        "loi": [str(exc)],
                        "dong": [{"madon": id_don, "trang_thai_tai": "Lỗi", "ghi_chu_tai": str(exc)}],
                    }
                tong.append(kq)
                if excel_kq:
                    try:
                        thd.ghi_excel_ket_qua(excel_kq, [d for k in tong for d in k["dong"]])
                    except PermissionError:
                        thd.log("⚠ Không ghi được Excel kết quả (file đang mở?)")
                if het_phien:
                    break
            self.bridge.progressSignal.emit(len(tong), len(ids), ket_thuc)

            so_file = sum(k["so_file"] for k in tong)
            thd.log("=" * 50)
            if chi_kiem_tra:
                thd.log(f"✅ {len(tong)} đơn | {so_file} file HSQ")
            else:
                thd.log(f"✅ {len(tong)} đơn | {so_file} file HSQ | đã tải {sum(k['da_tai'] for k in tong)} → {out_dir}")
            if excel_kq and tong:
                thd.log(f"Excel kết quả: {excel_kq}")
                self.bridge.excelTaiSignal.emit(excel_kq)
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self.bridge.doneSignal.emit(ket_thuc)

    def _worker_day(self, nguon: dict[str, Any], excel_path: str, ky_so_dir: str, chi_da_ky: bool, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            day.day_tu_excel(
                client,
                excel_path,
                ky_so_dir,
                chi_da_ky,
                chay_thu=chay_thu,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self.bridge.progressSignal.emit(done, total, label),
            )
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except core.HetPhien as exc:
            thd.log(f"❌ {exc}")
            thd.log("⛔ HẾT PHIÊN — đã dừng. Lấy session mới rồi bấm Đẩy lại (dòng đã đẩy sẽ được bỏ qua).")
            ket_thuc = "HẾT PHIÊN — đã dừng"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self.bridge.doneSignal.emit(ket_thuc)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self.login_client:
            threading.Thread(target=self.login_client.close_browser, daemon=True).start()
        super().closeEvent(event)


def main() -> int:
    app = QApplication(sys.argv)
    win = TaiHsqPySide6Window()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
