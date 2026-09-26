# -*- coding: utf-8 -*-
"""Giao diện Flet cho công cụ cập nhật hồ sơ quét MPLIS.

File này chỉ đảm nhiệm giao diện và điều phối. Toàn bộ nghiệp vụ đọc Excel,
tìm đơn/GCN/HSQ, gọi API và ghi kết quả được tái sử dụng trực tiếp từ
``cap_nhat_don_hsq.py``.
"""

from __future__ import annotations

import asyncio
import os
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import flet as ft

import cap_nhat_don_hsq as core


APP_DIR = Path(__file__).resolve().parent


class CyberTheme:
    BG_DARK = "#090D16"
    SURFACE_DARK = "#111827"
    SURFACE_CARD = "#172033"
    SURFACE_HOVER = "#1E293B"
    SURFACE_INPUT = "#0D1322"
    BORDER_COLOR = "#23304B"
    CYAN = "#00F0FF"
    BLUE = "#3B82F6"
    GREEN = "#10B981"
    RED = "#F43F5E"
    AMBER = "#F59E0B"
    PURPLE = "#8B5CF6"
    TEXT_MUTED = "#64748B"
    TEXT_LIGHT = "#94A3B8"
    TEXT_WHITE = "#F8FAFC"
    TERMINAL_BG = "#060A12"


def _native_choose_excel(initial_dir: Optional[str] = None) -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    try:
        root.withdraw()
        root.attributes("-topmost", True)
        return filedialog.askopenfilename(
            initialdir=initial_dir if initial_dir and os.path.isdir(initial_dir) else str(APP_DIR),
            title="Chọn file Excel cần cập nhật hồ sơ quét",
            filetypes=[
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro-Enabled", "*.xlsm"),
                ("Tất cả tập tin", "*.*"),
            ],
        ) or ""
    finally:
        root.destroy()


def _native_save_excel(initial_file: str = "ket_qua_cap_nhat_don_hsq.xlsx") -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    try:
        root.withdraw()
        root.attributes("-topmost", True)
        return filedialog.asksaveasfilename(
            initialdir=str(APP_DIR),
            initialfile=Path(initial_file).name,
            title="Chọn nơi lưu file kết quả",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        ) or ""
    finally:
        root.destroy()


def _native_choose_directory(initial_dir: Optional[str] = None) -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    try:
        root.withdraw()
        root.attributes("-topmost", True)
        return filedialog.askdirectory(
            initialdir=initial_dir if initial_dir and os.path.isdir(initial_dir) else str(APP_DIR),
            title="Chọn folder gốc chứa hồ sơ quét (bao gồm các folder con)",
        ) or ""
    finally:
        root.destroy()


def _gcn_lookup_key(value: str) -> str:
    """Chuẩn hóa GCN thành khóa chữ/số để lập chỉ mục tên PDF."""
    stem = Path(str(value or "").strip()).stem
    return "".join(re.findall(r"[A-Z0-9]+", core.remove_accents(stem).upper()))


def _split_excel_values(value: Any) -> list[str]:
    """Tách danh sách trong một ô Excel theo dấu chấm phẩy hoặc xuống dòng."""
    return [part.strip() for part in re.split(r"[;\r\n]+", str(value or "")) if part.strip()]


def expand_rows_by_sogcn(
    rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    """Mở rộng một dòng có nhiều ``sogcn`` thành một lượt xử lý cho mỗi GCN.

    Mã đơn, tờ/thửa và các thông tin dùng chung được giữ nguyên. Các trường nhận
    diện GCN phụ chỉ được ghép theo vị trí khi chúng cũng có đúng số phần tử;
    cách này tránh vô tình dùng một ``gcn_id`` cho nhiều giấy chứng nhận.
    """
    expanded_rows: list[dict[str, Any]] = []
    split_source_rows = 0

    for row in rows:
        raw_sogcn = str(row.get("sogcn") or "").strip()
        gcn_parts = _split_excel_values(raw_sogcn)

        # Loại GCN trùng trong cùng một ô nhưng vẫn giữ nguyên thứ tự nhập.
        unique_gcns: list[str] = []
        seen_keys: set[str] = set()
        for gcn in gcn_parts:
            key = _gcn_lookup_key(gcn) or gcn.upper()
            if key not in seen_keys:
                seen_keys.add(key)
                unique_gcns.append(gcn)

        if len(unique_gcns) <= 1:
            expanded_rows.append(row)
            continue

        split_source_rows += 1
        item_count = len(unique_gcns)
        raw_tenfile = str(row.get("tenfile") or "").strip()
        tenfile_was_auto_filled = raw_tenfile.casefold() == f"{raw_sogcn}.pdf".casefold()
        tenfile_parts = [] if tenfile_was_auto_filled else _split_excel_values(raw_tenfile)
        sovaoso_parts = _split_excel_values(row.get("sovaoso"))
        gcn_id_parts = _split_excel_values(row.get("gcn_id"))

        for index, so_gcn in enumerate(unique_gcns):
            expanded = dict(row)
            expanded["sogcn"] = so_gcn
            expanded["tenfile"] = (
                tenfile_parts[index]
                if len(tenfile_parts) == item_count
                else f"{so_gcn}.pdf"
            )
            expanded["sovaoso"] = (
                sovaoso_parts[index] if len(sovaoso_parts) == item_count else ""
            )
            expanded["gcn_id"] = (
                gcn_id_parts[index] if len(gcn_id_parts) == item_count else ""
            )
            expanded["_split_gcn_index"] = index + 1
            expanded["_split_gcn_total"] = item_count
            expanded_rows.append(expanded)

    return expanded_rows, split_source_rows


def resolve_pdfs_by_gcn(
    folder_upload: str, so_gcns: list[str]
) -> dict[str, tuple[str | None, str]]:
    """Quét folder đúng một lần và tìm PDF cho toàn bộ danh sách GCN."""
    display_by_key: dict[str, str] = {}
    for so_gcn in so_gcns:
        key = _gcn_lookup_key(so_gcn)
        if key:
            display_by_key.setdefault(key, str(so_gcn).strip())

    if not folder_upload or not os.path.isdir(folder_upload):
        return {key: (None, "Thư mục PDF không tồn tại.") for key in display_by_key}
    if not display_by_key:
        return {}

    scan_errors: list[OSError] = []

    def remember_scan_error(exc: OSError) -> None:
        scan_errors.append(exc)

    try:
        pdf_names = sorted(
            os.path.relpath(os.path.join(current_folder, name), folder_upload)
            for current_folder, _subfolders, files in os.walk(
                folder_upload, onerror=remember_scan_error
            )
            for name in files
            if name.lower().endswith(".pdf")
        )
    except OSError as exc:
        return {
            key: (None, f"Không đọc được thư mục PDF: {exc}")
            for key in display_by_key
        }

    requested_keys = set(display_by_key)
    max_key_length = max(len(key) for key in requested_keys)
    exact_matches: dict[str, set[str]] = {key: set() for key in requested_keys}
    embedded_matches: dict[str, set[str]] = {key: set() for key in requested_keys}

    for name in pdf_names:
        stem = Path(name).stem
        whole_key = _gcn_lookup_key(stem)
        if whole_key in requested_keys:
            exact_matches[whole_key].add(name)

        tokens = re.findall(r"[A-Z0-9]+", core.remove_accents(stem).upper())
        for start in range(len(tokens)):
            candidate = ""
            for token in tokens[start:]:
                candidate += token
                if len(candidate) > max_key_length:
                    break
                if candidate in requested_keys:
                    embedded_matches[candidate].add(name)

    results: dict[str, tuple[str | None, str]] = {}
    for key, display in display_by_key.items():
        exact = sorted(exact_matches[key])
        if len(exact) == 1:
            results[key] = (exact[0], "Khớp chính xác số GCN")
            continue
        if len(exact) > 1:
            results[key] = (
                exact[0],
                f"Có {len(exact)} file trùng chính xác; chọn file đầu tiên: {exact[0]}",
            )
            continue

        embedded = sorted(embedded_matches[key])
        if len(embedded) == 1:
            results[key] = (embedded[0], "Khớp số GCN nằm trong tên file")
        elif not embedded:
            scan_note = f"; có {len(scan_errors)} folder con không đọc được" if scan_errors else ""
            results[key] = (
                None,
                f"Không tìm thấy PDF chứa số GCN '{display}' trong folder gốc và các folder con{scan_note}.",
            )
        else:
            results[key] = (
                embedded[0],
                f"Có {len(embedded)} PDF cùng khớp GCN '{display}'; "
                f"chọn file đầu tiên: {embedded[0]}",
            )
    return results


def resolve_pdf_by_gcn(folder_upload: str, so_gcn: str) -> tuple[str | None, str]:
    """Tìm duy nhất một PDF theo số GCN, không chọn đại khi kết quả mơ hồ.

    Ví dụ ``DN 512781`` khớp ``24337-GCN-DN 512781 - 2509263_.pdf``.
    Quét cả folder gốc và toàn bộ folder con, trả về đường dẫn tương đối.
    """
    key = _gcn_lookup_key(so_gcn)
    if not key:
        return None, "Số GCN trống hoặc không có ký tự nhận diện."
    return resolve_pdfs_by_gcn(folder_upload, [so_gcn]).get(
        key, (None, f"Không tìm thấy PDF chứa số GCN '{so_gcn}'.")
    )


class CoreRowProcessor:
    """Adapter tối thiểu để tái sử dụng nguyên hàm xử lý dòng của giao diện Tk."""

    def __init__(self, client: core.MplisClient, log_fn):
        self.client = client
        self.log = log_fn
        self.stop_flag = False

    def process_row(self, **kwargs) -> list[dict[str, Any]]:
        return core.App._xu_ly_1_row(self, **kwargs)


class CapNhatDonHsqFletApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.loop = asyncio.get_running_loop()
        self.message_queue: asyncio.Queue = asyncio.Queue()

        self.client = core.MplisClient(self._thread_log)
        self.processor = CoreRowProcessor(self.client, self._thread_log)
        self.worker_thread: threading.Thread | None = None
        self.is_running = False
        self.session_ready = False
        self.browser_opened = False
        self.pending_run: dict[str, Any] | None = None
        self.logs_history: list[tuple[str, str, str]] = []

        self.current_nav_index = 0

    async def init_ui(self) -> None:
        self.page.title = "MPLIS • Cập nhật hồ sơ quét"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.bgcolor = CyberTheme.BG_DARK
        self.page.padding = 0
        self.page.spacing = 0
        self.page.window.width = 1260
        self.page.window.height = 800
        self.page.window.min_width = 1000
        self.page.window.min_height = 650
        await self.page.window.center()

        self._build_header()
        self._build_process_view()
        self._build_login_view()
        self._build_guide_view()
        self._build_footer()

        self.main_container = ft.Container(
            expand=True,
            content=self.process_view,
            padding=ft.Padding.only(left=16, right=16, top=10, bottom=10),
        )

        self.page.add(
            ft.Column(
                expand=True,
                spacing=0,
                controls=[self.header_bar, self.main_container, self.footer_bar],
            )
        )
        asyncio.create_task(self._consume_messages())
        self._append_log("Giao diện Flet đã sẵn sàng. Hãy tạo session MPLIS trước khi chạy.", "title")
        self._append_log(
            "Tên file dài được nhập ở cột tenfile; số phát hành thật, ví dụ DN 512781, nhập ở cột sogcn.",
            "info",
        )
        self.page.update()

    # ------------------------------------------------------------------ UI basics
    def _card(
        self,
        content: ft.Control,
        title: str = "",
        icon: str | None = None,
        expand: bool = False,
        header_actions: list[ft.Control] | None = None,
    ) -> ft.Container:
        controls: list[ft.Control] = []
        if title:
            header_controls = [
                ft.Icon(icon or ft.Icons.DASHBOARD_ROUNDED, color=CyberTheme.CYAN, size=17),
                ft.Text(title, color=CyberTheme.TEXT_WHITE, size=12, weight=ft.FontWeight.BOLD),
            ]
            if header_actions:
                header_row = ft.Row(
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Row(spacing=8, controls=header_controls),
                        ft.Row(spacing=6, controls=header_actions),
                    ],
                )
            else:
                header_row = ft.Row(spacing=8, controls=header_controls)
            controls.append(header_row)
            controls.append(ft.Divider(height=1, color=CyberTheme.BORDER_COLOR))
        controls.append(content)
        return ft.Container(
            expand=expand,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            border_radius=10,
            padding=12,
            content=ft.Column(expand=expand, spacing=8, controls=controls),
        )

    def _input(
        self,
        label: str,
        value: str = "",
        hint: str = "",
        password: bool = False,
        multiline: bool = False,
        expand: bool = False,
        width: int | None = None,
    ) -> ft.TextField:
        return ft.TextField(
            label=label,
            value=value,
            hint_text=hint,
            password=password,
            can_reveal_password=password,
            multiline=multiline,
            min_lines=2 if multiline else 1,
            max_lines=3 if multiline else 1,
            expand=expand,
            width=width,
            text_size=12,
            label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=11),
            color=CyberTheme.TEXT_WHITE,
            bgcolor=CyberTheme.SURFACE_INPUT,
            border={
                ft.ControlState.DEFAULT: ft.OutlineInputBorder(
                    side=ft.BorderSide(1, CyberTheme.BORDER_COLOR),
                    border_radius=8,
                ),
                ft.ControlState.FOCUSED: ft.OutlineInputBorder(
                    side=ft.BorderSide(1, CyberTheme.CYAN),
                    border_radius=8,
                ),
            },
            dense=True,
        )

    def _button(
        self,
        text: str,
        icon: str,
        color: str,
        callback,
        filled: bool = False,
        width: int | None = None,
        height: int = 38,
    ) -> ft.Container:
        foreground = CyberTheme.BG_DARK if filled else color
        return ft.Container(
            content=ft.Row(
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=7,
                controls=[
                    ft.Icon(icon, color=foreground, size=16),
                    ft.Text(text, color=foreground, size=12, weight=ft.FontWeight.W_600),
                ],
            ),
            width=width,
            height=height,
            bgcolor=color if filled else f"{color}18",
            border=ft.Border.all(1, color),
            border_radius=8,
            padding=ft.Padding.symmetric(horizontal=12, vertical=0),
            alignment=ft.Alignment(0, 0),
            ink=True,
            on_click=callback,
        )

    def _show_toast(self, message: str, error: bool = False, success: bool = False) -> None:
        color = CyberTheme.RED if error else (CyberTheme.GREEN if success else CyberTheme.BLUE)
        icon = ft.Icons.ERROR_OUTLINE_ROUNDED if error else (
            ft.Icons.CHECK_CIRCLE_OUTLINE_ROUNDED if success else ft.Icons.INFO_OUTLINE_ROUNDED
        )
        self.page.show_dialog(
            ft.SnackBar(
                content=ft.Row(
                    spacing=9,
                    controls=[ft.Icon(icon, color=color, size=19), ft.Text(message, color=CyberTheme.TEXT_WHITE, size=12)],
                ),
                bgcolor=CyberTheme.SURFACE_CARD,
                behavior=ft.SnackBarBehavior.FLOATING,
                margin=ft.Margin.all(16),
                duration=4000,
            )
        )

    # ------------------------------------------------------------------ header/nav
    def _build_header(self) -> None:
        self.status_icon = ft.Icon(ft.Icons.RADIO_BUTTON_CHECKED_ROUNDED, color=CyberTheme.AMBER, size=14)
        self.status_text = ft.Text("CHƯA CÓ SESSION", color=CyberTheme.AMBER, size=10, weight=ft.FontWeight.BOLD)
        self.status_badge = ft.Container(
            content=ft.Row(spacing=6, controls=[self.status_icon, self.status_text]),
            padding=ft.Padding.symmetric(horizontal=10, vertical=5),
            bgcolor=f"{CyberTheme.AMBER}15",
            border=ft.Border.all(1, f"{CyberTheme.AMBER}55"),
            border_radius=20,
        )

        self.nav_buttons = [
            self._nav_button("Xử lý Excel", ft.Icons.DASHBOARD_ROUNDED, 0, True),
            self._nav_button("Đăng nhập MPLIS", ft.Icons.SECURITY_ROUNDED, 1, False),
            self._nav_button("Hướng dẫn", ft.Icons.HELP_OUTLINE_ROUNDED, 2, False),
        ]

        self.btn_header_start = self._button(
            "Chạy cập nhật",
            ft.Icons.PLAY_ARROW_ROUNDED,
            CyberTheme.GREEN,
            self._prepare_run,
            filled=True,
            height=36,
        )

        self.header_bar = ft.Container(
            height=62,
            bgcolor=CyberTheme.SURFACE_DARK,
            border=ft.Border.only(bottom=ft.BorderSide(1, CyberTheme.BORDER_COLOR)),
            padding=ft.Padding.symmetric(horizontal=16),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                controls=[
                    ft.Row(
                        spacing=12,
                        controls=[
                            ft.Container(
                                width=38,
                                height=38,
                                border_radius=10,
                                gradient=ft.LinearGradient(
                                    colors=[CyberTheme.CYAN, CyberTheme.BLUE],
                                    begin=ft.Alignment(-1, -1),
                                    end=ft.Alignment(1, 1),
                                ),
                                alignment=ft.Alignment(0, 0),
                                content=ft.Icon(ft.Icons.DOCUMENT_SCANNER_ROUNDED, color=CyberTheme.BG_DARK, size=21),
                            ),
                            ft.Column(
                                spacing=0,
                                alignment=ft.MainAxisAlignment.CENTER,
                                controls=[
                                    ft.Text("CẬP NHẬT HỒ SƠ QUÉT", color=CyberTheme.TEXT_WHITE, size=14, weight=ft.FontWeight.BOLD),
                                    ft.Text("MPLIS REGISTRATION SYNC", color=CyberTheme.CYAN, size=9, weight=ft.FontWeight.W_700),
                                ],
                            ),
                            ft.VerticalDivider(width=16, color=CyberTheme.BORDER_COLOR),
                            self.status_badge,
                        ],
                    ),
                    ft.Row(spacing=6, controls=self.nav_buttons),
                    ft.Row(
                        spacing=8,
                        controls=[
                            self.btn_header_start,
                            self._button("Mở thư mục kết quả", ft.Icons.FOLDER_OPEN_ROUNDED, CyberTheme.CYAN, self._open_output_folder, height=36),
                        ],
                    ),
                ],
            ),
        )

    def _nav_button(self, title: str, icon: str, index: int, active: bool) -> ft.Container:
        color = CyberTheme.CYAN if active else CyberTheme.TEXT_LIGHT
        return ft.Container(
            content=ft.Row(spacing=6, controls=[ft.Icon(icon, color=color, size=16), ft.Text(title, color=color, size=11)]),
            bgcolor=f"{CyberTheme.CYAN}18" if active else "transparent",
            border=ft.Border.all(1, CyberTheme.CYAN if active else "transparent"),
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            border_radius=8,
            ink=True,
            on_click=lambda e, idx=index: self._switch_view(idx),
        )

    def _switch_view(self, index: int) -> None:
        self.current_nav_index = index
        for idx, button in enumerate(self.nav_buttons):
            active = idx == index
            color = CyberTheme.CYAN if active else CyberTheme.TEXT_LIGHT
            button.bgcolor = f"{CyberTheme.CYAN}18" if active else "transparent"
            button.border = ft.Border.all(1, CyberTheme.CYAN if active else "transparent")
            button.content.controls[0].color = color
            button.content.controls[1].color = color
        self.main_container.content = [self.process_view, self.login_view, self.guide_view][index]
        self.page.update()

    # ------------------------------------------------------------------ process view
    def _metric(self, title: str, value: ft.Text, icon: str, color: str) -> ft.Container:
        return ft.Container(
            expand=True,
            height=66,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            border_radius=10,
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.START,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=10,
                controls=[
                    ft.Container(
                        width=38,
                        height=38,
                        border_radius=9,
                        bgcolor=f"{color}18",
                        alignment=ft.Alignment(0, 0),
                        content=ft.Icon(icon, color=color, size=20),
                    ),
                    ft.Column(
                        spacing=1,
                        alignment=ft.MainAxisAlignment.CENTER,
                        controls=[
                            ft.Text(title, color=CyberTheme.TEXT_MUTED, size=9, weight=ft.FontWeight.BOLD),
                            value,
                        ],
                    ),
                ],
            ),
        )

    def _build_process_view(self) -> None:
        # Thẻ thống kê (Metrics Bar)
        self.metric_total = ft.Text("0", color=CyberTheme.TEXT_WHITE, size=18, weight=ft.FontWeight.BOLD)
        self.metric_progress = ft.Text("0 / 0", color=CyberTheme.CYAN, size=18, weight=ft.FontWeight.BOLD)
        self.metric_success = ft.Text("0", color=CyberTheme.GREEN, size=18, weight=ft.FontWeight.BOLD)
        self.metric_error = ft.Text("0", color=CyberTheme.RED, size=18, weight=ft.FontWeight.BOLD)

        metrics = ft.Row(
            spacing=10,
            controls=[
                self._metric("TỔNG NHÓM", self.metric_total, ft.Icons.DATA_ARRAY_ROUNDED, CyberTheme.BLUE),
                self._metric("TIẾN ĐỘ XỬ LÝ", self.metric_progress, ft.Icons.TIMELAPSE_ROUNDED, CyberTheme.CYAN),
                self._metric("THÀNH CÔNG", self.metric_success, ft.Icons.CHECK_CIRCLE_ROUNDED, CyberTheme.GREEN),
                self._metric("LỖI / BỎ QUA", self.metric_error, ft.Icons.ERROR_ROUNDED, CyberTheme.RED),
            ],
        )

        # Thanh điều khiển & NÚT CHẠY CHÍNH (BẮT ĐẦU CẬP NHẬT) - Nổi bật ngay trên đầu
        self.btn_start = ft.Container(
            content=ft.Row(
                spacing=8,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.PLAY_ARROW_ROUNDED, color=CyberTheme.BG_DARK, size=22),
                    ft.Text("BẮT ĐẦU CẬP NHẬT", color=CyberTheme.BG_DARK, size=13, weight=ft.FontWeight.BOLD),
                ],
            ),
            width=200,
            height=42,
            border_radius=8,
            gradient=ft.LinearGradient(
                colors=[CyberTheme.GREEN, "#059669"],
                begin=ft.Alignment(-1, 0),
                end=ft.Alignment(1, 0),
            ),
            border=ft.Border.all(1, CyberTheme.GREEN),
            ink=True,
            on_click=self._prepare_run,
        )

        self.btn_stop = ft.Container(
            content=ft.Row(
                spacing=7,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.STOP_ROUNDED, color=CyberTheme.TEXT_MUTED, size=20),
                    ft.Text("DỪNG", color=CyberTheme.TEXT_MUTED, size=13, weight=ft.FontWeight.BOLD),
                ],
            ),
            width=100,
            height=42,
            border_radius=8,
            bgcolor=CyberTheme.SURFACE_HOVER,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            ink=False,
            on_click=self._stop_run,
        )

        self.session_quick_info = ft.Container(
            content=ft.Row(
                spacing=8,
                controls=[
                    ft.Icon(ft.Icons.WARNING_AMBER_ROUNDED, color=CyberTheme.AMBER, size=16),
                    ft.Text("Chưa có Session MPLIS — Bấm để đăng nhập ngay", color=CyberTheme.AMBER, size=11, weight=ft.FontWeight.W_600),
                ],
            ),
            bgcolor=f"{CyberTheme.AMBER}15",
            border=ft.Border.all(1, f"{CyberTheme.AMBER}50"),
            border_radius=8,
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            ink=True,
            on_click=lambda e: self._switch_view(1),
        )

        primary_action_bar = ft.Container(
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            border_radius=10,
            padding=ft.Padding.symmetric(horizontal=14, vertical=8),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    self.session_quick_info,
                    ft.Row(
                        spacing=10,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            self.btn_start,
                            self.btn_stop,
                            ft.IconButton(
                                icon=ft.Icons.FILE_OPEN_ROUNDED,
                                icon_color=CyberTheme.PURPLE,
                                tooltip="Mở file Excel kết quả",
                                on_click=self._open_output_file,
                            ),
                            ft.IconButton(
                                icon=ft.Icons.FOLDER_OPEN_ROUNDED,
                                icon_color=CyberTheme.CYAN,
                                tooltip="Mở thư mục lưu kết quả",
                                on_click=self._open_output_folder,
                            ),
                        ],
                    ),
                ],
            ),
        )

        # 1. Tệp tin và Thư mục làm việc
        self.input_file = self._input("File Excel đầu vào (.xlsx, .xlsm)", expand=True, hint="Chọn file Excel chứa madon hoặc soto/sothua, sogcn...")
        self.folder_upload = self._input("Thư mục chứa file PDF hồ sơ quét", expand=True, hint="Chọn thư mục chứa file scan...")
        self.output_file = self._input("File kết quả sau cập nhật", expand=True, hint="ket_qua_cap_nhat_don_hsq.xlsx")

        file_section = self._card(
            ft.Column(
                spacing=8,
                controls=[
                    ft.Row(
                        controls=[
                            self.input_file,
                            self._button("CHỌN EXCEL", ft.Icons.UPLOAD_FILE_ROUNDED, CyberTheme.BLUE, self._choose_input, width=125, height=38),
                        ]
                    ),
                    ft.Row(
                        controls=[
                            self.folder_upload,
                            self._button("CHỌN FOLDER", ft.Icons.FOLDER_ROUNDED, CyberTheme.CYAN, self._choose_folder, width=125, height=38),
                        ]
                    ),
                    ft.Row(
                        controls=[
                            self.output_file,
                            self._button("CHỌN NƠI LƯU", ft.Icons.SAVE_ALT_ROUNDED, CyberTheme.PURPLE, self._choose_output, width=125, height=38),
                        ]
                    ),
                ],
            ),
            "1. Tệp tin & Thư mục làm việc",
            ft.Icons.FOLDER_COPY_ROUNDED,
        )

        # 2. Tham số và Tùy chọn nghiệp vụ
        self.xa_id = self._input("Mã xã (xaId)", width=130, hint="Tra theo tờ/thửa")
        self.ngay_dk = self._input("Ngày ĐK lần đầu", width=140, hint="dd/mm/yyyy")
        self.chu_id = self._input("ID thông tin chủ", expand=True, hint="Để trống nếu không đổi")

        self.loai_hsq = ft.Dropdown(
            value="1",
            label="Loại hồ sơ quét",
            expand=True,
            height=42,
            dense=True,
            text_size=11,
            color=CyberTheme.TEXT_WHITE,
            bgcolor=CyberTheme.SURFACE_INPUT,
            label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=10),
            options=[
                ft.DropdownOption(key=str(key), text=f"{key} · {label}")
                for key, label in core.LOAI_HO_SO_QUET_OPTIONS.items()
            ],
            border={
                ft.ControlState.DEFAULT: ft.OutlineInputBorder(
                    side=ft.BorderSide(1, CyberTheme.BORDER_COLOR),
                    border_radius=8,
                ),
                ft.ControlState.FOCUSED: ft.OutlineInputBorder(
                    side=ft.BorderSide(1, CyberTheme.CYAN),
                    border_radius=8,
                ),
            },
        )

        self.day_hsq = ft.Checkbox(
            label="Upload PDF",
            value=True,
            active_color=CyberTheme.CYAN,
            on_change=self._on_upload_changed,
            tooltip="Đẩy file PDF thật lên hồ sơ quét",
        )
        self.thay_moi_gcn = ft.Checkbox(
            label="Thay đúng GCN",
            value=True,
            active_color=CyberTheme.CYAN,
            tooltip="Khớp HSQ theo số GCN thay vì chỉ CHUACOGIAY",
        )
        self.dry_run = ft.Checkbox(
            label="DRY RUN",
            value=False,
            active_color=CyberTheme.AMBER,
            tooltip="Chỉ kiểm tra, không gửi update",
        )

        options_section = self._card(
            ft.Column(
                spacing=8,
                controls=[
                    ft.Row(spacing=8, controls=[self.xa_id, self.ngay_dk, self.chu_id]),
                    ft.Row(
                        spacing=8,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            self.loai_hsq,
                            ft.Container(
                                bgcolor=CyberTheme.SURFACE_INPUT,
                                border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
                                border_radius=8,
                                padding=ft.Padding.symmetric(horizontal=8, vertical=1),
                                content=ft.Row(
                                    spacing=6,
                                    controls=[self.day_hsq, self.thay_moi_gcn, self.dry_run],
                                ),
                            ),
                        ],
                    ),
                ],
            ),
            "2. Tham số & Tùy chọn nghiệp vụ",
            ft.Icons.TUNE_ROUNDED,
        )

        rules_card = ft.Container(
            bgcolor=f"{CyberTheme.BLUE}10",
            border=ft.Border.all(1, f"{CyberTheme.BLUE}40"),
            border_radius=10,
            padding=10,
            content=ft.Column(
                spacing=3,
                controls=[
                    ft.Row(
                        spacing=6,
                        controls=[
                            ft.Icon(ft.Icons.LIGHTBULB_OUTLINE_ROUNDED, color=CyberTheme.CYAN, size=15),
                            ft.Text("Quy tắc nhận diện tên file & số GCN", color=CyberTheme.CYAN, size=11, weight=ft.FontWeight.BOLD),
                        ],
                    ),
                    ft.Text(
                        "• Ví dụ: tenfile = '24337-GCN-DN 512781 - 2509263_.pdf'  →  nhập sogcn = 'DN 512781'.",
                        color=CyberTheme.TEXT_LIGHT,
                        size=10,
                        selectable=True,
                    ),
                    ft.Text(
                        "• Một đơn có nhiều GCN: nhập 'BQ 809137;BQ 809138;BQ 809132' trong cùng ô sogcn; công cụ tự tách và xử lý từng GCN.",
                        color=CyberTheme.TEXT_LIGHT,
                        size=10,
                        selectable=True,
                    ),
                    ft.Text(
                        "• Hệ thống tự tìm file PDF chứa số GCN này trong thư mục và cập nhật mô tả: Giấy chứng nhận DN 512781.",
                        color=CyberTheme.GREEN,
                        size=10,
                    ),
                ],
            ),
        )

        left_config_panel = ft.Column(
            expand=True,
            spacing=10,
            scroll=ft.ScrollMode.AUTO,
            controls=[
                file_section,
                options_section,
                rules_card,
            ],
        )

        # 3. Live Terminal & Nhật Ký
        self.log_counter_text = ft.Text("0 dòng", color=CyberTheme.TEXT_MUTED, size=10)
        self.log_list = ft.ListView(expand=True, spacing=4, auto_scroll=True)
        terminal_card = self._card(
            self.log_list,
            "Nhật ký xử lý hệ thống",
            ft.Icons.TERMINAL_ROUNDED,
            expand=True,
            header_actions=[
                self.log_counter_text,
                ft.IconButton(
                    icon=ft.Icons.DELETE_SWEEP_ROUNDED,
                    icon_color=CyberTheme.TEXT_MUTED,
                    icon_size=17,
                    tooltip="Xóa nhật ký hiển thị",
                    on_click=self._clear_logs,
                ),
            ],
        )

        # 4. Khu vực Workspace cân đối (Cột trái cấu hình, Cột phải Terminal)
        workspace = ft.Row(
            expand=True,
            spacing=12,
            controls=[
                ft.Container(expand=5, content=left_config_panel),
                ft.Container(expand=6, content=terminal_card),
            ],
        )

        self.process_view = ft.Column(
            expand=True,
            spacing=10,
            controls=[
                metrics,
                primary_action_bar,
                workspace,
            ],
        )

    # ------------------------------------------------------------------ login view
    def _build_login_view(self) -> None:
        self.username = self._input("Username MPLIS", expand=True)
        self.password = self._input("Password MPLIS", expand=True, password=True)
        self.token = self._input(
            "Dán __RequestVerificationToken tại đây",
            expand=True,
            hint="F12 → Network → Request Headers → __requestverificationtoken",
        )
        self.cookie = self._input(
            "Dán Cookie tại đây",
            expand=True,
            multiline=True,
            hint="F12 → Network → Request Headers → Cookie (dán nguyên chuỗi)",
        )
        self.login_result = ft.Text("Chưa tạo session.", color=CyberTheme.AMBER, size=11)

        browser_card = self._card(
            ft.Column(
                spacing=12,
                controls=[
                    ft.Text(
                        "Chrome sẽ mở để đăng nhập. Hoàn tất OTP nếu có, sau đó quay lại bấm Lấy session.",
                        color=CyberTheme.TEXT_LIGHT,
                        size=11,
                    ),
                    ft.Row(controls=[self.username, self.password]),
                    ft.Row(
                        spacing=10,
                        controls=[
                            self._button("1. Mở Chrome", ft.Icons.OPEN_IN_BROWSER_ROUNDED, CyberTheme.BLUE, self._open_chrome),
                            self._button("2. Lấy session và đóng Chrome", ft.Icons.VPN_KEY_ROUNDED, CyberTheme.GREEN, self._capture_session),
                        ],
                    ),
                ],
            ),
            "Đăng nhập qua Chrome",
            ft.Icons.WEB_ROUNDED,
        )

        manual_card = self._card(
            ft.Column(
                spacing=12,
                controls=[
                    ft.Container(
                        bgcolor=f"{CyberTheme.PURPLE}18",
                        border=ft.Border.all(1, f"{CyberTheme.PURPLE}55"),
                        border_radius=8,
                        padding=10,
                        content=ft.Row(
                            spacing=8,
                            controls=[
                                ft.Icon(ft.Icons.CONTENT_PASTE_ROUNDED, color=CyberTheme.PURPLE, size=19),
                                ft.Text(
                                    "DÁN TOKEN VÀ COOKIE VÀO HAI Ô BÊN DƯỚI",
                                    color=CyberTheme.TEXT_WHITE,
                                    size=11,
                                    weight=ft.FontWeight.BOLD,
                                ),
                            ],
                        ),
                    ),
                    ft.Text(
                        "Có thể bỏ qua Chrome bằng cách sao chép Token và Cookie từ F12 → Network.",
                        color=CyberTheme.TEXT_LIGHT,
                        size=11,
                    ),
                    self.token,
                    self.cookie,
                    self._button("Dùng Token / Cookie", ft.Icons.SECURITY_ROUNDED, CyberTheme.PURPLE, self._manual_session),
                ],
            ),
            "Token / Cookie thủ công",
            ft.Icons.CODE_ROUNDED,
        )

        self.login_view = ft.Column(
            expand=True,
            spacing=12,
            scroll=ft.ScrollMode.AUTO,
            controls=[
                manual_card,
                browser_card,
                ft.Container(
                    bgcolor=CyberTheme.SURFACE_CARD,
                    border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
                    border_radius=10,
                    padding=14,
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Row(
                                spacing=8,
                                controls=[ft.Icon(ft.Icons.INFO_OUTLINE_ROUNDED, color=CyberTheme.CYAN), self.login_result],
                            ),
                            self._button(
                                "Về trang Xử lý Excel",
                                ft.Icons.ARROW_BACK_ROUNDED,
                                CyberTheme.CYAN,
                                lambda e: self._switch_view(0),
                                filled=True,
                                height=36,
                            ),
                        ],
                    ),
                ),
            ],
        )

    # ------------------------------------------------------------------ guide view
    def _guide_step(self, number: str, title: str, body: str) -> ft.Container:
        return ft.Container(
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            border_radius=10,
            padding=14,
            content=ft.Row(
                vertical_alignment=ft.CrossAxisAlignment.START,
                controls=[
                    ft.Container(
                        width=34,
                        height=34,
                        border_radius=17,
                        bgcolor=f"{CyberTheme.CYAN}20",
                        alignment=ft.Alignment(0, 0),
                        content=ft.Text(number, color=CyberTheme.CYAN, weight=ft.FontWeight.BOLD),
                    ),
                    ft.Column(
                        expand=True,
                        spacing=3,
                        controls=[
                            ft.Text(title, color=CyberTheme.TEXT_WHITE, size=13, weight=ft.FontWeight.BOLD),
                            ft.Text(body, color=CyberTheme.TEXT_LIGHT, size=11),
                        ],
                    ),
                ],
            ),
        )

    def _build_guide_view(self) -> None:
        self.guide_view = ft.Column(
            expand=True,
            spacing=10,
            scroll=ft.ScrollMode.AUTO,
            controls=[
                ft.Text("QUY TRÌNH CẬP NHẬT HỒ SƠ QUÉT", color=CyberTheme.TEXT_WHITE, size=18, weight=ft.FontWeight.BOLD),
                self._guide_step("1", "Tạo session MPLIS", "Đăng nhập qua Chrome hoặc dán Token/Cookie thủ công."),
                self._guide_step(
                    "2",
                    "Chuẩn bị Excel",
                    "Dùng cột madon/tinhhinhdangkyid hoặc cặp soto+sothua. Với file 24337-GCN-DN 512781 - 2509263_.pdf, nhập sogcn là DN 512781. Nhiều GCN cùng đơn có thể đặt chung một ô và ngăn cách bằng dấu ';'.",
                ),
                self._guide_step("3", "Chọn PDF và tùy chọn", "Bật upload để thay file thật; tắt upload nếu chỉ sửa metadata của file HSQ đã có."),
                self._guide_step(
                    "4",
                    "Kiểm tra bằng DRY RUN",
                    "Nên chạy thử trước. DRY RUN vẫn đọc chi tiết đơn nhưng không gửi request cập nhật.",
                ),
                self._guide_step(
                    "5",
                    "Chạy cập nhật",
                    "Công cụ tìm đúng GCN, chọn HSQ thuộc thông tin đăng ký, cập nhật file/mô tả rồi ghi kết quả Excel.",
                ),
                ft.Container(
                    bgcolor=f"{CyberTheme.AMBER}12",
                    border=ft.Border.all(1, f"{CyberTheme.AMBER}55"),
                    border_radius=10,
                    padding=14,
                    content=ft.Text(
                        "Lưu ý: tra theo tờ/thửa sẽ xử lý tất cả đơn tìm thấy. Cập nhật HSQ và cập nhật ngày/chủ là hai request riêng biệt.",
                        color=CyberTheme.AMBER,
                        size=11,
                    ),
                ),
            ],
        )

    # ------------------------------------------------------------------ footer/logging
    def _build_footer(self) -> None:
        self.footer_status = ft.Text("Sẵn sàng", color=CyberTheme.TEXT_LIGHT, size=10)
        self.footer_percent = ft.Text("0%", color=CyberTheme.CYAN, size=10, weight=ft.FontWeight.BOLD)
        self.footer_brand = ft.Text(
            "Phòng dữ liệu - thông tin đất đai\nTổ ứng dụng và phát triển công nghệ",
            color=CyberTheme.TEXT_LIGHT,
            size=9,
            text_align=ft.TextAlign.CENTER,
        )
        self.progress_bar = ft.ProgressBar(value=0, color=CyberTheme.CYAN, bgcolor=CyberTheme.SURFACE_HOVER, height=4)
        self.footer_bar = ft.Container(
            bgcolor=CyberTheme.SURFACE_DARK,
            border=ft.Border.only(top=ft.BorderSide(1, CyberTheme.BORDER_COLOR)),
            padding=ft.Padding.only(left=20, right=20, top=7, bottom=8),
            content=ft.Column(
                spacing=5,
                controls=[
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[self.footer_status, self.footer_brand, self.footer_percent],
                    ),
                    self.progress_bar,
                ],
            ),
        )

    def _append_log(self, message: str, level: str = "info") -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        color = {
            "success": CyberTheme.GREEN,
            "error": CyberTheme.RED,
            "warn": CyberTheme.AMBER,
            "title": CyberTheme.CYAN,
        }.get(level, CyberTheme.TEXT_LIGHT)
        icon = {
            "success": ft.Icons.CHECK_CIRCLE_OUTLINE_ROUNDED,
            "error": ft.Icons.ERROR_OUTLINE_ROUNDED,
            "warn": ft.Icons.WARNING_AMBER_ROUNDED,
            "title": ft.Icons.CHEVRON_RIGHT_ROUNDED,
        }.get(level, ft.Icons.SUBDIRECTORY_ARROW_RIGHT_ROUNDED)
        self.logs_history.append((timestamp, message, level))
        self.log_list.controls.append(
            ft.Container(
                bgcolor=CyberTheme.TERMINAL_BG,
                border_radius=6,
                padding=ft.Padding.symmetric(horizontal=8, vertical=6),
                content=ft.Row(
                    vertical_alignment=ft.CrossAxisAlignment.START,
                    spacing=7,
                    controls=[
                        ft.Text(timestamp, color=CyberTheme.TEXT_MUTED, size=9),
                        ft.Icon(icon, color=color, size=13),
                        ft.Text(message, color=color, size=10, selectable=True, expand=True),
                    ],
                ),
            )
        )
        if len(self.log_list.controls) > 1000:
            del self.log_list.controls[:100]
        if hasattr(self, "log_counter_text") and self.log_counter_text.page:
            self.log_counter_text.value = f"{len(self.log_list.controls)} dòng"
            self.log_counter_text.update()
        if self.log_list.page:
            self.log_list.update()

    def _clear_logs(self, _=None) -> None:
        """Xóa phần nhật ký đang hiển thị trên giao diện."""
        self.logs_history.clear()
        self.log_list.controls.clear()
        if hasattr(self, "log_counter_text") and self.log_counter_text.page:
            self.log_counter_text.value = "0 dòng"
            self.log_counter_text.update()
        if self.log_list.page:
            self.log_list.update()
        self._show_toast("Đã xóa nhật ký hiển thị.", success=True)

    def _thread_log(self, message: str) -> None:
        level = "error" if "❌" in message or "lỗi" in message.lower() else (
            "success" if "✅" in message or "thành công" in message.lower() else "info"
        )
        self._emit("log", message, level)

    def _emit(self, kind: str, *args) -> None:
        self.loop.call_soon_threadsafe(self.message_queue.put_nowait, (kind, *args))

    # ------------------------------------------------------------------ file actions
    async def _choose_input(self, _=None) -> None:
        selected = await asyncio.to_thread(_native_choose_excel, str(APP_DIR))
        if not selected:
            return
        self.input_file.value = selected
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_file.value = str(Path(selected).with_name(f"ket_qua_cap_nhat_don_hsq_{stamp}.xlsx"))
        self.page.update()

    async def _choose_output(self, _=None) -> None:
        initial = self.output_file.value.strip() or "ket_qua_cap_nhat_don_hsq.xlsx"
        selected = await asyncio.to_thread(_native_save_excel, initial)
        if selected:
            self.output_file.value = selected
            self.output_file.update()

    async def _choose_folder(self, _=None) -> None:
        selected = await asyncio.to_thread(_native_choose_directory, self.folder_upload.value.strip())
        if selected:
            self.folder_upload.value = selected
            self.folder_upload.update()

    def _open_output_file(self, _=None) -> None:
        """Mở file Excel kết quả bằng ứng dụng mặc định của hệ điều hành."""
        raw = self.output_file.value.strip() if hasattr(self, "output_file") else ""
        if not raw or not os.path.isfile(raw):
            self._show_toast("File kết quả chưa tồn tại.", error=True)
            return
        try:
            if sys.platform == "win32":
                os.startfile(raw)
            else:
                subprocess.Popen(["xdg-open", raw])
        except Exception as exc:
            self._show_toast(f"Không mở được file kết quả: {exc}", error=True)

    def _open_output_folder(self, _=None) -> None:
        raw = self.output_file.value.strip() if hasattr(self, "output_file") else ""
        folder = str(Path(raw).resolve().parent) if raw else str(APP_DIR)
        if not os.path.isdir(folder):
            self._show_toast("Thư mục kết quả chưa tồn tại.", error=True)
            return
        try:
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            self._show_toast(f"Không mở được thư mục: {exc}", error=True)

    def _on_upload_changed(self, _=None) -> None:
        self.folder_upload.disabled = not bool(self.day_hsq.value)
        if self.folder_upload.page:
            self.folder_upload.update()

    # ------------------------------------------------------------------ session actions
    async def _open_chrome(self, _=None) -> None:
        if self.is_running:
            return
        username = self.username.value.strip()
        password = self.password.value or ""
        if not username or not password:
            self._show_toast("Nhập username và password trước.", error=True)
            return
        self.login_result.value = "Đang mở Chrome..."
        self.login_result.color = CyberTheme.CYAN
        self.login_result.update()
        try:
            await asyncio.to_thread(self.client.open_browser_and_fill_login, username, password)
        except Exception as exc:
            self.login_result.value = f"Lỗi mở Chrome: {exc}"
            self.login_result.color = CyberTheme.RED
            self.login_result.update()
            self._append_log(f"Lỗi mở Chrome: {exc}", "error")
            return
        self.browser_opened = True
        self.login_result.value = "Chrome đã mở. Hoàn tất đăng nhập/OTP rồi bấm Lấy session."
        self.login_result.color = CyberTheme.AMBER
        self.login_result.update()

    async def _capture_session(self, _=None) -> None:
        if not self.browser_opened and not self.client.driver:
            self._show_toast("Chưa mở Chrome từ ứng dụng.", error=True)
            return
        self.login_result.value = "Đang lấy token và cookie..."
        self.login_result.color = CyberTheme.CYAN
        self.login_result.update()
        try:
            await asyncio.to_thread(self.client.build_session_from_browser)
            await asyncio.to_thread(self.client.close_browser)
        except Exception as exc:
            self.login_result.value = str(exc)
            self.login_result.color = CyberTheme.RED
            self.login_result.update()
            return
        self.browser_opened = False
        self._set_session_ready("Đã lấy session từ Chrome và đóng trình duyệt.")

    async def _manual_session(self, _=None) -> None:
        token = self.token.value.strip()
        cookie = self.cookie.value.strip()
        if not token or not cookie:
            self._show_toast("Nhập đủ Token và Cookie.", error=True)
            return
        try:
            await asyncio.to_thread(self.client.build_session_from_manual, token, cookie)
        except Exception as exc:
            self.login_result.value = str(exc)
            self.login_result.color = CyberTheme.RED
            self.login_result.update()
            return
        self._set_session_ready("Đã tạo session từ Token/Cookie thủ công.")

    def _set_session_ready(self, message: str) -> None:
        self.session_ready = True
        self.status_text.value = "SESSION SẴN SÀNG"
        self.status_text.color = CyberTheme.GREEN
        self.status_icon.color = CyberTheme.GREEN
        self.status_badge.bgcolor = f"{CyberTheme.GREEN}15"
        self.status_badge.border = ft.Border.all(1, f"{CyberTheme.GREEN}55")
        self.login_result.value = message
        self.login_result.color = CyberTheme.GREEN
        if hasattr(self, "session_quick_info") and self.session_quick_info.page:
            self.session_quick_info.bgcolor = f"{CyberTheme.GREEN}15"
            self.session_quick_info.border = ft.Border.all(1, f"{CyberTheme.GREEN}50")
            row = self.session_quick_info.content
            row.controls[0].name = ft.Icons.CHECK_CIRCLE_ROUNDED
            row.controls[0].color = CyberTheme.GREEN
            row.controls[1].value = "Session MPLIS sẵn sàng • Có thể bắt đầu cập nhật ngay"
            row.controls[1].color = CyberTheme.GREEN
            self.session_quick_info.update()
        self.page.update()
        self._append_log(message, "success")
        self._show_toast(message, success=True)

    # ------------------------------------------------------------------ run validation/start
    def _prepare_run(self, _=None) -> None:
        """Callback đồng bộ để nút luôn phản hồi ngay trên Flet Desktop."""
        if self.is_running:
            self._show_toast("Tiến trình đang chạy.")
            return
        self.footer_status.value = "Đã nhận lệnh chạy — đang kiểm tra dữ liệu..."
        self.footer_status.color = CyberTheme.CYAN
        self.page.update()
        self._append_log("Đã bấm CHẠY CẬP NHẬT; bắt đầu kiểm tra session và dữ liệu đầu vào.", "title")
        self.page.run_task(self._prepare_run_safe)

    async def _prepare_run_safe(self) -> None:
        try:
            await self._prepare_run_impl()
        except Exception as exc:
            message = f"Lỗi chuẩn bị chạy: {exc}"
            self.footer_status.value = message
            self.footer_status.color = CyberTheme.RED
            self.page.update()
            self._append_log(message, "error")
            self._show_toast(message, error=True)

    async def _prepare_run_impl(self) -> None:
        if self.is_running:
            self._show_toast("Tiến trình đang chạy.")
            return
        if self.client.session is None:
            self._show_toast("Chưa có session MPLIS. Hãy đăng nhập trước.", error=True)
            self._switch_view(1)
            return

        input_file = self.input_file.value.strip()
        output_file = self.output_file.value.strip()
        folder_upload = self.folder_upload.value.strip()
        xa_id = self.xa_id.value.strip()
        ngay_dk = self.ngay_dk.value.strip()
        chu_id = self.chu_id.value.strip() or None
        day_hsq = bool(self.day_hsq.value)
        thay_moi_gcn = bool(self.thay_moi_gcn.value)
        dry_run = bool(self.dry_run.value)

        if not input_file or not os.path.isfile(input_file):
            self._show_toast("Chưa chọn file Excel hợp lệ.", error=True)
            return
        if not output_file:
            self._show_toast("Chưa chọn file kết quả.", error=True)
            return
        if Path(input_file).resolve() == Path(output_file).resolve():
            self._show_toast("File kết quả không được trùng file đầu vào.", error=True)
            return
        if day_hsq and not os.path.isdir(folder_upload):
            self._show_toast("Thư mục PDF không tồn tại.", error=True)
            return
        if ngay_dk:
            try:
                core.ddmmyyyy_to_iso_utc_start_of_day_vn(ngay_dk)
            except Exception:
                self._show_toast("Ngày đăng ký phải có dạng dd/mm/yyyy.", error=True)
                return

        try:
            rows = await asyncio.to_thread(core.doc_danh_sach, input_file)
        except Exception as exc:
            self._show_toast(f"Lỗi đọc Excel: {exc}", error=True)
            return

        source_row_count = len(rows)
        rows, split_source_rows = expand_rows_by_sogcn(rows)
        if split_source_rows:
            self._append_log(
                f"Đã tách {split_source_rows} dòng có nhiều GCN: "
                f"{source_row_count} dòng Excel → {len(rows)} lượt xử lý GCN.",
                "title",
            )

        if day_hsq:
            unique_gcns: dict[str, str] = {}
            for item in rows:
                so_gcn = str(item.get("sogcn") or "").strip()
                if so_gcn:
                    unique_gcns.setdefault(_gcn_lookup_key(so_gcn), so_gcn)

            self.footer_status.value = (
                f"Đang quét folder gốc và các folder con cho {len(unique_gcns)} số GCN..."
            )
            self.page.update()
            resolution_cache = await asyncio.to_thread(
                resolve_pdfs_by_gcn, folder_upload, list(unique_gcns.values())
            )
            resolution_errors: list[str] = []
            resolved_count = 0
            matched_unique: dict[str, tuple[str, str, str]] = {}
            for item in rows:
                so_gcn = str(item.get("sogcn") or "").strip()
                if not so_gcn:
                    continue
                cache_key = _gcn_lookup_key(so_gcn)
                matched_name, reason = resolution_cache.get(
                    cache_key, (None, "Số GCN không có ký tự nhận diện hợp lệ.")
                )
                if matched_name:
                    item["tenfile"] = matched_name
                    resolved_count += 1
                    matched_unique.setdefault(cache_key, (so_gcn, matched_name, reason))
                else:
                    item["_pdf_skip_reason"] = reason
                    resolution_errors.append(
                        f"Dòng {item.get('excel_row')}: GCN {so_gcn} — {reason}"
                    )

            if resolution_errors:
                for message in resolution_errors[:20]:
                    self._append_log(f"Bỏ qua — {message}", "warn")
                if len(resolution_errors) > 20:
                    self._append_log(
                        f"Còn {len(resolution_errors) - 20} dòng không có PDF sẽ được bỏ qua.",
                        "warn",
                    )
                self._append_log(
                    f"Có {len(resolution_errors)} dòng không xác định được PDF; "
                    "các dòng này sẽ ghi Bỏ qua, những dòng còn lại vẫn tiếp tục.",
                    "title",
                )
            if resolved_count:
                for so_gcn, matched_name, reason in list(matched_unique.values())[:10]:
                    self._append_log(f"GCN {so_gcn} → {matched_name} ({reason}).", "success")
                if len(matched_unique) > 10:
                    self._append_log(
                        f"... và {len(matched_unique) - 10} số GCN khác đã khớp file.",
                        "info",
                    )
                self._append_log(
                    f"Đã quét toàn bộ cây folder 1 lần và xác định {resolved_count} dòng PDF "
                    f"cho {len(matched_unique)} số GCN.",
                    "title",
                )

        is_mode_id = any(bool(item.get("id_don")) for item in rows)
        if not is_mode_id and (not xa_id or not xa_id.isdigit()):
            self._show_toast("Cần nhập mã xã dạng số khi tra theo tờ/thửa.", error=True)
            return
        if xa_id and not xa_id.isdigit():
            self._show_toast("Mã xã phải là số nguyên.", error=True)
            return

        grouped: dict[str, list[dict[str, Any]]] = {}
        for item in rows:
            grouped.setdefault(core.make_group_key(item), []).append(item)
        groups = list(grouped.items())
        loai_hsq = int(self.loai_hsq.value or "1")

        self.pending_run = {
            "xa_id": xa_id,
            "groups": groups,
            "rows": rows,
            "output_file": output_file,
            "day_hsq": day_hsq,
            "chu_id": chu_id,
            "loai_hsq": loai_hsq,
            "folder_upload": folder_upload,
            "ngay_dk": ngay_dk,
            "thay_moi_gcn": thay_moi_gcn,
            "dry_run": dry_run,
            "source_row_count": source_row_count,
            "split_source_rows": split_source_rows,
            "preflight_skipped": len(resolution_errors) if day_hsq else 0,
        }

        mode = "ID đơn" if is_mode_id else "Tờ / Thửa"
        row_summary = f"{source_row_count} dòng Excel"
        if split_source_rows:
            row_summary += f" → {len(rows)} lượt GCN"
        confirm_text = (
            f"Chế độ: {mode}\n"
            f"{row_summary} → {len(groups)} nhóm xử lý\n"
            f"Upload PDF: {'Có' if day_hsq else 'Không, chỉ sửa metadata'}\n"
            f"Thay theo GCN: {'Có' if thay_moi_gcn else 'Chỉ CHUACOGIAY'}\n"
            f"Loại HSQ: {loai_hsq} - {core.LOAI_HO_SO_QUET_OPTIONS.get(loai_hsq, '')}\n"
            f"Không có PDF sẽ bỏ qua: {self.pending_run['preflight_skipped']} dòng\n"
            f"DRY RUN: {'Có' if dry_run else 'Không — cập nhật thật'}"
        )

        def confirm(_event) -> None:
            self.page.pop_dialog()
            self._launch_worker()

        self.page.show_dialog(
            ft.AlertDialog(
                title=ft.Row(
                    spacing=8,
                    controls=[
                        ft.Icon(ft.Icons.WARNING_AMBER_ROUNDED, color=CyberTheme.AMBER),
                        ft.Text("Xác nhận cập nhật MPLIS", color=CyberTheme.TEXT_WHITE, weight=ft.FontWeight.BOLD),
                    ],
                ),
                content=ft.Text(confirm_text, color=CyberTheme.TEXT_LIGHT, size=12),
                actions=[
                    ft.TextButton("Hủy", on_click=lambda e: self.page.pop_dialog()),
                    ft.Button(
                        "Tiếp tục",
                        bgcolor=CyberTheme.GREEN,
                        color=CyberTheme.BG_DARK,
                        on_click=confirm,
                    ),
                ],
                bgcolor=CyberTheme.SURFACE_CARD,
            )
        )

    def _launch_worker(self) -> None:
        if not self.pending_run or self.is_running:
            return
        cfg = self.pending_run
        self.is_running = True
        self.processor.stop_flag = False
        core.DRY_RUN = bool(cfg["dry_run"])

        total = len(cfg["groups"])
        self.metric_total.value = str(total)
        self.metric_progress.value = f"0 / {total}"
        self.metric_success.value = "0"
        self.metric_error.value = "0"
        self.progress_bar.value = 0
        self.footer_percent.value = "0%"
        self.footer_status.value = "Đang khởi tạo tiến trình..."
        self.status_text.value = "ĐANG XỬ LÝ"
        self.status_text.color = CyberTheme.CYAN
        self.status_icon.color = CyberTheme.CYAN

        # Cập nhật trạng thái nút Chạy sang "Đang xử lý"
        self.btn_start.gradient = None
        self.btn_start.bgcolor = CyberTheme.SURFACE_HOVER
        self.btn_start.border = ft.Border.all(1, CyberTheme.BORDER_COLOR)
        self.btn_start.content.controls[0].name = ft.Icons.HOURGLASS_TOP_ROUNDED
        self.btn_start.content.controls[0].color = CyberTheme.TEXT_MUTED
        self.btn_start.content.controls[1].value = "ĐANG XỬ LÝ..."
        self.btn_start.content.controls[1].color = CyberTheme.TEXT_MUTED
        self.btn_start.ink = False

        if hasattr(self, "btn_header_start") and self.btn_header_start.page:
            self.btn_header_start.bgcolor = CyberTheme.SURFACE_HOVER
            self.btn_header_start.border = ft.Border.all(1, CyberTheme.BORDER_COLOR)
            self.btn_header_start.content.controls[0].name = ft.Icons.HOURGLASS_TOP_ROUNDED
            self.btn_header_start.content.controls[0].color = CyberTheme.TEXT_MUTED
            self.btn_header_start.content.controls[1].value = "Đang chạy..."
            self.btn_header_start.content.controls[1].color = CyberTheme.TEXT_MUTED
            self.btn_header_start.ink = False

        self.btn_stop.bgcolor = f"{CyberTheme.RED}25"
        self.btn_stop.border = ft.Border.all(1, CyberTheme.RED)
        self.btn_stop.content.controls[0].color = CyberTheme.RED
        self.btn_stop.content.controls[1].color = CyberTheme.RED
        self.btn_stop.ink = True

        self.page.update()

        self.worker_thread = threading.Thread(target=self._run_worker, args=(cfg,), daemon=True)
        self.worker_thread.start()

    def _stop_run(self, _=None) -> None:
        if not self.is_running:
            self._show_toast("Không có tiến trình đang chạy.")
            return
        self.processor.stop_flag = True
        self._append_log("Đã yêu cầu dừng; chương trình sẽ dừng sau request hiện tại và lưu kết quả.", "warn")
        self.footer_status.value = "Đang chờ request hiện tại kết thúc..."
        self.footer_status.update()

    # ------------------------------------------------------------------ core worker
    def _run_worker(self, cfg: dict[str, Any]) -> None:
        writer = core.ExcelResultWriter(cfg["output_file"])
        groups = cfg["groups"]
        processed = 0
        total_results = 0
        success = 0
        errors = 0
        skipped = 0
        fatal_error = ""

        try:
            self._thread_log("=" * 60)
            source_row_count = cfg.get("source_row_count", len(cfg["rows"]))
            self._thread_log(
                f"BẮT ĐẦU: {source_row_count} dòng Excel, "
                f"{len(cfg['rows'])} lượt GCN, {len(groups)} nhóm duy nhất."
            )
            self._thread_log(
                f"Đẩy HSQ: {cfg['day_hsq']} | Thay mới theo GCN: {cfg['thay_moi_gcn']} | DRY_RUN: {core.DRY_RUN}"
            )

            for group_index, (group_key, items) in enumerate(groups, start=1):
                if self.processor.stop_flag:
                    self._emit("log", "Đã dừng theo yêu cầu.", "warn")
                    break

                representative = items[0]
                gcn_label = str(representative.get("sogcn") or "").strip()
                if representative.get("id_don"):
                    label = f"ID đơn {representative['id_don']}"
                else:
                    label = f"tờ {representative.get('soto') or '?'} / thửa {representative.get('sothua') or '?'}"
                if gcn_label:
                    label += f" / GCN {gcn_label}"
                self._emit("group_start", group_index, len(groups), label)
                self._thread_log(
                    f"--- [{group_index}/{len(groups)}] Dòng Excel {representative['excel_row']}: {label} ---"
                )

                pdf_skip_reason = representative.get("_pdf_skip_reason")
                if pdf_skip_reason:
                    self._thread_log(f"Bỏ qua dòng {representative['excel_row']}: {pdf_skip_reason}")
                    master_results = [
                        {
                            "soto": representative.get("soto", ""),
                            "sothua": representative.get("sothua", ""),
                            "loaidat": representative.get("loaidat", ""),
                            "tenfile": representative.get("tenfile", ""),
                            "sogcn": representative.get("sogcn", ""),
                            "tinhhinhdangkyid": representative.get("id_don", ""),
                            "hosoquetid": "",
                            "mota_moi": "",
                            "chu_su_dung": "",
                            "trang_thai": "Bỏ qua",
                            "ghi_chu": f"Bỏ qua — {pdf_skip_reason}",
                        }
                    ]
                else:
                    master_results = self.processor.process_row(
                        xa_id=cfg["xa_id"],
                        row=representative,
                        day_hsq=cfg["day_hsq"],
                        chu_id=cfg["chu_id"],
                        loai_ho_so_quet=cfg["loai_hsq"],
                        folder_upload=cfg["folder_upload"],
                        ngay_dang_ky_lan_dau=cfg["ngay_dk"],
                        thay_moi_gcn=cfg["thay_moi_gcn"],
                    )

                for result in master_results:
                    if (
                        result.get("trang_thai") == "Lỗi"
                        and "Không tìm thấy file PDF" in str(result.get("ghi_chu") or "")
                    ):
                        result["trang_thai"] = "Bỏ qua"
                        result["ghi_chu"] = f"Bỏ qua — {result.get('ghi_chu')}"

                for result in master_results:
                    writer.append_result(result)
                    total_results += 1
                    state = result.get("trang_thai")
                    if state == "Thành công":
                        success += 1
                    elif state == "Bỏ qua":
                        skipped += 1
                    else:
                        errors += 1

                if len(items) > 1:
                    self._thread_log(
                        f"Nhóm '{group_key}' có {len(items)} dòng; chỉ gọi API cho dòng đại diện."
                    )
                    for sibling in items[1:]:
                        for result in master_results:
                            copied = core.App._make_copy_result(
                                result, sibling, representative["excel_row"], group_key
                            )
                            writer.append_result(copied)
                            total_results += 1
                            state = copied.get("trang_thai")
                            if state == "Thành công":
                                success += 1
                            elif state == "Bỏ qua":
                                skipped += 1
                            else:
                                errors += 1

                processed += 1
                if processed % core.SAVE_EVERY_ROWS == 0:
                    writer.save()
                    self._thread_log(f"Đã tự lưu sau {processed} nhóm.")
                self._emit(
                    "progress",
                    processed,
                    len(groups),
                    success,
                    errors,
                    skipped,
                    total_results,
                )

            writer.save()
            self._thread_log(f"Đã lưu kết quả: {cfg['output_file']}")
        except PermissionError:
            fatal_error = "Không lưu được file kết quả; hãy đóng file nếu đang mở trong Excel."
            self._emit("log", fatal_error, "error")
        except Exception as exc:
            fatal_error = f"Lỗi chương trình: {exc}"
            self._emit("log", fatal_error, "error")
            try:
                writer.save()
                self._thread_log("Đã cố gắng lưu phần kết quả đang có.")
            except Exception as save_exc:
                self._emit("log", f"Không lưu được kết quả dở dang: {save_exc}", "error")
        finally:
            writer.close()
            self._emit(
                "done",
                processed,
                len(groups),
                success,
                errors,
                skipped,
                total_results,
                cfg["output_file"],
                fatal_error,
                self.processor.stop_flag,
            )

    async def _consume_messages(self) -> None:
        while True:
            item = await self.message_queue.get()
            kind = item[0]
            if kind == "log":
                self._append_log(item[1], item[2] if len(item) > 2 else "info")
            elif kind == "group_start":
                index, total, label = item[1], item[2], item[3]
                self.footer_status.value = f"Đang xử lý {label} ({index}/{total})"
                self.footer_status.update()
            elif kind == "progress":
                processed, total, success, errors, skipped, _results = item[1:]
                ratio = processed / total if total else 0
                self.metric_progress.value = f"{processed} / {total}"
                self.metric_success.value = str(success)
                self.metric_error.value = str(errors + skipped)
                self.progress_bar.value = ratio
                self.footer_percent.value = f"{int(ratio * 100)}%"
                self.page.update()
            elif kind == "done":
                (
                    processed,
                    total,
                    success,
                    errors,
                    skipped,
                    total_results,
                    output_file,
                    fatal_error,
                    stopped,
                ) = item[1:]
                self.is_running = False
                self.metric_progress.value = f"{processed} / {total}"
                self.metric_success.value = str(success)
                self.metric_error.value = str(errors + skipped)
                ratio = processed / total if total else 0
                self.progress_bar.value = ratio
                self.footer_percent.value = f"{int(ratio * 100)}%"

                if fatal_error:
                    status, color = "LỖI", CyberTheme.RED
                    self.footer_status.value = fatal_error
                elif stopped:
                    status, color = "ĐÃ DỪNG", CyberTheme.AMBER
                    self.footer_status.value = f"Đã dừng tại {processed}/{total} nhóm; kết quả đã được lưu."
                else:
                    status, color = "HOÀN TẤT", CyberTheme.GREEN
                    self.footer_status.value = (
                        f"Hoàn tất: {total_results} kết quả · {success} thành công · "
                        f"{errors} lỗi · {skipped} bỏ qua"
                    )

                self.status_text.value = status
                self.status_text.color = color
                self.status_icon.color = color
                self.status_badge.bgcolor = f"{color}15"
                self.status_badge.border = ft.Border.all(1, f"{color}55")

                # Khôi phục trạng thái các nút
                self.btn_start.gradient = ft.LinearGradient(
                    colors=[CyberTheme.GREEN, "#059669"],
                    begin=ft.Alignment(-1, 0),
                    end=ft.Alignment(1, 0),
                )
                self.btn_start.border = ft.Border.all(1, CyberTheme.GREEN)
                self.btn_start.content.controls[0].name = ft.Icons.PLAY_ARROW_ROUNDED
                self.btn_start.content.controls[0].color = CyberTheme.BG_DARK
                self.btn_start.content.controls[1].value = "BẮT ĐẦU CẬP NHẬT"
                self.btn_start.content.controls[1].color = CyberTheme.BG_DARK
                self.btn_start.ink = True

                if hasattr(self, "btn_header_start") and self.btn_header_start.page:
                    self.btn_header_start.bgcolor = CyberTheme.GREEN
                    self.btn_header_start.border = ft.Border.all(1, CyberTheme.GREEN)
                    self.btn_header_start.content.controls[0].name = ft.Icons.PLAY_ARROW_ROUNDED
                    self.btn_header_start.content.controls[0].color = CyberTheme.BG_DARK
                    self.btn_header_start.content.controls[1].value = "Chạy cập nhật"
                    self.btn_header_start.content.controls[1].color = CyberTheme.BG_DARK
                    self.btn_header_start.ink = True

                self.btn_stop.bgcolor = CyberTheme.SURFACE_HOVER
                self.btn_stop.border = ft.Border.all(1, CyberTheme.BORDER_COLOR)
                self.btn_stop.content.controls[0].color = CyberTheme.TEXT_MUTED
                self.btn_stop.content.controls[1].color = CyberTheme.TEXT_MUTED
                self.btn_stop.ink = False

                self.page.update()
                self._append_log(self.footer_status.value, "error" if fatal_error else ("warn" if stopped else "success"))
                if not fatal_error:
                    self._show_toast(f"Đã lưu kết quả tại {output_file}", success=not stopped)


async def flet_main(page: ft.Page) -> None:
    app = CapNhatDonHsqFletApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
