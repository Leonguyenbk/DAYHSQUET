import asyncio
import datetime
import os
import queue
import subprocess
import sys
import threading
from typing import Dict, List, Optional, Set

import flet as ft

import app_dir
import config as config_module
import db
import downloader
import excel_reader
import file_locator
import folder_map as folder_map_module
import net_use

APP_DIR = app_dir.get_app_dir()


# ==============================================================================
# Native Dialog Helpers (Non-blocking Windows Dialogs via Background Thread)
# ==============================================================================
def _native_choose_directory(initial_dir: Optional[str] = None) -> str:
    """Mở hộp thoại chọn thư mục Windows chuẩn (Native) mà không phụ thuộc Zenity."""
    import tkinter as tk
    from tkinter import filedialog

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(
            initialdir=initial_dir if initial_dir and os.path.exists(initial_dir) else APP_DIR,
            title="Chọn thư mục lưu kết quả tải về",
        )
        root.destroy()
        return selected or ""
    except Exception:
        return ""


def _native_choose_excel(initial_dir: Optional[str] = None) -> str:
    """Mở hộp thoại chọn file Excel .xlsx chuẩn Windows."""
    import tkinter as tk
    from tkinter import filedialog

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            initialdir=initial_dir if initial_dir and os.path.exists(initial_dir) else APP_DIR,
            title="Chọn file Excel chứa danh sách số giấy chứng nhận",
            filetypes=[("File Excel", "*.xlsx"), ("Tất cả tập tin", "*.*")],
        )
        root.destroy()
        return selected or ""
    except Exception:
        return ""


# ==============================================================================
# Theme & Color Palette (Futuristic Dark / Cyberpunk Clean Aesthetic)
# ==============================================================================
class CyberTheme:
    BG_DARK = "#090D16"
    SURFACE_DARK = "#111827"
    SURFACE_CARD = "#172033"
    SURFACE_HOVER = "#1E293B"
    SURFACE_INPUT = "#0D1322"
    BORDER_COLOR = "#23304B"
    BORDER_GLOW = "#00D2FF"

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
    TERMINAL_BORDER = "#1B253B"


# ==============================================================================
# Main Flet Application
# ==============================================================================
class TaiHoSoQuetApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.cfg = config_module.load_config()
        self.folder_map: Dict[str, str] = {}
        self._load_folder_map(silent=True)

        # Trạng thái tải & luồng
        self.cancel_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self.is_running = False

        # Dữ liệu danh sách số GCN
        self.cert_list: List[str] = []
        self.selected_indices: Set[int] = set()
        self.cert_status_map: Dict[str, dict] = {}  # {cert: {"status": ..., "ok": 0, "fail": 0}}
        self.search_filter_text = ""

        # Thống kê
        self.stat_total_certs = 0
        self.stat_processed_certs = 0
        self.stat_ok_files = 0
        self.stat_fail_files = 0

        # Nhật ký
        self.logs_history: List[dict] = []  # [{"time": ..., "text": ..., "level": ...}]
        self.log_filter_level = "ALL"  # ALL, OK, ERROR
        self.auto_scroll_enabled = True

        # Queue truyền thông điệp bất đồng bộ
        self.msg_queue: asyncio.Queue = asyncio.Queue()
        self.loop = asyncio.get_running_loop()

        # View hiện tại (0: Dashboard, 1: Cài đặt, 2: Hướng dẫn)
        self.current_nav_index = 0

    def _load_folder_map(self, silent=False) -> bool:
        """Đọc file CSV danh mục mã xã -> thư mục gốc."""
        try:
            files = [f for f in self.cfg["paths"]["folder_map_files"].split(",") if f.strip()]
            self.folder_map = folder_map_module.load_folder_map(files, base_dir=APP_DIR)
            return True
        except Exception as exc:
            self.folder_map = {}
            if not silent:
                self._show_toast(f"Lỗi đọc CSV thư mục: {exc}", is_error=True)
            return False

    def _emit(self, kind: str, *args):
        """Thread-safe emit từ background thread sang asyncio queue."""
        self.loop.call_soon_threadsafe(self.msg_queue.put_nowait, (kind, *args))

    def _show_toast(self, message: str, is_error: bool = False, is_success: bool = False):
        """Hiển thị thông báo Toast / SnackBar hiện đại."""
        color = CyberTheme.RED if is_error else (CyberTheme.GREEN if is_success else CyberTheme.BLUE)
        icon = (
            ft.Icons.ERROR_OUTLINE_ROUNDED
            if is_error
            else (ft.Icons.CHECK_CIRCLE_OUTLINE_ROUNDED if is_success else ft.Icons.INFO_OUTLINE_ROUNDED)
        )

        snack = ft.SnackBar(
            content=ft.Row(
                controls=[
                    ft.Icon(icon, color=color, size=20),
                    ft.Text(message, color=CyberTheme.TEXT_WHITE, size=13, weight=ft.FontWeight.W_500),
                ],
                spacing=10,
            ),
            bgcolor=CyberTheme.SURFACE_CARD,
            duration=3500,
            behavior=ft.SnackBarBehavior.FLOATING,
            margin=ft.Margin.all(16),
            shape=ft.RoundedRectangleBorder(radius=10),
        )
        self.page.show_dialog(snack)

    # ==========================================================================
    # Initialization & UI Composition
    # ==========================================================================
    async def init_ui(self):
        self.page.title = "TaiHoSoQuet • Hệ Thống Tải Hồ Sơ Quét Chuyên Nghiệp"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.bgcolor = CyberTheme.BG_DARK
        self.page.padding = 0
        self.page.spacing = 0
        self.page.window.width = 1240
        self.page.window.height = 840
        self.page.window.min_width = 1000
        self.page.window.min_height = 680
        await self.page.window.center()

        # Build các thành phần giao diện
        self._build_header()
        self._build_footer()
        self._build_dashboard_view()
        self._build_settings_view()
        self._build_guide_view()

        # Ghép toàn bộ trang
        self.main_container = ft.Container(
            expand=True,
            content=self.dashboard_view,  # Mặc định mở dashboard
            padding=ft.Padding.only(left=20, right=20, top=14, bottom=14),
        )

        self.page.add(
            ft.Column(
                expand=True,
                spacing=0,
                controls=[
                    self.header_bar,
                    self.main_container,
                    self.footer_progress_bar,
                ],
            )
        )

        # Chạy listener tiêu thụ tin nhắn từ background thread
        asyncio.create_task(self._process_message_queue())

        # Ghi log chào mừng hệ thống
        self._append_log(
            "=== Hệ thống Tải Hồ Sơ Quét đã sẵn sàng (Flet v{} | Python {}.{}) ===".format(
                getattr(ft, "__version__", "1.0"), sys.version_info.major, sys.version_info.minor
            ),
            level="title",
        )
        self._append_log(
            f"Đã nạp {len(self.folder_map)} mã xã từ danh mục CSV ({self.cfg['paths']['folder_map_files']}).",
            level="info",
        )
        self.page.update()

    # ==========================================================================
    # Header Bar & Navigation
    # ==========================================================================
    def _build_header(self):
        self.status_badge = ft.Container(
            content=ft.Row(
                spacing=6,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.RADIO_BUTTON_CHECKED_ROUNDED, color=CyberTheme.GREEN, size=14),
                    ft.Text("SẴN SÀNG", color=CyberTheme.GREEN, size=11, weight=ft.FontWeight.BOLD),
                ],
            ),
            padding=ft.Padding.symmetric(horizontal=10, vertical=4),
            border_radius=20,
            bgcolor=f"{CyberTheme.GREEN}15",
            border=ft.Border.all(1, f"{CyberTheme.GREEN}40"),
        )

        # Nút chuyển Tab (Pills)
        self.nav_btn_dash = self._create_nav_button("Bảng Điều Khiển", ft.Icons.DASHBOARD_ROUNDED, 0, active=True)
        self.nav_btn_sett = self._create_nav_button("Cài Đặt Kết Nối", ft.Icons.TUNE_ROUNDED, 1, active=False)
        self.nav_btn_help = self._create_nav_button("Hướng Dẫn", ft.Icons.HELP_OUTLINE_ROUNDED, 2, active=False)

        # Nút mở thư mục kết quả nhanh trên header
        btn_open_folder = ft.Container(
            content=ft.Row(
                spacing=6,
                controls=[
                    ft.Icon(ft.Icons.FOLDER_SPECIAL_ROUNDED, color=CyberTheme.CYAN, size=16),
                    ft.Text("Mở Downloads", color=CyberTheme.CYAN, size=12, weight=ft.FontWeight.W_600),
                ],
            ),
            padding=ft.Padding.symmetric(horizontal=12, vertical=6),
            border_radius=8,
            bgcolor=f"{CyberTheme.CYAN}15",
            border=ft.Border.all(1, f"{CyberTheme.CYAN}35"),
            ink=True,
            on_click=lambda e: self._open_dest_folder(),
            tooltip="Mở thư mục lưu kết quả tải về trong Windows Explorer",
        )

        self.header_bar = ft.Container(
            height=66,
            padding=ft.Padding.symmetric(horizontal=20),
            bgcolor=CyberTheme.SURFACE_DARK,
            border=ft.Border.only(bottom=ft.BorderSide(1, CyberTheme.BORDER_COLOR)),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    # Logo & Tên ứng dụng
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
                                content=ft.Icon(
                                    ft.Icons.CLOUD_DOWNLOAD_ROUNDED,
                                    color=CyberTheme.BG_DARK,
                                    size=22,
                                ),
                                alignment=ft.Alignment(0, 0),
                            ),
                            ft.Column(
                                spacing=0,
                                alignment=ft.MainAxisAlignment.CENTER,
                                controls=[
                                    ft.Text(
                                        "TẢI HỒ SƠ QUÉT",
                                        color=CyberTheme.TEXT_WHITE,
                                        size=16,
                                        weight=ft.FontWeight.BOLD,
                                    ),
                                    ft.Row(
                                        spacing=6,
                                        controls=[
                                            ft.Text("LIS SCAN SYNC", color=CyberTheme.CYAN, size=10, weight=ft.FontWeight.W_700),
                                            ft.Text("•", color=CyberTheme.TEXT_MUTED, size=10),
                                            ft.Text("Đắk Lắk", color=CyberTheme.TEXT_LIGHT, size=10),
                                        ],
                                    ),
                                ],
                            ),
                            ft.VerticalDivider(width=20, color=CyberTheme.BORDER_COLOR),
                            self.status_badge,
                        ],
                    ),
                    # Tabs Navigation
                    ft.Row(
                        spacing=6,
                        controls=[
                            self.nav_btn_dash,
                            self.nav_btn_sett,
                            self.nav_btn_help,
                        ],
                    ),
                    # Nút thao tác nhanh bên phải
                    ft.Row(
                        spacing=10,
                        controls=[
                            btn_open_folder,
                        ],
                    ),
                ],
            ),
        )

    def _create_nav_button(self, title: str, icon: str, index: int, active: bool) -> ft.Container:
        bg = f"{CyberTheme.CYAN}20" if active else "transparent"
        border_c = CyberTheme.CYAN if active else "transparent"
        txt_c = CyberTheme.CYAN if active else CyberTheme.TEXT_LIGHT

        btn = ft.Container(
            content=ft.Row(
                spacing=6,
                controls=[
                    ft.Icon(icon, color=txt_c, size=16),
                    ft.Text(title, color=txt_c, size=12, weight=ft.FontWeight.W_600),
                ],
            ),
            padding=ft.Padding.symmetric(horizontal=14, vertical=7),
            border_radius=8,
            bgcolor=bg,
            border=ft.Border.all(1, border_c),
            ink=True,
            on_click=lambda e, idx=index: self._switch_view(idx),
        )
        return btn

    def _switch_view(self, index: int):
        self.current_nav_index = index
        # Cập nhật style nút nav
        for i, btn in enumerate([self.nav_btn_dash, self.nav_btn_sett, self.nav_btn_help]):
            is_act = i == index
            btn.bgcolor = f"{CyberTheme.CYAN}20" if is_act else "transparent"
            btn.border = ft.Border.all(1, CyberTheme.CYAN if is_act else "transparent")
            row = btn.content
            row.controls[0].color = CyberTheme.CYAN if is_act else CyberTheme.TEXT_LIGHT
            row.controls[1].color = CyberTheme.CYAN if is_act else CyberTheme.TEXT_LIGHT
            btn.update()

        # Đổi nội dung container chính
        if index == 0:
            self.main_container.content = self.dashboard_view
        elif index == 1:
            self.main_container.content = self.settings_view
        else:
            self.main_container.content = self.guide_view
        self.main_container.update()

    # ==========================================================================
    # Dashboard View (Primary Workspace)
    # ==========================================================================
    def _build_dashboard_view(self):
        # 1. Thẻ Thống Kê (Metric Cards)
        self.metric_total_txt = ft.Text("0", size=24, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE)
        self.metric_proc_txt = ft.Text("0 / 0", size=24, weight=ft.FontWeight.BOLD, color=CyberTheme.CYAN)
        self.metric_ok_txt = ft.Text("0", size=24, weight=ft.FontWeight.BOLD, color=CyberTheme.GREEN)
        self.metric_fail_txt = ft.Text("0", size=24, weight=ft.FontWeight.BOLD, color=CyberTheme.RED)

        metrics_row = ft.Row(
            spacing=12,
            controls=[
                self._create_metric_card("TỔNG SỐ GIẤY", self.metric_total_txt, ft.Icons.FORMAT_LIST_NUMBERED_ROUNDED, CyberTheme.BLUE),
                self._create_metric_card("TIẾN ĐỘ XỬ LÝ", self.metric_proc_txt, ft.Icons.TIMELAPSE_ROUNDED, CyberTheme.CYAN),
                self._create_metric_card("FILE THÀNH CÔNG", self.metric_ok_txt, ft.Icons.CHECK_CIRCLE_ROUNDED, CyberTheme.GREEN),
                self._create_metric_card("FILE LỖI / BỎ QUA", self.metric_fail_txt, ft.Icons.ERROR_ROUNDED, CyberTheme.RED),
            ],
        )

        # 2. Thanh Cấu Hình Tải Nhanh (Destination & Threshold Bar)
        default_dest = os.path.join(APP_DIR, "downloads")
        self.dest_field = ft.TextField(
            value=default_dest,
            hint_text="Đường dẫn thư mục lưu file scan tải về...",
            prefix_icon=ft.Icons.FOLDER_ROUNDED,
            border_color=CyberTheme.BORDER_COLOR,
            focused_border_color=CyberTheme.CYAN,
            text_size=12,
            height=40,
            content_padding=ft.Padding.symmetric(horizontal=10, vertical=0),
            expand=True,
            bgcolor=CyberTheme.SURFACE_INPUT,
        )

        self.threshold_field = ft.TextField(
            value="0",
            hint_text="Ngưỡng ID",
            prefix_icon=ft.Icons.FILTER_LIST_ROUNDED,
            border_color=CyberTheme.BORDER_COLOR,
            focused_border_color=CyberTheme.CYAN,
            text_size=12,
            width=140,
            height=40,
            content_padding=ft.Padding.symmetric(horizontal=10, vertical=0),
            bgcolor=CyberTheme.SURFACE_INPUT,
            tooltip="Chỉ tải hồ sơ có HoSoQuetId lớn hơn giá trị này (mặc định: 0)",
        )

        # Nút Bắt Đầu / Dừng
        self.btn_start = ft.Container(
            content=ft.Row(
                spacing=8,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.PLAY_ARROW_ROUNDED, color=CyberTheme.BG_DARK, size=20),
                    ft.Text("BẮT ĐẦU TẢI", color=CyberTheme.BG_DARK, size=13, weight=ft.FontWeight.BOLD),
                ],
            ),
            width=150,
            height=40,
            border_radius=8,
            gradient=ft.LinearGradient(
                colors=[CyberTheme.GREEN, "#059669"],
                begin=ft.Alignment(-1, 0),
                end=ft.Alignment(1, 0),
            ),
            ink=True,
            on_click=lambda e: self._start_download_flow(),
        )

        self.btn_stop = ft.Container(
            content=ft.Row(
                spacing=8,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.STOP_ROUNDED, color=CyberTheme.TEXT_MUTED, size=20),
                    ft.Text("DỪNG", color=CyberTheme.TEXT_MUTED, size=13, weight=ft.FontWeight.BOLD),
                ],
            ),
            width=100,
            height=40,
            border_radius=8,
            bgcolor=CyberTheme.SURFACE_HOVER,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            ink=False,
            on_click=lambda e: self._stop_download_flow(),
        )

        config_bar = ft.Container(
            padding=14,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=10,
                controls=[
                    ft.Row(
                        spacing=10,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Text("Thư mục lưu kết quả:", size=12, weight=ft.FontWeight.W_600, color=CyberTheme.TEXT_LIGHT, width=135),
                            self.dest_field,
                            ft.IconButton(
                                icon=ft.Icons.FOLDER_OPEN_ROUNDED,
                                icon_color=CyberTheme.CYAN,
                                tooltip="Duyệt thư mục (Chọn nơi lưu)",
                                on_click=lambda e: self._choose_dest_directory(),
                            ),
                            ft.IconButton(
                                icon=ft.Icons.LAUNCH_ROUNDED,
                                icon_color=CyberTheme.BLUE,
                                tooltip="Mở thư mục này trong Windows Explorer",
                                on_click=lambda e: self._open_dest_folder(),
                            ),
                        ],
                    ),
                    ft.Row(
                        spacing=10,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Text("HoSoQuetId lớn hơn:", size=12, weight=ft.FontWeight.W_600, color=CyberTheme.TEXT_LIGHT, width=135),
                            self.threshold_field,
                            ft.Container(expand=True),
                            self.btn_start,
                            self.btn_stop,
                        ],
                    ),
                ],
            ),
        )

        # 3. Khu Vực Làm Việc Chính (Hai Cột: Danh Sách GCN & Nhật Ký Live)
        self._build_cert_list_panel()
        self._build_terminal_panel()

        workspace_row = ft.Row(
            expand=True,
            spacing=14,
            controls=[
                ft.Container(expand=4, content=self.cert_panel),
                ft.Container(expand=6, content=self.terminal_panel),
            ],
        )

        self.dashboard_view = ft.Column(
            expand=True,
            spacing=14,
            controls=[
                metrics_row,
                config_bar,
                workspace_row,
            ],
        )

    def _create_metric_card(self, title: str, val_control: ft.Text, icon: str, color: str) -> ft.Container:
        return ft.Container(
            expand=True,
            height=78,
            padding=ft.Padding.symmetric(horizontal=16, vertical=10),
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Column(
                        spacing=2,
                        alignment=ft.MainAxisAlignment.CENTER,
                        controls=[
                            ft.Text(title, size=11, weight=ft.FontWeight.W_700, color=CyberTheme.TEXT_MUTED),
                            val_control,
                        ],
                    ),
                    ft.Container(
                        width=42,
                        height=42,
                        border_radius=10,
                        bgcolor=f"{color}18",
                        border=ft.Border.all(1, f"{color}40"),
                        content=ft.Icon(icon, color=color, size=22),
                        alignment=ft.Alignment(0, 0),
                    ),
                ],
            ),
        )

    # --------------------------------------------------------------------------
    # Left Panel: Quản Lý Số Giấy Chứng Nhận
    # --------------------------------------------------------------------------
    def _build_cert_list_panel(self):
        self.cert_count_badge = ft.Text("0", size=11, weight=ft.FontWeight.BOLD, color=CyberTheme.CYAN)
        self.manual_input_field = ft.TextField(
            hint_text="Nhập số GCN rồi nhấn Enter...",
            text_size=12,
            height=38,
            content_padding=ft.Padding.symmetric(horizontal=10, vertical=0),
            expand=True,
            bgcolor=CyberTheme.SURFACE_INPUT,
            border_color=CyberTheme.BORDER_COLOR,
            focused_border_color=CyberTheme.CYAN,
            on_submit=lambda e: self._add_manual_cert(),
        )

        self.search_field = ft.TextField(
            hint_text="Tìm kiếm trong danh sách...",
            prefix_icon=ft.Icons.SEARCH_ROUNDED,
            text_size=11,
            height=34,
            content_padding=ft.Padding.symmetric(horizontal=8, vertical=0),
            expand=True,
            bgcolor=CyberTheme.SURFACE_INPUT,
            border_color=CyberTheme.BORDER_COLOR,
            focused_border_color=CyberTheme.CYAN,
            on_change=lambda e: self._on_search_change(e.control.value),
        )

        # Action Buttons
        btn_import_excel = self._create_mini_btn("Excel", ft.Icons.UPLOAD_FILE_ROUNDED, CyberTheme.GREEN, self._import_excel_flow)
        btn_batch_paste = self._create_mini_btn("Dán nhiều", ft.Icons.CONTENT_PASTE_ROUNDED, CyberTheme.BLUE, self._open_batch_dialog)
        btn_del_selected = self._create_mini_btn("Xóa chọn", ft.Icons.DELETE_OUTLINE_ROUNDED, CyberTheme.AMBER, self._remove_selected_certs)
        btn_clear_all = self._create_mini_btn("Xóa hết", ft.Icons.DELETE_SWEEP_ROUNDED, CyberTheme.RED, self._confirm_clear_all)

        # Danh sách dạng cuộn
        self.cert_list_view = ft.ListView(
            expand=True,
            spacing=4,
            padding=4,
            auto_scroll=False,
        )

        self.cert_panel = ft.Container(
            padding=12,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=8,
                controls=[
                    # Panel Header
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Row(
                                spacing=8,
                                controls=[
                                    ft.Icon(ft.Icons.DESCRIPTION_ROUNDED, color=CyberTheme.CYAN, size=18),
                                    ft.Text("Danh Sách Số GCN", size=13, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                                    ft.Container(
                                        content=self.cert_count_badge,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=2),
                                        border_radius=12,
                                        bgcolor=f"{CyberTheme.CYAN}20",
                                        border=ft.Border.all(1, f"{CyberTheme.CYAN}50"),
                                    ),
                                ],
                            ),
                        ],
                    ),
                    # Thêm thủ công
                    ft.Row(
                        spacing=6,
                        controls=[
                            self.manual_input_field,
                            ft.IconButton(
                                icon=ft.Icons.ADD_ROUNDED,
                                icon_color=CyberTheme.CYAN,
                                tooltip="Thêm vào danh sách",
                                on_click=lambda e: self._add_manual_cert(),
                            ),
                        ],
                    ),
                    # Toolbar Thao Tác
                    ft.Row(
                        spacing=6,
                        controls=[
                            btn_import_excel,
                            btn_batch_paste,
                            btn_del_selected,
                            btn_clear_all,
                        ],
                    ),
                    # Thanh tìm kiếm
                    ft.Row(
                        spacing=6,
                        controls=[
                            self.search_field,
                        ],
                    ),
                    ft.Divider(height=1, color=CyberTheme.BORDER_COLOR),
                    # Container chứa list
                    ft.Container(
                        expand=True,
                        border_radius=8,
                        bgcolor=CyberTheme.SURFACE_INPUT,
                        border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
                        content=self.cert_list_view,
                    ),
                ],
            ),
        )

    def _create_mini_btn(self, text: str, icon: str, color: str, callback) -> ft.Container:
        return ft.Container(
            content=ft.Row(
                spacing=4,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(icon, color=color, size=14),
                    ft.Text(text, color=color, size=11, weight=ft.FontWeight.W_600),
                ],
            ),
            padding=ft.Padding.symmetric(horizontal=8, vertical=5),
            border_radius=6,
            bgcolor=f"{color}15",
            border=ft.Border.all(1, f"{color}35"),
            ink=True,
            on_click=lambda e: callback(),
        )

    # --------------------------------------------------------------------------
    # Right Panel: Hacker Console / Live Terminal Logs
    # --------------------------------------------------------------------------
    def _build_terminal_panel(self):
        self.terminal_list_view = ft.ListView(
            expand=True,
            spacing=3,
            padding=10,
            auto_scroll=True,
        )

        self.log_count_txt = ft.Text("0 dòng", size=11, color=CyberTheme.TEXT_MUTED)

        # Bộ lọc nhật ký
        self.filter_chip_all = self._create_filter_chip("Tất cả", "ALL", active=True)
        self.filter_chip_ok = self._create_filter_chip("Thành công", "OK", active=False)
        self.filter_chip_err = self._create_filter_chip("Lỗi", "ERROR", active=False)

        btn_copy_logs = ft.IconButton(
            icon=ft.Icons.COPY_ALL_ROUNDED,
            icon_color=CyberTheme.TEXT_LIGHT,
            icon_size=16,
            tooltip="Sao chép toàn bộ nhật ký",
            on_click=lambda e: self._copy_logs_to_clipboard(),
        )
        btn_save_logs = ft.IconButton(
            icon=ft.Icons.SAVE_ALT_ROUNDED,
            icon_color=CyberTheme.TEXT_LIGHT,
            icon_size=16,
            tooltip="Xuất nhật ký ra file text",
            on_click=lambda e: self._export_logs_file(),
        )
        btn_clear_logs = ft.IconButton(
            icon=ft.Icons.DELETE_SWEEP_ROUNDED,
            icon_color=CyberTheme.RED,
            icon_size=16,
            tooltip="Xóa màn hình nhật ký",
            on_click=lambda e: self._clear_terminal(),
        )

        self.terminal_panel = ft.Container(
            padding=12,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=8,
                controls=[
                    # Terminal Header (Cyber / Mac dot style)
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Row(
                                spacing=8,
                                controls=[
                                    ft.Container(width=10, height=10, border_radius=5, bgcolor="#EF4444"),
                                    ft.Container(width=10, height=10, border_radius=5, bgcolor="#F59E0B"),
                                    ft.Container(width=10, height=10, border_radius=5, bgcolor="#10B981"),
                                    ft.Container(width=6),
                                    ft.Icon(ft.Icons.TERMINAL_ROUNDED, color=CyberTheme.CYAN, size=16),
                                    ft.Text("NHẬT KÝ HOẠT ĐỘNG (REALTIME)", size=12, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                                ],
                            ),
                            ft.Row(
                                spacing=4,
                                controls=[
                                    self.log_count_txt,
                                    btn_copy_logs,
                                    btn_save_logs,
                                    btn_clear_logs,
                                ],
                            ),
                        ],
                    ),
                    # Filter Bar
                    ft.Row(
                        spacing=6,
                        controls=[
                            self.filter_chip_all,
                            self.filter_chip_ok,
                            self.filter_chip_err,
                        ],
                    ),
                    # Màn hình Terminal
                    ft.Container(
                        expand=True,
                        border_radius=8,
                        bgcolor=CyberTheme.TERMINAL_BG,
                        border=ft.Border.all(1, CyberTheme.TERMINAL_BORDER),
                        content=self.terminal_list_view,
                    ),
                ],
            ),
        )

    def _create_filter_chip(self, label: str, level: str, active: bool) -> ft.Container:
        bg = f"{CyberTheme.CYAN}20" if active else "transparent"
        border_c = CyberTheme.CYAN if active else CyberTheme.BORDER_COLOR
        txt_c = CyberTheme.CYAN if active else CyberTheme.TEXT_MUTED

        return ft.Container(
            content=ft.Text(label, size=11, color=txt_c, weight=ft.FontWeight.W_600),
            padding=ft.Padding.symmetric(horizontal=10, vertical=4),
            border_radius=12,
            bgcolor=bg,
            border=ft.Border.all(1, border_c),
            ink=True,
            on_click=lambda e, lvl=level: self._set_log_filter(lvl),
        )

    def _set_log_filter(self, level: str):
        self.log_filter_level = level
        # Cập nhật style
        chips = [
            (self.filter_chip_all, "ALL"),
            (self.filter_chip_ok, "OK"),
            (self.filter_chip_err, "ERROR"),
        ]
        for chip, lvl in chips:
            is_act = lvl == level
            chip.bgcolor = f"{CyberTheme.CYAN}20" if is_act else "transparent"
            chip.border = ft.Border.all(1, CyberTheme.CYAN if is_act else CyberTheme.BORDER_COLOR)
            chip.content.color = CyberTheme.CYAN if is_act else CyberTheme.TEXT_MUTED
            chip.update()

        self._render_filtered_logs()

    # ==========================================================================
    # Footer Progress Bar & Live Status
    # ==========================================================================
    def _build_footer(self):
        self.progress_bar = ft.ProgressBar(
            value=0.0,
            color=CyberTheme.CYAN,
            bgcolor=CyberTheme.SURFACE_HOVER,
            height=4,
        )
        self.footer_status_txt = ft.Text(
            "Hệ thống sẵn sàng.",
            size=11,
            color=CyberTheme.TEXT_LIGHT,
            weight=ft.FontWeight.W_500,
        )
        self.footer_percent_txt = ft.Text("0%", size=11, color=CyberTheme.CYAN, weight=ft.FontWeight.BOLD)

        self.footer_progress_bar = ft.Container(
            height=34,
            padding=ft.Padding.symmetric(horizontal=20),
            bgcolor=CyberTheme.SURFACE_DARK,
            border=ft.Border.only(top=ft.BorderSide(1, CyberTheme.BORDER_COLOR)),
            content=ft.Column(
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=4,
                controls=[
                    self.progress_bar,
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            self.footer_status_txt,
                            self.footer_percent_txt,
                        ],
                    ),
                ],
            ),
        )

    # ==========================================================================
    # Settings View (Cài Đặt Kết Nối SQL Server & Share UNC)
    # ==========================================================================
    def _build_settings_view(self):
        db_cfg = self.cfg["database"]
        fs_cfg = self.cfg["fileserver"]

        # 1. Thẻ CSDL SQL Server
        self.sett_db_server = self._create_input("SQL Server (Host / IP)", db_cfg.get("server", ""), ft.Icons.DNS_ROUNDED)
        self.sett_db_port = self._create_input("Cổng (Port)", db_cfg.get("port", "1433"), ft.Icons.NUMBERS_ROUNDED, width=120)
        self.sett_db_database = self._create_input("Tên Database", db_cfg.get("database", "LIS"), ft.Icons.STORAGE_ROUNDED)
        self.sett_db_user = self._create_input("Username CSDL", db_cfg.get("username", "vpdktinh"), ft.Icons.PERSON_ROUNDED)
        self.sett_db_pass = self._create_input("Password CSDL", db_cfg.get("password", ""), ft.Icons.PASSWORD_ROUNDED, password=True)
        self.sett_db_driver = self._create_input("ODBC Driver", db_cfg.get("driver", "ODBC Driver 18 for SQL Server"), ft.Icons.TERMINAL_ROUNDED)

        self.db_test_result = ft.Text("", size=12, color=CyberTheme.TEXT_LIGHT)
        btn_test_sql = self._create_action_btn("Kiểm Tra Kết Nối SQL Server", ft.Icons.BOLT_ROUNDED, CyberTheme.BLUE, self._test_sql_connection)

        card_db = ft.Container(
            padding=16,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=12,
                controls=[
                    ft.Row(
                        spacing=8,
                        controls=[
                            ft.Icon(ft.Icons.STORAGE_ROUNDED, color=CyberTheme.CYAN, size=20),
                            ft.Text("Cơ Sở Dữ Liệu SQL Server (LIS)", size=14, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                        ],
                    ),
                    ft.Row(spacing=12, controls=[self.sett_db_server, self.sett_db_port, self.sett_db_database]),
                    ft.Row(spacing=12, controls=[self.sett_db_user, self.sett_db_pass, self.sett_db_driver]),
                    ft.Row(
                        spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[btn_test_sql, self.db_test_result],
                    ),
                ],
            ),
        )

        # 2. Thẻ Máy Chủ File Scan
        self.sett_fs_ip = self._create_input("IP Máy Chủ File Scan", fs_cfg.get("ip", ""), ft.Icons.ROUTER_ROUNDED)
        self.sett_fs_suffix = self._create_input("Hậu tố Admin Share", fs_cfg.get("share_suffix", "$"), ft.Icons.TAG_ROUNDED, width=120)
        self.sett_fs_map = self._create_input(
            "Ánh xạ Share (VD: E:\\BACKUP-SYN=HOSOQUET_02;...)",
            fs_cfg.get("share_map", ""),
            ft.Icons.HUB_ROUNDED,
            expand=True,
        )
        self.sett_fs_user = self._create_input("User Share (để trống nếu dùng phiên Windows)", fs_cfg.get("username", ""), ft.Icons.SECURITY_ROUNDED)
        self.sett_fs_pass = self._create_input("Password User Share", fs_cfg.get("password", ""), ft.Icons.PASSWORD_ROUNDED, password=True)

        self.fs_test_result = ft.Text("", size=12, color=CyberTheme.TEXT_LIGHT)
        btn_test_fs = self._create_action_btn("Kiểm Tra Kết Nối Máy Chủ File", ft.Icons.NETWORK_CHECK_ROUNDED, CyberTheme.PURPLE, self._test_fs_connection)

        card_fs = ft.Container(
            padding=16,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=12,
                controls=[
                    ft.Row(
                        spacing=8,
                        controls=[
                            ft.Icon(ft.Icons.FOLDER_SHARED_ROUNDED, color=CyberTheme.PURPLE, size=20),
                            ft.Text("Máy Chủ Lưu File Scan (UNC / SMB Share)", size=14, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                        ],
                    ),
                    ft.Row(spacing=12, controls=[self.sett_fs_ip, self.sett_fs_suffix]),
                    self.sett_fs_map,
                    ft.Row(spacing=12, controls=[self.sett_fs_user, self.sett_fs_pass]),
                    ft.Row(
                        spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[btn_test_fs, self.fs_test_result],
                    ),
                ],
            ),
        )

        # 3. Thẻ Danh Mục CSV
        self.sett_paths_csv = self._create_input("File CSV danh mục thư mục cấp 2", self.cfg["paths"].get("folder_map_files", ""), ft.Icons.TABLE_CHART_ROUNDED, expand=True)
        self.csv_status_txt = ft.Text(f"Hiện đã nạp: {len(self.folder_map)} mã xã.", size=12, color=CyberTheme.TEXT_LIGHT)
        btn_reload_csv = self._create_action_btn("Nạp Lại Danh Mục CSV", ft.Icons.REFRESH_ROUNDED, CyberTheme.AMBER, self._reload_folder_map_ui)

        card_csv = ft.Container(
            padding=16,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=12,
                controls=[
                    ft.Row(
                        spacing=8,
                        controls=[
                            ft.Icon(ft.Icons.DATA_OBJECT_ROUNDED, color=CyberTheme.AMBER, size=20),
                            ft.Text("Danh Mục & Ánh Xạ Thư Mục Cấp 2", size=14, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                        ],
                    ),
                    self.sett_paths_csv,
                    ft.Row(
                        spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[btn_reload_csv, self.csv_status_txt],
                    ),
                ],
            ),
        )

        # Footer Actions
        btn_save_config = ft.Container(
            content=ft.Row(
                spacing=8,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.SAVE_ROUNDED, color=CyberTheme.BG_DARK, size=18),
                    ft.Text("LƯU CẤU HÌNH", color=CyberTheme.BG_DARK, size=13, weight=ft.FontWeight.BOLD),
                ],
            ),
            width=160,
            height=42,
            border_radius=8,
            gradient=ft.LinearGradient(
                colors=[CyberTheme.CYAN, CyberTheme.BLUE],
                begin=ft.Alignment(-1, 0),
                end=ft.Alignment(1, 0),
            ),
            ink=True,
            on_click=lambda e: self._save_settings(),
        )

        btn_reset_defaults = ft.Container(
            content=ft.Row(
                spacing=8,
                alignment=ft.MainAxisAlignment.CENTER,
                controls=[
                    ft.Icon(ft.Icons.RESTORE_ROUNDED, color=CyberTheme.TEXT_LIGHT, size=18),
                    ft.Text("Khôi Phục Mặc Định", color=CyberTheme.TEXT_LIGHT, size=12, weight=ft.FontWeight.W_600),
                ],
            ),
            width=180,
            height=42,
            border_radius=8,
            bgcolor=CyberTheme.SURFACE_HOVER,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            ink=True,
            on_click=lambda e: self._reset_default_settings(),
        )

        self.settings_view = ft.ListView(
            expand=True,
            spacing=16,
            controls=[
                card_db,
                card_fs,
                card_csv,
                ft.Row(
                    alignment=ft.MainAxisAlignment.END,
                    spacing=12,
                    controls=[btn_reset_defaults, btn_save_config],
                ),
            ],
        )

    def _create_input(
        self,
        label: str,
        value: str,
        icon: str,
        width: Optional[int] = None,
        expand: bool = False,
        password: bool = False,
    ) -> ft.TextField:
        return ft.TextField(
            label=label,
            value=value,
            prefix_icon=icon,
            password=password,
            can_reveal_password=password,
            text_size=12,
            label_style=ft.TextStyle(size=12, color=CyberTheme.TEXT_LIGHT),
            bgcolor=CyberTheme.SURFACE_INPUT,
            border_color=CyberTheme.BORDER_COLOR,
            focused_border_color=CyberTheme.CYAN,
            height=48,
            width=width,
            expand=expand,
        )

    def _create_action_btn(self, text: str, icon: str, color: str, callback) -> ft.Container:
        return ft.Container(
            content=ft.Row(
                spacing=6,
                controls=[
                    ft.Icon(icon, color=color, size=16),
                    ft.Text(text, color=color, size=12, weight=ft.FontWeight.W_600),
                ],
            ),
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            border_radius=8,
            bgcolor=f"{color}15",
            border=ft.Border.all(1, f"{color}35"),
            ink=True,
            on_click=lambda e: callback(),
        )

    # ==========================================================================
    # Guide View (Hướng Dẫn & Kiến Trúc Phần Mềm)
    # ==========================================================================
    def _build_guide_view(self):
        guide_card = ft.Container(
            padding=24,
            border_radius=12,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            content=ft.Column(
                spacing=16,
                controls=[
                    ft.Row(
                        spacing=10,
                        controls=[
                            ft.Icon(ft.Icons.MENU_BOOK_ROUNDED, color=CyberTheme.CYAN, size=24),
                            ft.Text("HƯỚNG DẪN SỬ DỤNG VÀ NGUYÊN LÝ HOẠT ĐỘNG", size=16, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                        ],
                    ),
                    ft.Divider(color=CyberTheme.BORDER_COLOR),
                    self._create_guide_step(
                        "1",
                        "Chuẩn bị danh sách số Giấy Chứng Nhận",
                        "Bạn có thể nhập từng số vào ô nhập liệu hoặc bấm nút 'Excel' để tải danh sách từ cột A của file .xlsx, hoặc bấm 'Dán nhiều' để dán hàng loạt số giấy chứng nhận.",
                    ),
                    self._create_guide_step(
                        "2",
                        "Cấu hình thư mục lưu kết quả & ngưỡng ID",
                        "Chọn thư mục sẽ lưu file tải về (mặc định là thư mục downloads cùng cấp với chương trình). Nhập ngưỡng HoSoQuetId (mặc định là 0 để tải tất cả file mới).",
                    ),
                    self._create_guide_step(
                        "3",
                        "Khởi động tiến trình tải",
                        "Bấm 'BẮT ĐẦU TẢI'. Phần mềm sẽ tự động kết nối SQL Server LIS, tra cứu bảng HoSoQuet và các bảng Paper tương ứng, sau đó sao chép file scan qua đường dẫn mạng UNC về máy bạn.",
                    ),
                    self._create_guide_step(
                        "4",
                        "Kiểm tra & theo dõi",
                        "Toàn bộ diễn biến được hiển thị thời gian thực trên màn hình Live Terminal. Khi tải xong, bạn có thể bấm 'Mở Downloads' trên thanh tiêu đề để xem toàn bộ file đã lưu.",
                    ),
                ],
            ),
        )

        self.guide_view = ft.ListView(
            expand=True,
            spacing=16,
            controls=[guide_card],
        )

    def _create_guide_step(self, num: str, title: str, desc: str) -> ft.Row:
        return ft.Row(
            spacing=14,
            vertical_alignment=ft.CrossAxisAlignment.START,
            controls=[
                ft.Container(
                    width=28,
                    height=28,
                    border_radius=14,
                    bgcolor=f"{CyberTheme.CYAN}20",
                    border=ft.Border.all(1, CyberTheme.CYAN),
                    content=ft.Text(num, size=12, weight=ft.FontWeight.BOLD, color=CyberTheme.CYAN),
                    alignment=ft.Alignment(0, 0),
                ),
                ft.Column(
                    spacing=4,
                    expand=True,
                    controls=[
                        ft.Text(title, size=13, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                        ft.Text(desc, size=12, color=CyberTheme.TEXT_LIGHT),
                    ],
                ),
            ],
        )

    # ==========================================================================
    # Logic: Quản Lý Số Giấy Chứng Nhận (Thêm, Xóa, Nhập Excel, Dán)
    # ==========================================================================
    def _add_manual_cert(self):
        text = self.manual_input_field.value.strip()
        if not text:
            return
        added = 0
        for part in text.replace(",", "\n").split("\n"):
            part = part.strip()
            if part and part not in self.cert_list:
                self.cert_list.append(part)
                self.cert_status_map[part] = {"status": "pending", "ok": 0, "fail": 0}
                added += 1

        self.manual_input_field.value = ""
        self.manual_input_field.update()
        if added > 0:
            self._update_cert_list_ui()
            self._show_toast(f"Đã thêm {added} số giấy chứng nhận.", is_success=True)

    def _open_batch_dialog(self):
        """Mở dialog dán danh sách hàng loạt."""
        batch_text_field = ft.TextField(
            multiline=True,
            min_lines=8,
            max_lines=12,
            hint_text="Dán danh sách số giấy chứng nhận tại đây (mỗi số một dòng hoặc cách nhau bởi dấu phẩy, khoảng trắng)...",
            text_size=12,
            bgcolor=CyberTheme.SURFACE_INPUT,
            border_color=CyberTheme.BORDER_COLOR,
            focused_border_color=CyberTheme.CYAN,
        )

        def on_confirm(e):
            raw = batch_text_field.value or ""
            parts = [p.strip() for p in raw.replace(",", "\n").replace("\t", "\n").split("\n")]
            added = 0
            for p in parts:
                if p and p not in self.cert_list:
                    self.cert_list.append(p)
                    self.cert_status_map[p] = {"status": "pending", "ok": 0, "fail": 0}
                    added += 1

            self.page.pop_dialog()
            if added > 0:
                self._update_cert_list_ui()
                self._show_toast(f"Đã thêm {added} số giấy chứng nhận từ văn bản dán.", is_success=True)

        dlg = ft.AlertDialog(
            title=ft.Row(
                spacing=8,
                controls=[
                    ft.Icon(ft.Icons.CONTENT_PASTE_ROUNDED, color=CyberTheme.CYAN, size=20),
                    ft.Text("Dán Hàng Loạt Số GCN", size=15, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                ],
            ),
            content=ft.Container(
                width=460,
                content=batch_text_field,
            ),
            actions=[
                ft.TextButton("Hủy", on_click=lambda e: self.page.pop_dialog()),
                ft.ElevatedButton("Thêm Vào Danh Sách", bgcolor=CyberTheme.CYAN, color=CyberTheme.BG_DARK, on_click=on_confirm),
            ],
            bgcolor=CyberTheme.SURFACE_CARD,
        )
        self.page.show_dialog(dlg)

    def _import_excel_flow(self):
        """Mở hộp thoại chọn file Excel trong thread nền."""
        asyncio.create_task(self._async_import_excel())

    async def _async_import_excel(self):
        path = await asyncio.to_thread(_native_choose_excel, APP_DIR)
        if not path:
            return

        try:
            values = await asyncio.to_thread(excel_reader.read_certificates_from_excel, path)
        except Exception as exc:
            self._show_toast(f"Lỗi đọc file Excel: {exc}", is_error=True)
            return

        added = 0
        for v in values:
            if v and v not in self.cert_list:
                self.cert_list.append(v)
                self.cert_status_map[v] = {"status": "pending", "ok": 0, "fail": 0}
                added += 1

        self._update_cert_list_ui()
        self._show_toast(f"Đã nạp thành công {added} số giấy chứng nhận từ Excel.", is_success=True)
        self._append_log(f"Đã nhập {len(values)} dòng từ file Excel: {os.path.basename(path)} ({added} số mới).", level="info")

    def _remove_selected_certs(self):
        if not self.selected_indices:
            self._show_toast("Chưa chọn mục nào để xóa.", is_error=False)
            return

        # Lọc danh sách giữ lại các mục không được chọn
        new_list = [cert for i, cert in enumerate(self.cert_list) if i not in self.selected_indices]
        removed_count = len(self.cert_list) - len(new_list)
        self.cert_list = new_list
        self.selected_indices.clear()
        self._update_cert_list_ui()
        self._show_toast(f"Đã xóa {removed_count} số giấy chứng nhận.", is_success=True)

    def _confirm_clear_all(self):
        if not self.cert_list:
            return

        def on_confirm(e):
            self.cert_list.clear()
            self.selected_indices.clear()
            self.cert_status_map.clear()
            self.page.pop_dialog()
            self._update_cert_list_ui()
            self._show_toast("Đã xóa toàn bộ danh sách.", is_success=True)

        dlg = ft.AlertDialog(
            title=ft.Row(
                spacing=8,
                controls=[
                    ft.Icon(ft.Icons.WARNING_ROUNDED, color=CyberTheme.AMBER, size=20),
                    ft.Text("Xác Nhận Xóa Hết", size=15, weight=ft.FontWeight.BOLD, color=CyberTheme.TEXT_WHITE),
                ],
            ),
            content=ft.Text("Bạn có chắc chắn muốn xóa toàn bộ danh sách số giấy chứng nhận không?"),
            actions=[
                ft.TextButton("Hủy", on_click=lambda e: self.page.pop_dialog()),
                ft.ElevatedButton("Xóa Tất Cả", bgcolor=CyberTheme.RED, color=CyberTheme.TEXT_WHITE, on_click=on_confirm),
            ],
            bgcolor=CyberTheme.SURFACE_CARD,
        )
        self.page.show_dialog(dlg)

    def _on_search_change(self, query: str):
        self.search_filter_text = (query or "").strip().lower()
        self._render_cert_list()

    def _update_cert_list_ui(self):
        self.cert_count_badge.value = str(len(self.cert_list))
        self.metric_total_txt.value = str(len(self.cert_list))
        self.cert_count_badge.update()
        self.metric_total_txt.update()
        self._render_cert_list()

    def _render_cert_list(self):
        self.cert_list_view.controls.clear()

        query = self.search_filter_text
        for i, cert in enumerate(self.cert_list):
            if query and query not in cert.lower():
                continue

            status_info = self.cert_status_map.get(cert, {"status": "pending", "ok": 0, "fail": 0})
            status = status_info.get("status", "pending")
            ok_cnt = status_info.get("ok", 0)
            fail_cnt = status_info.get("fail", 0)

            # Badge trạng thái từng dòng
            if status == "running":
                badge = ft.Container(
                    content=ft.Row(
                        spacing=4,
                        controls=[
                            ft.ProgressRing(width=10, height=10, stroke_width=2, color=CyberTheme.CYAN),
                            ft.Text("Đang tải...", size=10, color=CyberTheme.CYAN, weight=ft.FontWeight.BOLD),
                        ],
                    ),
                    padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                    border_radius=10,
                    bgcolor=f"{CyberTheme.CYAN}15",
                )
            elif status == "ok":
                badge = ft.Container(
                    content=ft.Text(f"✓ Xong ({ok_cnt})", size=10, color=CyberTheme.GREEN, weight=ft.FontWeight.BOLD),
                    padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                    border_radius=10,
                    bgcolor=f"{CyberTheme.GREEN}15",
                )
            elif status == "error":
                badge = ft.Container(
                    content=ft.Text(f"✕ Lỗi ({fail_cnt})", size=10, color=CyberTheme.RED, weight=ft.FontWeight.BOLD),
                    padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                    border_radius=10,
                    bgcolor=f"{CyberTheme.RED}15",
                )
            elif status == "empty":
                badge = ft.Container(
                    content=ft.Text("0 file mới", size=10, color=CyberTheme.AMBER, weight=ft.FontWeight.W_500),
                    padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                    border_radius=10,
                    bgcolor=f"{CyberTheme.AMBER}15",
                )
            else:
                badge = ft.Container(
                    content=ft.Text("Chờ", size=10, color=CyberTheme.TEXT_MUTED),
                    padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                    border_radius=10,
                    bgcolor=CyberTheme.SURFACE_HOVER,
                )

            is_selected = i in self.selected_indices
            chk = ft.Checkbox(
                value=is_selected,
                scale=0.75,
                fill_color=CyberTheme.CYAN,
                check_color=CyberTheme.BG_DARK,
                on_change=lambda e, idx=i: self._toggle_cert_selection(idx, e.control.value),
            )

            btn_del_single = ft.IconButton(
                icon=ft.Icons.CLOSE_ROUNDED,
                icon_color=CyberTheme.TEXT_MUTED,
                icon_size=14,
                tooltip="Xóa số này",
                on_click=lambda e, idx=i: self._delete_single_cert(idx),
            )

            row_container = ft.Container(
                padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                border_radius=6,
                bgcolor=CyberTheme.SURFACE_CARD if not is_selected else f"{CyberTheme.CYAN}10",
                border=ft.Border.all(1, CyberTheme.BORDER_COLOR if not is_selected else f"{CyberTheme.CYAN}40"),
                content=ft.Row(
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Row(
                            spacing=4,
                            controls=[
                                chk,
                                ft.Text(f"#{i+1}", size=10, color=CyberTheme.TEXT_MUTED, width=28),
                                ft.Text(cert, size=12, weight=ft.FontWeight.W_600, color=CyberTheme.TEXT_WHITE),
                            ],
                        ),
                        ft.Row(
                            spacing=4,
                            controls=[
                                badge,
                                btn_del_single,
                            ],
                        ),
                    ],
                ),
            )
            self.cert_list_view.controls.append(row_container)

        self.cert_list_view.update()

    def _toggle_cert_selection(self, index: int, selected: bool):
        if selected:
            self.selected_indices.add(index)
        else:
            self.selected_indices.discard(index)

    def _delete_single_cert(self, index: int):
        if 0 <= index < len(self.cert_list):
            cert = self.cert_list.pop(index)
            self.selected_indices.discard(index)
            self.cert_status_map.pop(cert, None)
            self._update_cert_list_ui()

    # ==========================================================================
    # Logic: Terminal Logs
    # ==========================================================================
    def _append_log(self, text: str, level: str = "info"):
        now_str = datetime.datetime.now().strftime("%H:%M:%S")
        self.logs_history.append({"time": now_str, "text": text, "level": level})
        self.log_count_txt.value = f"{len(self.logs_history)} dòng"
        self.log_count_txt.update()

        # Render nếu khớp bộ lọc
        if self._matches_filter(level):
            log_item_widget = self._create_log_item_widget(now_str, text, level)
            self.terminal_list_view.controls.append(log_item_widget)
            self.terminal_list_view.update()

    def _matches_filter(self, level: str) -> bool:
        if self.log_filter_level == "ALL":
            return True
        if self.log_filter_level == "OK" and level in ("ok", "success"):
            return True
        if self.log_filter_level == "ERROR" and level in ("error", "fail"):
            return True
        return False

    def _render_filtered_logs(self):
        self.terminal_list_view.controls.clear()
        for item in self.logs_history:
            if self._matches_filter(item["level"]):
                widget = self._create_log_item_widget(item["time"], item["text"], item["level"])
                self.terminal_list_view.controls.append(widget)
        self.terminal_list_view.update()

    def _create_log_item_widget(self, time_str: str, text: str, level: str) -> ft.Container:
        c_tag = CyberTheme.CYAN
        c_text = CyberTheme.TEXT_WHITE
        bg_tag = f"{CyberTheme.CYAN}20"
        tag_text = "INFO"

        lower_t = text.lower()
        if level == "title" or text.startswith("==="):
            c_tag = CyberTheme.CYAN
            c_text = CyberTheme.CYAN
            bg_tag = f"{CyberTheme.CYAN}25"
            tag_text = "EXEC"
        elif level in ("ok", "success") or " ok: " in lower_t or "thanh cong" in lower_t:
            c_tag = CyberTheme.GREEN
            c_text = CyberTheme.TEXT_WHITE
            bg_tag = f"{CyberTheme.GREEN}25"
            tag_text = " OK "
        elif level in ("error", "fail") or "loi" in lower_t or "fail" in lower_t or "error" in lower_t:
            c_tag = CyberTheme.RED
            c_text = CyberTheme.RED
            bg_tag = f"{CyberTheme.RED}25"
            tag_text = "FAIL"
        elif "khong tim thay" in lower_t or "khong co file" in lower_t:
            c_tag = CyberTheme.AMBER
            c_text = CyberTheme.TEXT_LIGHT
            bg_tag = f"{CyberTheme.AMBER}25"
            tag_text = "WARN"

        return ft.Container(
            padding=ft.Padding.symmetric(vertical=1),
            content=ft.Row(
                spacing=8,
                vertical_alignment=ft.CrossAxisAlignment.START,
                controls=[
                    ft.Text(f"[{time_str}]", size=11, color=CyberTheme.TEXT_MUTED, font_family="Consolas"),
                    ft.Container(
                        content=ft.Text(tag_text, size=9, weight=ft.FontWeight.BOLD, color=c_tag, font_family="Consolas"),
                        padding=ft.Padding.symmetric(horizontal=4, vertical=1),
                        border_radius=4,
                        bgcolor=bg_tag,
                    ),
                    ft.Text(
                        text,
                        size=11,
                        color=c_text,
                        font_family="Consolas",
                        selectable=True,
                        expand=True,
                    ),
                ],
            ),
        )

    def _clear_terminal(self):
        self.logs_history.clear()
        self.terminal_list_view.controls.clear()
        self.log_count_txt.value = "0 dòng"
        self.log_count_txt.update()
        self.terminal_list_view.update()
        self._show_toast("Đã xóa sạch màn hình nhật ký.", is_success=True)

    def _copy_logs_to_clipboard(self):
        if not self.logs_history:
            self._show_toast("Nhật ký trống, không có nội dung để sao chép.", is_error=False)
            return

        lines = [f"[{item['time']}] [{item['level'].upper()}] {item['text']}" for item in self.logs_history]
        all_text = "\n".join(lines)

        try:
            subprocess.run(
                ["clip"],
                input=all_text.encode("utf-8"),
                check=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
            )
            self._show_toast("Đã sao chép toàn bộ nhật ký vào Clipboard!", is_success=True)
        except Exception as exc:
            self._show_toast(f"Lỗi sao chép: {exc}", is_error=True)

    def _export_logs_file(self):
        if not self.logs_history:
            self._show_toast("Nhật ký trống.", is_error=False)
            return

        now_tag = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dest_dir = self.dest_field.value.strip() or os.path.join(APP_DIR, "downloads")
        os.makedirs(dest_dir, exist_ok=True)
        file_path = os.path.join(dest_dir, f"TaiHoSoQuet_Log_{now_tag}.txt")

        try:
            lines = [f"[{item['time']}] [{item['level'].upper()}] {item['text']}\n" for item in self.logs_history]
            with open(file_path, "w", encoding="utf-8") as f:
                f.writelines(lines)

            self._show_toast(f"Đã xuất log ra: {os.path.basename(file_path)}", is_success=True)
        except Exception as exc:
            self._show_toast(f"Lỗi xuất file log: {exc}", is_error=True)

    # ==========================================================================
    # Logic: Folder Browser & Opener
    # ==========================================================================
    def _choose_dest_directory(self):
        asyncio.create_task(self._async_choose_dir())

    async def _async_choose_dir(self):
        curr = self.dest_field.value.strip()
        selected = await asyncio.to_thread(_native_choose_directory, curr)
        if selected:
            self.dest_field.value = selected
            self.dest_field.update()
            self._show_toast(f"Đã chọn thư mục: {selected}", is_success=True)

    def _open_dest_folder(self):
        path = self.dest_field.value.strip() or os.path.join(APP_DIR, "downloads")
        os.makedirs(path, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            self._show_toast(f"Không thể mở thư mục: {exc}", is_error=True)

    # ==========================================================================
    # Download Execution Flow (Background Worker & Event Loop Bridge)
    # ==========================================================================
    def _start_download_flow(self):
        if self.is_running:
            return

        if not self.cert_list:
            self._show_toast("Chưa có số giấy chứng nhận nào trong danh sách!", is_error=True)
            return

        try:
            threshold = int(self.threshold_field.value.strip() or "0")
        except ValueError:
            self._show_toast("Ngưỡng HoSoQuetId phải là số nguyên hợp lệ!", is_error=True)
            return

        dest_root = self.dest_field.value.strip()
        if not dest_root:
            self._show_toast("Vui lòng chọn hoặc nhập thư mục lưu kết quả!", is_error=True)
            return
        os.makedirs(dest_root, exist_ok=True)

        if not self.folder_map:
            self._load_folder_map(silent=True)
            if not self.folder_map:
                self._show_toast("Chưa nạp được danh mục thư mục từ file CSV!", is_error=True)
                return

        server_ip = self.cfg["fileserver"]["ip"].strip()
        if not server_ip:
            self._show_toast("Chưa nhập IP máy chủ trong Cài đặt kết nối!", is_error=True)
            return

        # Đổi trạng thái UI
        self.is_running = True
        self.cancel_event.clear()

        self.btn_start.gradient = None
        self.btn_start.bgcolor = CyberTheme.SURFACE_HOVER
        self.btn_start.content.controls[0].color = CyberTheme.TEXT_MUTED
        self.btn_start.content.controls[1].color = CyberTheme.TEXT_MUTED
        self.btn_start.ink = False
        self.btn_start.update()

        self.btn_stop.bgcolor = f"{CyberTheme.RED}25"
        self.btn_stop.border = ft.Border.all(1, CyberTheme.RED)
        self.btn_stop.content.controls[0].color = CyberTheme.RED
        self.btn_stop.content.controls[1].color = CyberTheme.RED
        self.btn_stop.ink = True
        self.btn_stop.update()

        self.status_badge.bgcolor = f"{CyberTheme.CYAN}20"
        self.status_badge.border = ft.Border.all(1, CyberTheme.CYAN)
        self.status_badge.content.controls[0].color = CyberTheme.CYAN
        self.status_badge.content.controls[1].color = CyberTheme.CYAN
        self.status_badge.content.controls[1].value = "ĐANG TẢI..."
        self.status_badge.update()

        self.progress_bar.value = 0.0
        self.progress_bar.update()

        self.stat_total_certs = len(self.cert_list)
        self.stat_processed_certs = 0
        self.stat_ok_files = 0
        self.stat_fail_files = 0

        self.metric_proc_txt.value = f"0 / {self.stat_total_certs}"
        self.metric_ok_txt.value = "0"
        self.metric_fail_txt.value = "0"
        self.metric_proc_txt.update()
        self.metric_ok_txt.update()
        self.metric_fail_txt.update()

        # Reset status trong map
        for c in self.cert_list:
            self.cert_status_map[c] = {"status": "pending", "ok": 0, "fail": 0}
        self._render_cert_list()

        certs_to_download = list(self.cert_list)
        self._append_log(
            f"=== Bắt đầu tải {len(certs_to_download)} số giấy chứng nhận (Ngưỡng HoSoQuetId > {threshold}) ===",
            level="title",
        )

        # Khởi chạy luồng Worker (giữ 100% logic nguyên bản)
        self.worker_thread = threading.Thread(
            target=self._run_worker_thread,
            args=(certs_to_download, threshold, dest_root, server_ip),
            daemon=True,
        )
        self.worker_thread.start()

    def _stop_download_flow(self):
        if not self.is_running:
            return
        self.cancel_event.set()
        self._append_log("Đã gửi yêu cầu dừng, sẽ kết thúc sau khi xong số giấy chứng nhận hiện tại...", level="warn")
        self.footer_status_txt.value = "Đang dừng tiến trình..."
        self.footer_status_txt.update()

    def _run_worker_thread(self, certs: List[str], threshold: int, dest_root: str, server_ip: str):
        fs_cfg = self.cfg["fileserver"]
        share_suffix = fs_cfg["share_suffix"].strip() or "$"
        share_map = file_locator.parse_share_map(fs_cfg.get("share_map", ""))
        share_user = fs_cfg.get("username", "").strip()
        share_pass = fs_cfg.get("password", "")

        def log(msg):
            self._emit("log", msg)

        share_names = sorted({
            file_locator.share_name_for_path(p, share_suffix, share_map)
            for p in self.folder_map.values()
        })

        if share_user:
            for share_name, ok, msg in net_use.connect_all(server_ip, share_names, share_user, share_pass):
                if ok:
                    log(f"Đã đăng nhập share \\\\{server_ip}\\{share_name} bằng user {share_user}.")
                else:
                    log(f"Không đăng nhập được share \\\\{server_ip}\\{share_name}: {msg}")

        try:
            conn = db.connect(self.cfg, retries=3, retry_delay=3, log=log)
        except Exception as exc:
            log(f"Không kết nối được SQL Server sau nhiều lần thử: {exc}")
            self._emit("done", 0, 0)
            return

        total_ok = 0
        total_fail = 0
        try:
            for i, cert in enumerate(certs, start=1):
                if self.cancel_event.is_set():
                    log("Đã dừng theo yêu cầu của người dùng.")
                    break

                self._emit("cert_start", cert, i, len(certs))

                ok, fail = downloader.download_for_certificate(
                    conn, cert, threshold, self.folder_map, server_ip, share_suffix,
                    dest_root, log, share_map,
                )
                total_ok += ok
                total_fail += fail

                self._emit("cert_done", cert, ok, fail, i, len(certs), total_ok, total_fail)
        finally:
            conn.close()
            if share_user:
                net_use.disconnect_all(server_ip, share_names)

        self._emit("done", total_ok, total_fail)

    async def _process_message_queue(self):
        """Vòng lặp tiêu thụ thông điệp từ worker thread và cập nhật giao diện bất đồng bộ."""
        while True:
            item = await self.msg_queue.get()
            kind = item[0]

            if kind == "log":
                self._append_log(item[1])

            elif kind == "cert_start":
                cert, idx, total = item[1], item[2], item[3]
                self.cert_status_map[cert] = {"status": "running", "ok": 0, "fail": 0}
                self.footer_status_txt.value = f"Đang xử lý GCN {cert} ({idx}/{total})..."
                self.footer_status_txt.update()
                self._render_cert_list()

            elif kind == "cert_done":
                cert, ok, fail, idx, total, tot_ok, tot_fail = item[1], item[2], item[3], item[4], item[5], item[6], item[7]
                status = "ok" if (ok > 0 and fail == 0) else ("error" if fail > 0 else "empty")
                self.cert_status_map[cert] = {"status": status, "ok": ok, "fail": fail}

                self.stat_processed_certs = idx
                self.stat_ok_files = tot_ok
                self.stat_fail_files = tot_fail

                pct = idx / total if total > 0 else 0
                self.progress_bar.value = pct
                self.footer_percent_txt.value = f"{int(pct * 100)}%"
                self.metric_proc_txt.value = f"{idx} / {total}"
                self.metric_ok_txt.value = str(tot_ok)
                self.metric_fail_txt.value = str(tot_fail)

                self.progress_bar.update()
                self.footer_percent_txt.update()
                self.metric_proc_txt.update()
                self.metric_ok_txt.update()
                self.metric_fail_txt.update()
                self._render_cert_list()

            elif kind == "done":
                tot_ok, tot_fail = item[1], item[2]
                self._on_download_completed(tot_ok, tot_fail)

    def _on_download_completed(self, total_ok: int, total_fail: int):
        self.is_running = False

        # Khôi phục trạng thái các nút
        self.btn_start.gradient = ft.LinearGradient(
            colors=[CyberTheme.GREEN, "#059669"],
            begin=ft.Alignment(-1, 0),
            end=ft.Alignment(1, 0),
        )
        self.btn_start.content.controls[0].color = CyberTheme.BG_DARK
        self.btn_start.content.controls[1].color = CyberTheme.BG_DARK
        self.btn_start.ink = True
        self.btn_start.update()

        self.btn_stop.bgcolor = CyberTheme.SURFACE_HOVER
        self.btn_stop.border = ft.Border.all(1, CyberTheme.BORDER_COLOR)
        self.btn_stop.content.controls[0].color = CyberTheme.TEXT_MUTED
        self.btn_stop.content.controls[1].color = CyberTheme.TEXT_MUTED
        self.btn_stop.ink = False
        self.btn_stop.update()

        status_text = "ĐÃ DỪNG" if self.cancel_event.is_set() else "HOÀN TẤT"
        status_color = CyberTheme.AMBER if self.cancel_event.is_set() else CyberTheme.GREEN

        self.status_badge.bgcolor = f"{status_color}20"
        self.status_badge.border = ft.Border.all(1, status_color)
        self.status_badge.content.controls[0].color = status_color
        self.status_badge.content.controls[1].color = status_color
        self.status_badge.content.controls[1].value = status_text
        self.status_badge.update()

        self.footer_status_txt.value = f"Đã hoàn thành: {total_ok} file thành công, {total_fail} file lỗi."
        self.footer_status_txt.update()

        self._append_log(
            f"=== Hoàn tất tiến trình: {total_ok} file tải thành công, {total_fail} file lỗi ===",
            level="title" if total_fail == 0 else "warn",
        )
        self._show_toast(f"Đã hoàn tất tải: {total_ok} file thành công!", is_success=True)

    # ==========================================================================
    # Logic: Settings, Save, Reset & Connection Diagnostics
    # ==========================================================================
    def _save_settings(self):
        self.cfg["database"]["server"] = self.sett_db_server.value.strip()
        self.cfg["database"]["port"] = self.sett_db_port.value.strip()
        self.cfg["database"]["database"] = self.sett_db_database.value.strip()
        self.cfg["database"]["username"] = self.sett_db_user.value.strip()
        self.cfg["database"]["password"] = self.sett_db_pass.value.strip()
        self.cfg["database"]["driver"] = self.sett_db_driver.value.strip()

        self.cfg["fileserver"]["ip"] = self.sett_fs_ip.value.strip()
        self.cfg["fileserver"]["share_suffix"] = self.sett_fs_suffix.value.strip()
        self.cfg["fileserver"]["share_map"] = self.sett_fs_map.value.strip()
        self.cfg["fileserver"]["username"] = self.sett_fs_user.value.strip()
        self.cfg["fileserver"]["password"] = self.sett_fs_pass.value.strip()

        self.cfg["paths"]["folder_map_files"] = self.sett_paths_csv.value.strip()

        config_module.save_config(self.cfg)
        self._load_folder_map(silent=True)
        self._show_toast("Đã lưu toàn bộ cấu hình vào config.ini thành công!", is_success=True)

    def _reset_default_settings(self):
        defaults = config_module.DEFAULTS
        self.sett_db_server.value = defaults["database"]["server"]
        self.sett_db_port.value = defaults["database"]["port"]
        self.sett_db_database.value = defaults["database"]["database"]
        self.sett_db_user.value = defaults["database"]["username"]
        self.sett_db_pass.value = defaults["database"]["password"]
        self.sett_db_driver.value = defaults["database"]["driver"]

        self.sett_fs_ip.value = defaults["fileserver"]["ip"]
        self.sett_fs_suffix.value = defaults["fileserver"]["share_suffix"]
        self.sett_fs_map.value = defaults["fileserver"]["share_map"]
        self.sett_fs_user.value = defaults["fileserver"]["username"]
        self.sett_fs_pass.value = defaults["fileserver"]["password"]

        self.sett_paths_csv.value = defaults["paths"]["folder_map_files"]
        self.settings_view.update()
        self._show_toast("Đã khôi phục các giá trị mặc định.", is_success=True)

    def _reload_folder_map_ui(self):
        # Đồng bộ đường dẫn từ textfield
        self.cfg["paths"]["folder_map_files"] = self.sett_paths_csv.value.strip()
        ok = self._load_folder_map(silent=False)
        if ok:
            self.csv_status_txt.value = f"Hiện đã nạp: {len(self.folder_map)} mã xã."
            self.csv_status_txt.color = CyberTheme.GREEN
            self.csv_status_txt.update()
            self._show_toast(f"Đã nạp lại thành công {len(self.folder_map)} mã xã từ CSV.", is_success=True)
            self._append_log(f"Đã nạp lại {len(self.folder_map)} mã xã từ các file CSV.", level="info")

    def _test_sql_connection(self):
        self.db_test_result.value = "Đang kiểm tra kết nối SQL Server..."
        self.db_test_result.color = CyberTheme.CYAN
        self.db_test_result.update()

        asyncio.create_task(self._async_test_sql())

    async def _async_test_sql(self):
        # Tạo bản sao cấu hình hiện tại từ các input
        temp_cfg = {
            "database": {
                "server": self.sett_db_server.value.strip(),
                "port": self.sett_db_port.value.strip(),
                "database": self.sett_db_database.value.strip(),
                "username": self.sett_db_user.value.strip(),
                "password": self.sett_db_pass.value.strip(),
                "driver": self.sett_db_driver.value.strip(),
            }
        }

        def worker():
            try:
                conn = db.connect(temp_cfg, retries=1, retry_delay=1)
                cur = conn.cursor()
                cur.execute("SELECT @@VERSION")
                ver = cur.fetchone()[0]
                conn.close()
                return True, f"Kết nối SQL Server thành công! ({ver[:45]}...)"
            except Exception as e:
                return False, f"Lỗi kết nối SQL Server: {e}"

        ok, msg = await asyncio.to_thread(worker)
        self.db_test_result.value = msg
        self.db_test_result.color = CyberTheme.GREEN if ok else CyberTheme.RED
        self.db_test_result.update()

    def _test_fs_connection(self):
        self.fs_test_result.value = "Đang kiểm tra kết nối máy chủ file..."
        self.fs_test_result.color = CyberTheme.PURPLE
        self.fs_test_result.update()

        asyncio.create_task(self._async_test_fs())

    async def _async_test_fs(self):
        ip = self.sett_fs_ip.value.strip()
        user = self.sett_fs_user.value.strip()
        password = self.sett_fs_pass.value
        share_map_str = self.sett_fs_map.value.strip()
        share_suffix = self.sett_fs_suffix.value.strip() or "$"

        def worker():
            if not ip:
                return False, "Chưa nhập IP máy chủ!"

            share_map = file_locator.parse_share_map(share_map_str)
            target_share = list(share_map.values())[0] if share_map else f"C{share_suffix}"

            if user:
                ok, msg = net_use.connect(ip, target_share, user, password)
                if ok:
                    net_use.disconnect(ip, target_share)
                    return True, f"Đăng nhập thành công share \\\\{ip}\\{target_share} với user '{user}'!"
                else:
                    return False, f"Lỗi đăng nhập share \\\\{ip}\\{target_share}: {msg}"
            else:
                # Kiểm tra cổng SMB 445
                import socket

                try:
                    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    s.settimeout(3.0)
                    ret = s.connect_ex((ip, 445))
                    s.close()
                    if ret == 0:
                        return True, f"Máy chủ {ip} đang mở cổng SMB (445) và phản hồi tốt."
                    else:
                        return False, f"Máy chủ {ip} không phản hồi cổng SMB 445 (Mã lỗi: {ret})."
                except Exception as e:
                    return False, f"Lỗi mạng tới {ip}: {e}"

        ok, msg = await asyncio.to_thread(worker)
        self.fs_test_result.value = msg
        self.fs_test_result.color = CyberTheme.GREEN if ok else CyberTheme.RED
        self.fs_test_result.update()


# ==============================================================================
# Entry Point
# ==============================================================================
async def flet_main(page: ft.Page):
    app = TaiHoSoQuetApp(page)
    await app.init_ui()


def main():
    """Hàm khởi chạy ứng dụng Flet (tương thích main.py)."""
    ft.run(flet_main)


if __name__ == "__main__":
    main()
