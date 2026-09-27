# -*- coding: utf-8 -*-
"""Giao diện Flet: tải hồ sơ quét theo đơn → ký số → đẩy lại thay thế đúng vị trí.

Phiên đăng nhập: đăng nhập Chrome (Selenium) hoặc dán token + cookie (F12).
Tab "Tải HSQ":  danh sách tinhHinhDangKyId (hoặc Excel cột madon)
                → tải file về <thư mục lưu>/<madon>/ + Excel kết quả (mỗi file 1 dòng).
Tab "Đẩy HSQ ký số": chọn Excel kết quả đó → đẩy file đã ký số lên, giữ nguyên metadata,
                chỉ thay file (day_hsq_ky_so.py). Trạng thái ghi ngược vào Excel.
"""

from __future__ import annotations

import asyncio
import os
import subprocess
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

import flet as ft

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import tai_hsq_don as thd
from cap_nhat_don_hsq_flet import CyberTheme, _native_choose_directory, _native_choose_excel

APP_DIR = Path(thd.THU_MUC_APP)  # cạnh file .exe khi đóng gói, cạnh file .py khi chạy thường


class TaiHsqApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.loop = asyncio.get_running_loop()
        self.queue: asyncio.Queue = asyncio.Queue()
        self.is_running = False
        self.stop_flag = False
        # Chế độ "Đăng nhập": client giữ Chrome + session lấy từ Chrome.
        self.login_client: core.MplisClient | None = None

    # ------------------------------------------------------------------ UI helpers
    def _input(self, label: str, hint: str = "", multiline: bool = False, min_lines: int = 1,
               max_lines: int = 1, expand: bool | int = False, password: bool = False) -> ft.TextField:
        return ft.TextField(
            label=label,
            hint_text=hint,
            password=password,
            can_reveal_password=password,
            multiline=multiline,
            min_lines=min_lines,
            max_lines=max_lines,
            expand=expand,
            text_size=12,
            label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=11),
            hint_style=ft.TextStyle(color=CyberTheme.TEXT_MUTED, size=11),
            color=CyberTheme.TEXT_WHITE,
            bgcolor=CyberTheme.SURFACE_INPUT,
            border={
                ft.ControlState.DEFAULT: ft.OutlineInputBorder(
                    side=ft.BorderSide(1, CyberTheme.BORDER_COLOR), border_radius=8
                ),
                ft.ControlState.FOCUSED: ft.OutlineInputBorder(
                    side=ft.BorderSide(1, CyberTheme.CYAN), border_radius=8
                ),
            },
            dense=True,
        )

    def _button(self, text: str, icon: str, color: str, callback, filled: bool = False) -> ft.Container:
        fg = CyberTheme.BG_DARK if filled else color
        return ft.Container(
            content=ft.Row(
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=7,
                controls=[
                    ft.Icon(icon, color=fg, size=16),
                    ft.Text(text, color=fg, size=12, weight=ft.FontWeight.W_600),
                ],
            ),
            height=38,
            bgcolor=color if filled else f"{color}18",
            border=ft.Border.all(1, color),
            border_radius=8,
            padding=ft.Padding.symmetric(horizontal=14, vertical=0),
            ink=True,
            on_click=callback,
        )

    def _icon_btn(self, icon: str, tooltip: str, callback, color: str = CyberTheme.CYAN) -> ft.IconButton:
        return ft.IconButton(icon, icon_color=color, icon_size=20, tooltip=tooltip, on_click=callback)

    def _card(self, title: str, icon: str, content: ft.Control, expand: bool | int = False) -> ft.Container:
        return ft.Container(
            expand=expand,
            bgcolor=CyberTheme.SURFACE_CARD,
            border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
            border_radius=10,
            padding=12,
            content=ft.Column(
                expand=expand,
                spacing=8,
                controls=[
                    ft.Row(spacing=8, controls=[
                        ft.Icon(icon, color=CyberTheme.CYAN, size=17),
                        ft.Text(title, color=CyberTheme.TEXT_WHITE, size=12, weight=ft.FontWeight.BOLD),
                    ]),
                    ft.Divider(height=1, color=CyberTheme.BORDER_COLOR),
                    content,
                ],
            ),
        )

    def _note(self, text: str) -> ft.Text:
        return ft.Text(text, color=CyberTheme.TEXT_MUTED, size=10)

    def _toast(self, message: str, error: bool = False) -> None:
        self.page.show_dialog(ft.SnackBar(
            content=ft.Text(message, color=CyberTheme.TEXT_WHITE, size=12),
            bgcolor=CyberTheme.RED if error else CyberTheme.SURFACE_HOVER,
        ))

    # ------------------------------------------------------------------ build
    # Tiêu đề cửa sổ; lớp con (tool khác) đổi lại.
    WINDOW_TITLE = "MPLIS • Tải / đẩy hồ sơ quét ký số"

    async def init_ui(self) -> None:
        p = self.page
        p.title = self.WINDOW_TITLE
        p.theme_mode = ft.ThemeMode.DARK
        p.bgcolor = CyberTheme.BG_DARK
        p.padding = 16
        p.window.width = 1260
        p.window.height = 820
        p.window.min_width = 980
        p.window.min_height = 640
        await p.window.center()

        self.btn_stop = self._button("Dừng", ft.Icons.STOP_ROUNDED, CyberTheme.RED, self._stop)
        self.btn_stop.disabled = True
        self.btn_stop.opacity = 0.4

        left = ft.Column(width=470, spacing=12, controls=[
            self._build_session_card(),
            ft.Container(
                expand=True,
                bgcolor=CyberTheme.SURFACE_CARD,
                border=ft.Border.all(1, CyberTheme.BORDER_COLOR),
                border_radius=10,
                padding=ft.Padding.symmetric(horizontal=12, vertical=4),
                content=self._build_body(),
            ),
        ])
        p.add(ft.Row(expand=True, spacing=14, vertical_alignment=ft.CrossAxisAlignment.STRETCH,
                     controls=[left, self._build_log_card()]))

        # Tắt app thì đóng luôn Chrome đăng nhập.
        p.window.prevent_close = True
        p.window.on_event = self._on_window_event

        thd.dat_ham_log(lambda msg: self._emit("log", msg))
        asyncio.create_task(self._consume())
        mau = thd.doc_mau_da_luu()
        self._append_log("Endpoint tải: " + (mau["url"] if mau else thd.MAU_MAC_DINH[0]["url"]), "title")
        p.update()

    def _build_session_card(self) -> ft.Container:
        """Khung phiên đăng nhập: đăng nhập Chrome (Selenium) hoặc dán token/cookie."""
        self.che_do = ft.RadioGroup(
            value="login",
            on_change=self._doi_che_do,
            content=ft.Row(spacing=16, controls=[
                ft.Radio(value="login", label="Đăng nhập (Chrome)",
                         label_style=ft.TextStyle(color=CyberTheme.TEXT_WHITE, size=12)),
                ft.Radio(value="cookie", label="Dán token + cookie",
                         label_style=ft.TextStyle(color=CyberTheme.TEXT_WHITE, size=12)),
            ]),
        )
        self.username = self._input("Username", expand=True)
        self.password = self._input("Password", expand=True, password=True)
        self.password.on_submit = self._mo_chrome
        self.login_status = ft.Text("Chưa đăng nhập", color=CyberTheme.TEXT_MUTED, size=11)
        self.btn_chrome = self._button("1. Mở Chrome", ft.Icons.OPEN_IN_BROWSER_ROUNDED, CyberTheme.BLUE,
                                       self._mo_chrome)
        self.btn_lay_session = self._button("2. Lấy session", ft.Icons.KEY_ROUNDED, CyberTheme.GREEN,
                                            self._lay_session_chrome)
        self.panel_login = ft.Column(spacing=10, controls=[
            ft.Row(spacing=8, controls=[self.username, self.password]),
            ft.Row(spacing=8, controls=[self.btn_chrome, self.btn_lay_session]),
            self.login_status,
        ])

        self.token = self._input("__RequestVerificationToken", "F12 > Network > Request Headers")
        self.cookie = self._input("Cookie", "Dán nguyên dòng Cookie", multiline=True, min_lines=2, max_lines=4)
        self.panel_cookie = ft.Column(spacing=10, visible=False, controls=[self.token, self.cookie])
        return self._card("Phiên đăng nhập", ft.Icons.KEY_ROUNDED,
                          ft.Column(spacing=10, controls=[self.che_do, self.panel_login, self.panel_cookie]))

    def _build_log_card(self) -> ft.Container:
        self.progress = ft.ProgressBar(value=0, color=CyberTheme.CYAN, bgcolor=CyberTheme.BORDER_COLOR)
        self.status = ft.Text("Sẵn sàng", color=CyberTheme.TEXT_LIGHT, size=11)
        self.log_list = ft.ListView(expand=True, spacing=2, auto_scroll=True)
        return self._card(
            "Nhật ký", ft.Icons.TERMINAL_ROUNDED,
            ft.Column(expand=True, spacing=8, controls=[
                ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                    self.status,
                    ft.Row(spacing=0, controls=[
                        self.btn_stop,
                        self._icon_btn(ft.Icons.TABLE_VIEW_ROUNDED, "Mở Excel kết quả", self._mo_excel,
                                       CyberTheme.TEXT_LIGHT),
                        self._icon_btn(ft.Icons.FOLDER_ROUNDED, "Mở thư mục lưu", self._mo_thu_muc,
                                       CyberTheme.TEXT_LIGHT),
                        self._icon_btn(ft.Icons.DELETE_SWEEP_ROUNDED, "Xóa nhật ký", self._clear_log,
                                       CyberTheme.TEXT_LIGHT),
                    ]),
                ]),
                self.progress,
                ft.Container(expand=True, bgcolor=CyberTheme.TERMINAL_BG, border_radius=8, padding=8,
                             content=self.log_list),
            ]),
            expand=True,
        )

    def _build_body(self) -> ft.Control:
        """Phần chức năng dưới khung đăng nhập (tool này: 2 tab Tải / Đẩy)."""
        # --- tab Tải
        self.ids = self._input("tinhHinhDangKyId", "Mỗi dòng 1 ID (hoặc cách nhau dấu phẩy / khoảng trắng)",
                               multiline=True, min_lines=6, max_lines=10)
        self.out_dir = self._input("Thư mục lưu", expand=True)
        self.out_dir.value = str(APP_DIR / "hsq_tai_ve")
        self.excel_tai = ft.Text("", color=CyberTheme.GREEN, size=10, selectable=True)
        self.btn_check = self._button("Kiểm tra", ft.Icons.SEARCH_ROUNDED, CyberTheme.BLUE,
                                      lambda e: self._start_tai(chi_kiem_tra=True))
        self.btn_tai = self._button("Tải toàn bộ", ft.Icons.DOWNLOAD_ROUNDED, CyberTheme.CYAN,
                                    lambda e: self._start_tai(chi_kiem_tra=False), filled=True)

        tab_tai = ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=2),
            ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                self._note("Nhập ID hoặc nạp từ Excel có cột madon (= tinhHinhDangKyId)."),
                self._button("Nạp Excel", ft.Icons.UPLOAD_FILE_ROUNDED, CyberTheme.PURPLE, self._nap_excel_ids),
            ]),
            self.ids,
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            ft.Row(spacing=8, controls=[self.btn_check, self.btn_tai]),
            self.excel_tai,
        ])

        # --- tab Đẩy
        self.excel_day = self._input("Excel kết quả tải (ket_qua_tai_hsq_*.xlsx)", expand=True)
        self.ky_so_dir = self._input("Thư mục file ký số (để trống = cạnh file đã tải)", expand=True)
        self.chi_da_ky = ft.Checkbox(label="Chỉ đẩy file đã có chữ ký số", value=True,
                                     label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=12))
        self.chay_thu = ft.Checkbox(label="Chạy thử (chỉ in kế hoạch, không gửi lên MPLIS)", value=True,
                                    label_style=ft.TextStyle(color=CyberTheme.AMBER, size=12))
        self.btn_day = self._button("Đẩy lên MPLIS", ft.Icons.CLOUD_UPLOAD_ROUNDED, CyberTheme.GREEN,
                                    self._xac_nhan_day, filled=True)

        tab_day = ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=2),
            ft.Row(spacing=4, controls=[
                self.excel_day, self._icon_btn(ft.Icons.FILE_OPEN_ROUNDED, "Chọn Excel", self._chon_excel_day),
            ]),
            ft.Row(spacing=4, controls=[
                self.ky_so_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_ky_so_dir),
            ]),
            self.chi_da_ky,
            self.chay_thu,
            self._note(
                "• Mỗi HSQ gửi 1 lần kèm ĐỦ mọi thành phần: thành phần không đổi giữ nguyên, "
                "thành phần có file ký số được thay đúng vị trí (thanhPhanHoSoQuetId), giữ mô tả/loại/GCN.\n"
                "• Sau khi đẩy, nếu MPLIS bị giảm số thành phần → dừng toàn bộ ngay.\n"
                "• File ký số: cột file_ky_so trong Excel, hoặc cùng tên (hay tên + _signed…) cạnh file đã tải.\n"
                "• Bỏ qua nếu file trên MPLIS đã khác lúc tải (nodeId đổi). Dòng 'Đã đẩy' không đẩy lại.\n"
                "• Trạng thái ghi ngược vào chính file Excel (đóng Excel trước khi chạy)."
            ),
            ft.Row(spacing=8, controls=[self.btn_day]),
        ])

        tabs = ft.Tabs(
            length=2,
            selected_index=0,
            expand=True,
            content=ft.Column(expand=True, spacing=0, controls=[
                ft.TabBar(
                    tabs=[
                        ft.Tab(label="Tải HSQ", icon=ft.Icons.DOWNLOAD_ROUNDED),
                        ft.Tab(label="Đẩy HSQ ký số", icon=ft.Icons.CLOUD_UPLOAD_ROUNDED),
                    ],
                    indicator_color=CyberTheme.CYAN,
                    label_color=CyberTheme.CYAN,
                    unselected_label_color=CyberTheme.TEXT_LIGHT,
                    divider_color=CyberTheme.BORDER_COLOR,
                ),
                ft.TabBarView(expand=True, controls=[tab_tai, tab_day]),
            ]),
        )
        return tabs

    # ------------------------------------------------------------------ log / queue
    def _append_log(self, message: str, level: str | None = None) -> None:
        if level is None:
            level = ("error" if "❌" in message else "success" if "✅" in message
                     else "warn" if ("⚠" in message or "⏹" in message) else "info")
        color = {
            "success": CyberTheme.GREEN,
            "error": CyberTheme.RED,
            "warn": CyberTheme.AMBER,
            "title": CyberTheme.CYAN,
        }.get(level, CyberTheme.TEXT_LIGHT)
        self.log_list.controls.append(ft.Row(spacing=8, vertical_alignment=ft.CrossAxisAlignment.START, controls=[
            ft.Text(datetime.now().strftime("%H:%M:%S"), color=CyberTheme.TEXT_MUTED, size=10),
            ft.Text(message, color=color, size=11, selectable=True, expand=True),
        ]))
        if len(self.log_list.controls) > 2000:
            del self.log_list.controls[:200]
        self.log_list.update()

    def _clear_log(self, _=None) -> None:
        self.log_list.controls.clear()
        self.log_list.update()

    def _emit(self, kind: str, *args) -> None:
        self.loop.call_soon_threadsafe(self.queue.put_nowait, (kind, *args))

    async def _consume(self) -> None:
        while True:
            kind, *args = await self.queue.get()
            if kind == "log":
                self._append_log(args[0])
            elif kind == "progress":
                done, total, label = args
                self.progress.value = done / total if total else 0
                self.status.value = label
                self.page.update()
            elif kind == "excel_tai":
                self.excel_tai.value = f"Excel kết quả: {args[0]}"
                self.excel_day.value = args[0]  # sẵn cho tab Đẩy
                self.page.update()
            elif kind == "login":
                text, color, chrome_mo = args
                self.login_status.value = text
                self.login_status.color = color
                # "Mở Chrome" chỉ bật lại khi Chrome chưa mở / mở lỗi.
                self.btn_chrome.disabled = chrome_mo
                self.btn_chrome.opacity = 0.4 if chrome_mo else 1
                self.page.update()
            elif kind == "done":
                self.is_running = False
                self._set_running(False)
                self.status.value = args[0]
                self.page.update()

    # ------------------------------------------------------------------ chọn file/thư mục
    async def _nap_excel_ids(self, _=None) -> None:
        path = await asyncio.to_thread(_native_choose_excel, str(APP_DIR))
        if not path:
            return
        try:
            ids = thd.tach_danh_sach_id(thd.doc_id_tu_excel(path))
        except Exception as exc:
            self._toast(f"Không đọc được Excel: {exc}", error=True)
            return
        cu = thd.tach_danh_sach_id(self.ids.value or "")
        gop = cu + [i for i in ids if i not in cu]
        self.ids.value = "\n".join(str(i) for i in gop)
        self.ids.update()
        self._append_log(f"Nạp {len(ids)} ID từ {os.path.basename(path)} (tổng {len(gop)}).", "title")

    async def _chon_out_dir(self, _=None) -> None:
        selected = await asyncio.to_thread(_native_choose_directory, self.out_dir.value.strip())
        if selected:
            self.out_dir.value = selected
            self.out_dir.update()

    async def _chon_ky_so_dir(self, _=None) -> None:
        selected = await asyncio.to_thread(_native_choose_directory, self.ky_so_dir.value.strip()
                                           or self.out_dir.value.strip())
        if selected:
            self.ky_so_dir.value = selected
            self.ky_so_dir.update()

    async def _chon_excel_day(self, _=None) -> None:
        path = await asyncio.to_thread(_native_choose_excel, self.out_dir.value.strip())
        if path:
            self.excel_day.value = path
            self.excel_day.update()

    def _mo_excel(self, _=None) -> None:
        path = (self.excel_day.value or "").strip()
        if not os.path.isfile(path):
            self._toast("Chưa có Excel kết quả.", error=True)
            return
        os.startfile(path)

    def _mo_thu_muc(self, _=None) -> None:
        path = (self.out_dir.value or "").strip()
        if not os.path.isdir(path):
            self._toast("Thư mục lưu chưa tồn tại.", error=True)
            return
        subprocess.Popen(["explorer", os.path.normpath(path)])

    # ------------------------------------------------------------------ chạy / dừng
    def _set_running(self, running: bool) -> None:
        for c in (self.btn_check, self.btn_tai, self.btn_day):
            c.disabled = running
            c.opacity = 0.4 if running else 1
        self.btn_stop.disabled = not running
        self.btn_stop.opacity = 1 if running else 0.4

    def _stop(self, _=None) -> None:
        if self.is_running:
            self.stop_flag = True
            self._append_log("⏹ Đã yêu cầu dừng, sẽ dừng sau file đang xử lý...", "warn")

    def _nguon_session(self) -> dict[str, Any] | None:
        """Nguồn session cho worker: client Chrome đã lấy session, hoặc token + cookie dán tay."""
        if self.che_do.value == "login":
            if self.login_client is None or self.login_client.session is None:
                self._toast("Đăng nhập Chrome rồi bấm '2. Lấy session' trước.", error=True)
                return None
            return {"client": self.login_client}
        token = (self.token.value or "").strip()
        cookie = (self.cookie.value or "").strip()
        if not token or not cookie:
            self._toast("Dán token và cookie trước.", error=True)
            return None
        return {"token": token, "cookie": cookie}

    # ------------------------------------------------------------------ đăng nhập Chrome
    def _doi_che_do(self, _=None) -> None:
        dang_nhap = self.che_do.value == "login"
        self.panel_login.visible = dang_nhap
        self.panel_cookie.visible = not dang_nhap
        self.page.update()

    def _mo_chrome(self, _=None) -> None:
        user = (self.username.value or "").strip()
        pwd = self.password.value or ""
        if not user or not pwd:
            self._toast("Nhập username và password.", error=True)
            return
        if self.login_client and self.login_client.driver:
            self._toast("Chrome đang mở — đăng nhập xong thì bấm '2. Lấy session'.")
            return
        self.btn_chrome.disabled = True
        self.btn_chrome.opacity = 0.4
        self.login_status.value = "Đang mở Chrome..."
        self.login_status.color = CyberTheme.AMBER
        self.page.update()

        def work():
            client = core.MplisClient(thd.log)
            try:
                client.open_browser_and_fill_login(user, pwd)
                self.login_client = client
                self._emit("login", "Chrome đã mở. Đăng nhập xong (OTP nếu có) thì bấm '2. Lấy session'.",
                           CyberTheme.AMBER, True)
            except Exception as exc:
                client.close_browser()
                thd.log(f"❌ Mở Chrome lỗi: {core.rut_gon_text(exc, 300)}")
                self._emit("login", "Mở Chrome lỗi", CyberTheme.RED, False)

        threading.Thread(target=work, daemon=True).start()

    def _lay_session_chrome(self, _=None) -> None:
        client = self.login_client
        if client is None or client.driver is None:
            self._toast("Bấm '1. Mở Chrome' và đăng nhập trước.", error=True)
            return

        def work():
            try:
                client.build_session_from_browser()
                self._emit("login", "✅ Đã có session. Chrome vẫn mở — hết hạn thì bấm lại '2. Lấy session'.",
                           CyberTheme.GREEN, True)
            except Exception as exc:
                thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
                self._emit("login", "Chưa lấy được session — kiểm tra đã đăng nhập xong chưa.",
                           CyberTheme.RED, True)

        threading.Thread(target=work, daemon=True).start()

    async def _on_window_event(self, e) -> None:
        if e.type == ft.WindowEventType.CLOSE:
            if self.login_client:
                await asyncio.to_thread(self.login_client.close_browser)
            await self.page.window.destroy()

    def _run_thread(self, target, *args) -> None:
        self.is_running = True
        self.stop_flag = False
        self._set_running(True)
        self.progress.value = 0
        self.page.update()
        threading.Thread(target=target, args=args, daemon=True).start()

    def _start_tai(self, chi_kiem_tra: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        ids = thd.tach_danh_sach_id(self.ids.value or "")
        if not ids:
            self._toast("Chưa có tinhHinhDangKyId hợp lệ.", error=True)
            return
        out_dir = (self.out_dir.value or "").strip() or str(APP_DIR / "hsq_tai_ve")
        self._run_thread(self._worker_tai, nguon, ids, out_dir, chi_kiem_tra)

    def _xac_nhan_day(self, _=None) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        path = (self.excel_day.value or "").strip()
        if not os.path.isfile(path):
            self._toast("Chọn Excel kết quả tải trước.", error=True)
            return

        chay_thu = bool(self.chay_thu.value)

        def bat_dau():
            self._run_thread(self._worker_day, nguon, path, (self.ky_so_dir.value or "").strip(),
                             bool(self.chi_da_ky.value), chay_thu)

        if chay_thu:
            bat_dau()
            return
        self._hoi_xac_nhan(
            "Xác nhận đẩy HSQ",
            f"Đẩy file ký số thay thế file hiện có trên MPLIS theo:\n{path}\n\n"
            + ("Chỉ đẩy file đã có chữ ký số." if self.chi_da_ky.value else "⚠ Đẩy cả file CHƯA ký số.")
            + "\nThao tác GHI THẬT lên MPLIS (đã tắt Chạy thử). Tiếp tục?",
            "Đẩy lên",
            bat_dau,
        )

    def _hoi_xac_nhan(self, title: str, noi_dung: str, nut_dong_y: str, khi_dong_y) -> None:
        def huy(_e=None):
            self.page.pop_dialog()

        def dong_y(_e=None):
            self.page.pop_dialog()
            khi_dong_y()

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text(title, size=15),
            content=ft.Text(noi_dung, size=12),
            actions=[ft.TextButton("Hủy", on_click=huy), ft.TextButton(nut_dong_y, on_click=dong_y)],
        ))

    # ------------------------------------------------------------------ workers (thread)
    def _tao_client(self, nguon: dict[str, Any]) -> core.MplisClient:
        if "client" in nguon:
            return nguon["client"]
        client = core.MplisClient(thd.log)
        client.build_session_from_manual(nguon["token"], nguon["cookie"])
        return client

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
                self._emit("progress", n - 1, len(ids), f"Đơn {id_don} ({n}/{len(ids)})")
                thd.log(f"--- [{n}/{len(ids)}] Đơn {id_don} ---")
                try:
                    kq = thd.xu_ly_don(client, downloader, id_don, out_dir, chi_kiem_tra,
                                       nen_dung=lambda: self.stop_flag)
                except Exception as exc:
                    thd.log(f"❌ Đơn {id_don}: {core.rut_gon_text(exc, 300)}")
                    kq = {"id_don": id_don, "so_file": 0, "da_tai": 0, "loi": [str(exc)],
                          "dong": [{"madon": id_don, "trang_thai_tai": "Lỗi", "ghi_chu_tai": str(exc)}]}
                tong.append(kq)
                if excel_kq:
                    try:
                        thd.ghi_excel_ket_qua(excel_kq, [d for k in tong for d in k["dong"]])
                    except PermissionError:
                        thd.log("⚠ Không ghi được Excel kết quả (file đang mở?)")
            self._emit("progress", len(tong), len(ids), ket_thuc)

            so_file = sum(k["so_file"] for k in tong)
            thd.log("=" * 50)
            if chi_kiem_tra:
                thd.log(f"✅ {len(tong)} đơn | {so_file} file HSQ")
            else:
                thd.log(f"✅ {len(tong)} đơn | {so_file} file HSQ | đã tải "
                        f"{sum(k['da_tai'] for k in tong)} → {out_dir}")
            khong_hsq = [str(k["id_don"]) for k in tong if k["so_file"] == 0 and not k["loi"]]
            if khong_hsq:
                thd.log(f"⚠ Đơn không có HSQ: {', '.join(khong_hsq)}")
            co_loi = [str(k["id_don"]) for k in tong if k["loi"]]
            if co_loi:
                thd.log(f"❌ Đơn có lỗi: {', '.join(co_loi)}")
            if excel_kq and tong:
                thd.log(f"Excel kết quả: {excel_kq}")
                self._emit("excel_tai", excel_kq)
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self._emit("done", ket_thuc)

    def _worker_day(self, nguon: dict[str, Any], excel_path: str, ky_so_dir: str, chi_da_ky: bool,
                    chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            day.day_tu_excel(
                client, excel_path, ky_so_dir, chi_da_ky,
                chay_thu=chay_thu,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self._emit("progress", done, total, label),
            )
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self._emit("done", ket_thuc)


async def flet_main(page: ft.Page) -> None:
    app = TaiHsqApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
