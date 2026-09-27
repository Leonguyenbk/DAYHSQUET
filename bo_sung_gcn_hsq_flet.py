# -*- coding: utf-8 -*-
"""Giao diện Flet: bổ sung thành phần HSQ cho GCN còn thiếu (dùng chung 1 file GCN).

Đơn có nhiều GCN nhưng HSQ chỉ có 1 file GCN (chứa cả các GCN) → tải file đó về và thêm
thành phần HSQ cho từng GCN còn thiếu (bo_sung_gcn_hsq.py).
Khung đăng nhập (Chrome / token + cookie) và nhật ký dùng lại từ tai_hsq_don_flet.TaiHsqApp.
"""

from __future__ import annotations

import os
from typing import Any

import flet as ft

import bo_sung_gcn_hsq as bs
import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import tai_hsq_don as thd
from tai_hsq_don_flet import APP_DIR, CyberTheme, TaiHsqApp


class BoSungGcnApp(TaiHsqApp):
    WINDOW_TITLE = "MPLIS • Bổ sung HSQ cho GCN còn thiếu"

    def __init__(self, page: ft.Page):
        super().__init__(page)
        self.excel_kq = ""

    def _build_body(self) -> ft.Control:
        self.ids = self._input("tinhHinhDangKyId", "Mỗi dòng 1 ID (hoặc cách nhau dấu phẩy / khoảng trắng)",
                               multiline=True, min_lines=8, max_lines=12)
        self.out_dir = self._input("Thư mục lưu file GCN tải về", expand=True)
        self.out_dir.value = str(APP_DIR / "hsq_bo_sung_gcn")
        self.excel_text = ft.Text("", color=CyberTheme.GREEN, size=10, selectable=True)
        self.btn_kiem_tra = self._button("Kiểm tra (chạy thử)", ft.Icons.SEARCH_ROUNDED, CyberTheme.BLUE,
                                         lambda e: self._start(chay_thu=True))
        self.btn_bo_sung = self._button("Bổ sung lên MPLIS", ft.Icons.CLOUD_UPLOAD_ROUNDED, CyberTheme.GREEN,
                                        lambda e: self._start(chay_thu=False), filled=True)
        return ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=6),
            ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                ft.Text("Bổ sung HSQ cho GCN", color=CyberTheme.TEXT_WHITE, size=13, weight=ft.FontWeight.BOLD),
                self._button("Nạp Excel", ft.Icons.UPLOAD_FILE_ROUNDED, CyberTheme.PURPLE, self._nap_excel_ids),
            ]),
            self.ids,
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            self._note(
                "• Đơn có ≥ 2 GCN mà HSQ chỉ có đúng 1 file GCN → tải file đó, thêm thành phần "
                "'Giấy chứng nhận <số phát hành>' cho từng GCN chưa có file (cùng file).\n"
                "• Mỗi HSQ gửi 1 lần kèm ĐỦ thành phần cũ (giữ nguyên) + thành phần mới.\n"
                "• Bỏ qua đơn có từ 2 file GCN trở lên (không rõ file nào là nguồn).\n"
                "• Sau khi gửi: kiểm tra số thành phần; nếu MPLIS bị giảm → dừng toàn bộ.\n"
                "• 'Kiểm tra' chỉ tải file + in kế hoạch, KHÔNG gửi. Excel kết quả ghi vào thư mục lưu."
            ),
            ft.Row(spacing=8, wrap=True, controls=[self.btn_kiem_tra, self.btn_bo_sung]),
            self.excel_text,
        ])

    async def init_ui(self) -> None:
        await super().init_ui()
        self.log_list.controls.clear()
        self._append_log("Bổ sung HSQ cho GCN còn thiếu — bấm 'Kiểm tra' trước khi bổ sung thật.", "title")

    # ------------------------------------------------------------------ overrides
    def _set_running(self, running: bool) -> None:
        for c in (self.btn_kiem_tra, self.btn_bo_sung):
            c.disabled = running
            c.opacity = 0.4 if running else 1
        self.btn_stop.disabled = not running
        self.btn_stop.opacity = 1 if running else 0.4

    def _mo_excel(self, _=None) -> None:
        if not self.excel_kq or not os.path.isfile(self.excel_kq):
            self._toast("Chưa có Excel kết quả.", error=True)
            return
        os.startfile(self.excel_kq)

    async def _consume(self) -> None:
        # Bắt thêm sự kiện "excel_kq" rồi chuyển phần còn lại cho lớp cha.
        while True:
            kind, *args = await self.queue.get()
            if kind == "excel_kq":
                self.excel_kq = args[0]
                self.excel_text.value = f"Excel kết quả: {args[0]}"
                self.excel_text.update()
                continue
            await self._xu_ly_su_kien(kind, args)

    async def _xu_ly_su_kien(self, kind: str, args: list) -> None:
        if kind == "log":
            self._append_log(args[0])
        elif kind == "progress":
            done, total, label = args
            self.progress.value = done / total if total else 0
            self.status.value = label
            self.page.update()
        elif kind == "login":
            text, color, chrome_mo = args
            self.login_status.value = text
            self.login_status.color = color
            self.btn_chrome.disabled = chrome_mo
            self.btn_chrome.opacity = 0.4 if chrome_mo else 1
            self.page.update()
        elif kind == "done":
            self.is_running = False
            self._set_running(False)
            self.status.value = args[0]
            self.page.update()

    # ------------------------------------------------------------------ chạy
    def _start(self, chay_thu: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        ids = thd.tach_danh_sach_id(self.ids.value or "")
        if not ids:
            self._toast("Chưa có tinhHinhDangKyId hợp lệ.", error=True)
            return
        out_dir = (self.out_dir.value or "").strip() or str(APP_DIR / "hsq_bo_sung_gcn")

        def bat_dau():
            self._run_thread(self._worker, nguon, ids, out_dir, chay_thu)

        if chay_thu:
            bat_dau()
            return
        self._hoi_xac_nhan(
            "Xác nhận bổ sung HSQ",
            f"Thêm thành phần HSQ cho GCN còn thiếu của {len(ids)} đơn (GHI THẬT lên MPLIS).\n"
            "Nên bấm 'Kiểm tra' trước để xem kế hoạch. Tiếp tục?",
            "Bổ sung",
            bat_dau,
        )

    def _worker(self, nguon: dict[str, Any], ids: list[int], out_dir: str, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = bs.chay(
                client, ids, out_dir, chay_thu,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self._emit("progress", done, total, label),
            )
            self._emit("excel_kq", excel)
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except day.DungKhanCap as exc:
            thd.log(f"❌ {exc}")
            ket_thuc = "ĐÃ DỪNG KHẨN CẤP"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self._emit("done", ket_thuc)


async def flet_main(page: ft.Page) -> None:
    app = BoSungGcnApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
