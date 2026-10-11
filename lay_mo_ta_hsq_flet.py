# -*- coding: utf-8 -*-
"""Giao diện Flet: lấy mô tả tất cả hồ sơ quét của các đơn đăng ký → Excel.

Gọi POST /dc/DangKyAjax/GetThongTinDangKyByTinhHinhDangKyIds
     payload {"getHoSoQuet": true, "tinhHinhDangKyIds": [<id>]} cho từng đơn,
đọc moTa trong ListHoSoQuet[].ListFileHoSoQuet.ListFileHoSoQuet[] (lay_mo_ta_hsq.py).
Chỉ đọc, không tải file, không ghi gì lên MPLIS.
Khung đăng nhập (Chrome / token + cookie) và nhật ký dùng lại từ tai_hsq_don_flet.TaiHsqApp.
"""

from __future__ import annotations

import os
from typing import Any

import flet as ft

import cap_nhat_don_hsq as core
import lay_mo_ta_hsq as lmt
import tai_hsq_don as thd
from tai_hsq_don_flet import APP_DIR, CyberTheme, TaiHsqApp


class LayMoTaHsqApp(TaiHsqApp):
    WINDOW_TITLE = "MPLIS • Lấy mô tả hồ sơ quét"

    def __init__(self, page: ft.Page):
        super().__init__(page)
        self.excel_kq = ""

    def _build_body(self) -> ft.Control:
        self.ids = self._input("tinhHinhDangKyId", "Mỗi dòng 1 ID (hoặc cách nhau dấu phẩy / khoảng trắng)",
                               multiline=True, min_lines=10, max_lines=14)
        self.out_dir = self._input("Thư mục lưu Excel kết quả", expand=True)
        self.out_dir.value = str(APP_DIR)
        self.excel_text = ft.Text("", color=CyberTheme.GREEN, size=10, selectable=True)
        self.btn_lay = self._button("Lấy mô tả", ft.Icons.DESCRIPTION_ROUNDED, CyberTheme.CYAN,
                                    self._start, filled=True)
        return ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=6),
            ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                ft.Text("Lấy mô tả HSQ", color=CyberTheme.TEXT_WHITE, size=13, weight=ft.FontWeight.BOLD),
                ft.Row(spacing=8, controls=[
                    self._button("Xóa hết", ft.Icons.DELETE_SWEEP_ROUNDED, CyberTheme.RED, self._xoa_het_ids),
                    self._button("Nạp Excel", ft.Icons.UPLOAD_FILE_ROUNDED, CyberTheme.PURPLE, self._nap_excel_ids),
                ]),
            ]),
            self.ids,
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            self._note(
                "• Mỗi đơn gọi GetThongTinDangKyByTinhHinhDangKyIds (getHoSoQuet=true), chỉ ĐỌC.\n"
                "• Excel: sheet ChiTiet (mỗi file HSQ 1 dòng: madon, tờ, thửa, hosoquetid, vị trí, mô tả, "
                "loại, GCN, ký số) + sheet TongHop (mô tả khác nhau, số file, số đơn, ds mã đơn).\n"
                "• Tự lưu Excel mỗi 20 đơn; bấm Dừng vẫn ghi phần đã lấy."
            ),
            ft.Row(spacing=8, controls=[self.btn_lay]),
            self.excel_text,
        ])

    async def init_ui(self) -> None:
        await super().init_ui()
        self.log_list.controls.clear()
        self._append_log("Lấy mô tả hồ sơ quét của các đơn — chỉ đọc, không ghi lên MPLIS.", "title")
        self._append_log(f"POST {core.URL_CHI_TIET}", "title")
        self._append_log('payload: {"getHoSoQuet": true, "tinhHinhDangKyIds": [<id>]}', "title")

    # ------------------------------------------------------------------ overrides
    def _set_running(self, running: bool) -> None:
        self.btn_lay.disabled = running
        self.btn_lay.opacity = 0.4 if running else 1
        self.btn_stop.disabled = not running
        self.btn_stop.opacity = 1 if running else 0.4

    def _xoa_het_ids(self, _=None) -> None:
        so_id = len(thd.tach_danh_sach_id(self.ids.value or ""))
        self.ids.value = ""
        self.ids.update()
        self._append_log(f"Đã xóa {so_id} ID khỏi danh sách.", "title")

    def _mo_excel(self, _=None) -> None:
        if not self.excel_kq or not os.path.isfile(self.excel_kq):
            self._toast("Chưa có Excel kết quả.", error=True)
            return
        os.startfile(self.excel_kq)

    async def _consume(self) -> None:
        while True:
            kind, *args = await self.queue.get()
            if kind == "log":
                self._append_log(args[0])
            elif kind == "excel_kq":
                self.excel_kq = args[0]
                self.excel_text.value = f"Excel kết quả: {args[0]}"
                self.excel_text.update()
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
    def _start(self, _=None) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        ids = thd.tach_danh_sach_id(self.ids.value or "")
        if not ids:
            self._toast("Chưa có tinhHinhDangKyId hợp lệ.", error=True)
            return
        out_dir = (self.out_dir.value or "").strip() or str(APP_DIR)
        self._run_thread(self._worker, nguon, ids, out_dir)

    def _worker(self, nguon: dict[str, Any], ids: list[int], out_dir: str) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            os.makedirs(out_dir, exist_ok=True)
            excel = lmt.ten_excel_mac_dinh(out_dir)
            thd.log(f"Lấy mô tả HSQ của {len(ids)} đơn...")
            lmt.chay(
                client, ids, excel,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self._emit("progress", done, total, label),
            )
            if os.path.isfile(excel):
                self._emit("excel_kq", excel)
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self._emit("done", ket_thuc)


async def flet_main(page: ft.Page) -> None:
    app = LayMoTaHsqApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
