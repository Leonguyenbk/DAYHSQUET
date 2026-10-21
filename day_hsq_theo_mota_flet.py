# -*- coding: utf-8 -*-
"""Giao diện Flet: đẩy file HSQ theo mô tả, mỗi đơn 1 thư mục con đặt tên theo mã đơn.

<thư mục gốc>/<madon>/<mô tả>.pdf → thay đúng thành phần HSQ có mô tả trùng tên file (day_hsq_theo_mota.py).
Khung đăng nhập (Chrome / token + cookie) và nhật ký dùng lại từ tai_hsq_don_flet.TaiHsqApp.
"""

from __future__ import annotations

import os
from typing import Any

import flet as ft

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import day_hsq_theo_mota as dm
import tai_hsq_don as thd
from tai_hsq_don_flet import APP_DIR, CyberTheme, TaiHsqApp


class DayHsqTheoMoTaApp(TaiHsqApp):
    WINDOW_TITLE = "MPLIS • Đẩy HSQ theo mô tả (thư mục mã đơn)"

    def __init__(self, page: ft.Page):
        super().__init__(page)
        self.excel_kq = ""

    def _build_body(self) -> ft.Control:
        # Đặt tên out_dir để dùng lại _chon_out_dir của lớp cha.
        self.out_dir = self._input("Thư mục gốc (chứa các thư mục <mã đơn>)", expand=True)
        self.out_dir.value = str(APP_DIR / "hsq_tai_ve")
        self.excel_text = ft.Text("", color=CyberTheme.GREEN, size=10, selectable=True)
        self.btn_quet = self._button("Quét thư mục", ft.Icons.FOLDER_ZIP_ROUNDED, CyberTheme.PURPLE, self._quet)
        self.btn_kiem_tra = self._button("Kiểm tra (chạy thử)", ft.Icons.SEARCH_ROUNDED, CyberTheme.BLUE,
                                         lambda e: self._start(chay_thu=True))
        self.btn_day = self._button("Đẩy lên MPLIS", ft.Icons.CLOUD_UPLOAD_ROUNDED, CyberTheme.GREEN,
                                    lambda e: self._start(chay_thu=False), filled=True)
        return ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=6),
            ft.Text("Đẩy HSQ theo mô tả", color=CyberTheme.TEXT_WHITE, size=13, weight=ft.FontWeight.BOLD),
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            self._note(
                "• Mỗi đơn 1 thư mục con, tên bắt đầu bằng mã đơn: 13456477 hoặc 13456477_Nguyen Van A.\n"
                "• Tên file = mô tả HSQ trên MPLIS: 'Đơn đăng ký.pdf', 'Giấy chứng nhận CP 099228.pdf'. "
                "Không phân biệt dấu / hoa thường / ký tự đặc biệt; được thêm '_signed'; file tải bằng "
                "TaiDayHSQ ('03_Đơn đăng ký.pdf') cũng được.\n"
                "• Mô tả trùng nhau trong 1 đơn → cần tiền tố số thứ tự 'NN_' (như file TaiDayHSQ tải về).\n"
                "• Chỉ THAY file thành phần đã có; giữ nguyên mô tả, loại HSQ, GCN và các thành phần khác.\n"
                "• File có mô tả chứa 'GCN' + số phát hành GCN hiện hành của đơn (vd 24221_3_96_GCN_DI 923410) "
                "→ ghi nhận loại Giấy chứng nhận + gắn GCN đó. HSQ vẫn không có thành phần loại GCN / Đơn đăng ký"
                " / Thông báo xác nhận → không gửi (MPLIS sẽ từ chối), báo trong Excel.\n"
                "• Thêm '-GCN' cuối tên file → thay file + ghi nhận Giấy chứng nhận, gắn GCN có SỐ nằm trong "
                "tên file / mô tả (vd '24232_GCN_DH 770214-GCN.pdf' hoặc 'PC_A-GCN DH 770214.pdf'). Không có số "
                "GCN của đơn → không đẩy file đó.\n"
                "• Không tạo HSQ mới: đơn chưa có HSQ / file không khớp mô tả nào → báo trong Excel.\n"
                "• Sau khi gửi kiểm tra lại; MPLIS bị giảm thành phần → dừng toàn bộ.\n"
                "• Excel kết quả ghi vào thư mục gốc."
            ),
            ft.Row(spacing=8, wrap=True, controls=[self.btn_quet, self.btn_kiem_tra, self.btn_day]),
            self.excel_text,
        ])

    async def init_ui(self) -> None:
        await super().init_ui()
        self.log_list.controls.clear()
        self._append_log("Đẩy HSQ theo mô tả — bấm 'Kiểm tra' trước khi đẩy thật.", "title")

    # ------------------------------------------------------------------ overrides
    def _set_running(self, running: bool) -> None:
        for c in (self.btn_quet, self.btn_kiem_tra, self.btn_day):
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
    def _goc(self) -> str | None:
        goc = (self.out_dir.value or "").strip()
        if not os.path.isdir(goc):
            self._toast("Chọn thư mục gốc hợp lệ.", error=True)
            return None
        return goc

    def _quet(self, _=None) -> None:
        """Chỉ đọc thư mục (không cần đăng nhập): liệt kê đơn + file sẽ xử lý."""
        goc = self._goc()
        if not goc:
            return
        theo_don, khong_nhan = dm.quet_thu_muc_goc(goc)
        self._append_log(f"Thư mục gốc: {goc}", "title")
        for madon, pdfs in theo_don.items():
            self._append_log(f"Đơn {madon}: {len(pdfs)} file — " + ", ".join(os.path.basename(p) for p in pdfs[:6])
                             + (" ..." if len(pdfs) > 6 else ""))
        if khong_nhan:
            self._append_log(f"Bỏ qua thư mục có PDF nhưng không nằm trong thư mục mã đơn nào ({len(khong_nhan)}): {', '.join(khong_nhan[:10])}{' ...' if len(khong_nhan) > 10 else ''}", "warn")
        self._append_log(f"Tổng: {len(theo_don)} đơn | {sum(len(v) for v in theo_don.values())} file PDF", "title")

    def _start(self, chay_thu: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        goc = self._goc()
        if not goc:
            return

        def bat_dau():
            self._run_thread(self._worker, nguon, goc, chay_thu)

        if chay_thu:
            bat_dau()
            return
        so_don = len(dm.quet_thu_muc_goc(goc)[0])
        self._hoi_xac_nhan(
            "Xác nhận đẩy HSQ",
            f"Thay file HSQ theo mô tả cho {so_don} đơn trong:\n{goc}\n\n"
            "Thao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra' trước để xem file nào thay vào đâu. Tiếp tục?",
            "Đẩy lên",
            bat_dau,
        )

    def _worker(self, nguon: dict[str, Any], goc: str, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = dm.chay(
                client, goc, chay_thu,
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
    app = DayHsqTheoMoTaApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
