# -*- coding: utf-8 -*-
"""Giao diện Flet: đẩy file Đơn đăng ký theo thư mục mã đơn cho đơn chưa có GCN + cập nhật đăng ký
(có quyền quản lý, chưa đủ điều kiện cấp giấy).

<thư mục gốc>/<madon>/<tên>-DON.pdf (day_don_theo_thu_muc.py).
Khung đăng nhập, nhật ký, quét thư mục dùng lại từ day_hsq_theo_mota_flet.DayHsqTheoMoTaApp.
"""

from __future__ import annotations

from typing import Any

import flet as ft

import cap_nhat_don_hsq as core
import day_don_theo_thu_muc as dd
import day_hsq_ky_so as day
import day_hsq_theo_mota as dm
import tai_hsq_don as thd
from day_hsq_theo_mota_flet import DayHsqTheoMoTaApp
from tai_hsq_don_flet import APP_DIR, CyberTheme


class DayDonTheoThuMucApp(DayHsqTheoMoTaApp):
    WINDOW_TITLE = "MPLIS • Đẩy đơn đăng ký theo thư mục mã đơn"

    def _build_body(self) -> ft.Control:
        self.out_dir = self._input("Thư mục gốc (chứa các thư mục <mã đơn>)", expand=True)
        self.out_dir.value = str(APP_DIR / "hsq_tai_ve" / "DON")
        self.excel_text = ft.Text("", color=CyberTheme.GREEN, size=10, selectable=True)
        self.btn_quet = self._button("Quét thư mục", ft.Icons.FOLDER_ZIP_ROUNDED, CyberTheme.PURPLE, self._quet)
        self.btn_kiem_tra = self._button("Kiểm tra (chạy thử)", ft.Icons.SEARCH_ROUNDED, CyberTheme.BLUE,
                                         lambda e: self._start(chay_thu=True))
        self.btn_day = self._button("Đẩy đơn + cập nhật đăng ký", ft.Icons.CLOUD_UPLOAD_ROUNDED, CyberTheme.GREEN,
                                    lambda e: self._start(chay_thu=False), filled=True)
        return ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=6),
            ft.Text("Đẩy đơn đăng ký theo thư mục mã đơn (đơn chưa có GCN)", color=CyberTheme.TEXT_WHITE, size=13,
                    weight=ft.FontWeight.BOLD),
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            self._note(
                "• Mỗi đơn 1 thư mục con, tên bắt đầu bằng mã đơn (vd thư mục DON\\ do Tải & tách HSQ tạo).\n"
                "• Chỉ xử lý đơn CHƯA CÓ GCN hiện hành; đơn đã có GCN (kể cả GCN chưa có số) → bỏ qua.\n"
                "• File đơn: tên có hậu tố '-DON' hoặc là 'Đơn đăng ký' / 'Đơn đề nghị'. Được đuôi '_signed'.\n"
                "• Ghi nhận loại Đơn đăng ký: HSQ đã có thành phần Đơn đăng ký → không đẩy; thành phần có mô tả "
                "đúng bằng tên file → thay file; không có → thêm thành phần 'Đơn đăng ký'; chưa có HSQ → tạo mới.\n"
                "• Sau đó cập nhật đăng ký: có quyền quản lý, chưa đủ điều kiện cấp giấy (dieuKienCapGiay = 3).\n"
                "• Sau khi gửi kiểm tra lại; MPLIS bị giảm thành phần / đổi số đăng ký quyền → dừng toàn bộ.\n"
                "• Excel kết quả ghi vào thư mục gốc."
            ),
            ft.Row(spacing=8, wrap=True, controls=[self.btn_quet, self.btn_kiem_tra, self.btn_day]),
            self.excel_text,
        ])

    async def init_ui(self) -> None:
        await super().init_ui()
        self.log_list.controls.clear()
        self._append_log("Đẩy đơn đăng ký theo thư mục mã đơn — bấm 'Kiểm tra' trước khi đẩy thật.", "title")

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
            "Xác nhận đẩy đơn",
            f"Đẩy file đơn + cập nhật đăng ký (quyền quản lý, chưa đủ điều kiện cấp giấy) cho tối đa {so_don} "
            f"đơn trong:\n{goc}\n\nThao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra' trước. Tiếp tục?",
            "Đẩy lên",
            bat_dau,
        )

    def _worker(self, nguon: dict[str, Any], goc: str, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = dd.chay(
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
    app = DayDonTheoThuMucApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
