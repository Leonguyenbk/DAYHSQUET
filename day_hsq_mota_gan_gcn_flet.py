# -*- coding: utf-8 -*-
"""Giao diện Flet: đẩy HSQ theo mô tả (thư mục mã đơn) — file đuôi '-GCN' gán cho mọi GCN của đơn.

<thư mục gốc>/<madon>/<mô tả>.pdf, <...>-GCN.pdf (day_hsq_mota_gan_gcn.py).
Khung đăng nhập, nhật ký, quét thư mục dùng lại từ day_hsq_theo_mota_flet.DayHsqTheoMoTaApp.
"""

from __future__ import annotations

from typing import Any

import flet as ft

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import day_hsq_mota_gan_gcn as gg
import day_hsq_theo_mota as dm
import tai_hsq_don as thd
from day_hsq_theo_mota_flet import DayHsqTheoMoTaApp
from tai_hsq_don_flet import APP_DIR, CyberTheme


class DayHsqMoTaGanGcnApp(DayHsqTheoMoTaApp):
    WINDOW_TITLE = "MPLIS • Đẩy HSQ theo mô tả + file -GCN cho mọi GCN"

    def _build_body(self) -> ft.Control:
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
            ft.Text("Đẩy HSQ theo mô tả + file -GCN", color=CyberTheme.TEXT_WHITE, size=13,
                    weight=ft.FontWeight.BOLD),
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            self._note(
                "• Mỗi đơn 1 thư mục con, tên bắt đầu bằng mã đơn: 13456477 hoặc 13456477_Nguyen Van A.\n"
                "• File thường: tên file = mô tả HSQ ('Đơn đăng ký.pdf', '03_Trích lục.pdf', được '_signed', tên "
                "chứa nguyên mô tả dài cũng được) → thay file thành phần đó, giữ nguyên loại HSQ.\n"
                "• File có '-GCN' cuối tên (vd 'PC_NGUYEN VAN A-GCN.pdf') → gán cho MỌI GCN hiện hành của đơn: "
                "thay thành phần đã gắn GCN → thay thành phần có mô tả trùng tên file (bỏ '-GCN') → thêm thành "
                "phần 'Giấy chứng nhận <số>' → đơn chưa có HSQ thì tạo HSQ mới. Mỗi GCN 1 thành phần, ghi nhận "
                "loại Giấy chứng nhận.\n"
                "• '...-GCN DH 770214.pdf' → chỉ cho GCN đó; GCN còn lại dùng file '-GCN' không ghi số.\n"
                "• HSQ không có thành phần loại GCN / Đơn đăng ký / Thông báo xác nhận → không gửi, báo Excel.\n"
                "• Sau khi gửi kiểm tra lại; MPLIS bị giảm thành phần → dừng toàn bộ; 3 đơn liên tiếp không đọc "
                "được (hết phiên) → dừng.\n"
                "• Excel kết quả ghi vào thư mục gốc."
            ),
            ft.Row(spacing=8, wrap=True, controls=[self.btn_quet, self.btn_kiem_tra, self.btn_day]),
            self.excel_text,
        ])

    async def init_ui(self) -> None:
        await super().init_ui()
        self.log_list.controls.clear()
        self._append_log("Đẩy HSQ theo mô tả + file -GCN — bấm 'Kiểm tra' trước khi đẩy thật.", "title")

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
            f"Đẩy HSQ cho {so_don} đơn trong:\n{goc}\n\n"
            "Thao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra' trước để xem file nào thay thành phần nào. Tiếp tục?",
            "Đẩy lên",
            bat_dau,
        )

    def _worker(self, nguon: dict[str, Any], goc: str, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = gg.chay(
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
    app = DayHsqMoTaGanGcnApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
