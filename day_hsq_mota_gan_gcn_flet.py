# -*- coding: utf-8 -*-
"""Giao diện Flet: đẩy HSQ theo mô tả (thư mục mã đơn) — file đuôi '-GCN' gán cho GCN có số trong tên file.

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
    WINDOW_TITLE = "MPLIS • Đẩy HSQ theo mô tả + file -GCN theo số GCN"

    def _build_body(self) -> ft.Control:
        self.out_dir = self._input("Thư mục gốc (chứa các thư mục <mã đơn>)", expand=True)
        self.out_dir.value = str(APP_DIR / "hsq_tai_ve")
        self.giay_to = ft.Checkbox(
            label="HSQ có đẩy GCN: chuyển mọi thành phần khác của HSQ đó sang loại Giấy tờ", value=True,
            label_style=ft.TextStyle(color=CyberTheme.AMBER, size=12))
        self.bo_qua_da_xong = ft.Checkbox(
            label=f"Bỏ qua đơn đã đẩy xong (nhớ trong {gg.FILE_TIEN_DO} ở thư mục gốc)", value=True,
            label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=12))
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
            self.giay_to,
            self.bo_qua_da_xong,
            self._note(
                "• Mỗi đơn 1 thư mục con, tên bắt đầu bằng mã đơn: 13456477 hoặc 13456477_Nguyen Van A.\n"
                "• File thường: tên file = mô tả HSQ ('Đơn đăng ký.pdf', '03_Trích lục.pdf', được '_signed', tên "
                "chứa nguyên mô tả dài cũng được) → thay file thành phần đó, giữ nguyên loại HSQ.\n"
                "• File có '-GCN' cuối tên → chỉ gán cho GCN hiện hành có SỐ PHÁT HÀNH nằm trong tên file "
                "(vd '24221_GCN_DH 770214-GCN.pdf', 'PC_A-GCN DH 770214.pdf'; tên chứa 2 số → gán cả 2 GCN). "
                "Tên file không có số GCN nào của đơn → KHÔNG đẩy, báo Excel.\n"
                "• Mỗi GCN: thay thành phần đã gắn GCN → thay thành phần có mô tả trùng tên file (bỏ '-GCN') → "
                "thêm thành phần 'Giấy chứng nhận <số>' → đơn chưa có HSQ thì tạo HSQ mới; ghi nhận loại "
                "Giấy chứng nhận.\n"
                "• HSQ không có thành phần loại GCN / Đơn đăng ký / Thông báo xác nhận → không gửi, báo Excel.\n"
                "• Tích 'chuyển sang Giấy tờ': HSQ nào có đẩy GCN thì mọi thành phần khác (kể cả Đơn đăng ký, "
                "GCN cũ) → Giấy tờ, không đổi file; thành phần GCN gắn GCN hiện hành chưa có file lần này vẫn giữ. "
                "HSQ không đẩy GCN → không đụng.\n"
                "• Sau khi gửi kiểm tra lại; MPLIS bị giảm thành phần → dừng toàn bộ; 3 đơn liên tiếp không đọc "
                "được (hết phiên) → dừng.\n"
                "• Excel kết quả ghi vào thư mục gốc.\n"
                f"• Đơn đẩy xong không lỗi được nhớ trong {gg.FILE_TIEN_DO}: chạy lại thì bỏ qua, trừ khi file "
                "trong thư mục đơn đã đổi. Xoá file đó (hoặc bỏ tích) để chạy lại từ đầu."
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
            self._run_thread(self._worker, nguon, goc, chay_thu, bool(self.giay_to.value),
                             bool(self.bo_qua_da_xong.value))

        if chay_thu:
            bat_dau()
            return
        theo_don = dm.quet_thu_muc_goc(goc)[0]
        so_don = len(theo_don)
        if self.bo_qua_da_xong.value:
            tien_do = gg.doc_tien_do(goc)
            so_don -= sum(gg.don_da_xong(tien_do, goc, m, p) for m, p in theo_don.items())
        self._hoi_xac_nhan(
            "Xác nhận đẩy HSQ",
            f"Đẩy HSQ cho {so_don} đơn{' chưa xong' if self.bo_qua_da_xong.value else ''} trong:\n{goc}\n\n"
            + ("HSQ có đẩy GCN: mọi thành phần khác sẽ chuyển sang Giấy tờ.\n\n" if self.giay_to.value else "")
            + "Thao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra' trước để xem file nào thay thành phần nào. Tiếp tục?",
            "Đẩy lên",
            bat_dau,
        )

    def _worker(self, nguon: dict[str, Any], goc: str, chay_thu: bool, giay_to: bool = False,
                bo_qua_da_xong: bool = True) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = gg.chay(
                client, goc, chay_thu,
                nen_dung=lambda: self.stop_flag,
                giay_to=giay_to,
                bao_tien_do=lambda done, total, label: self._emit("progress", done, total, label),
                bo_qua_da_xong=bo_qua_da_xong,
            )
            if excel:
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
