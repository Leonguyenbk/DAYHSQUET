# -*- coding: utf-8 -*-
"""Giao diện Flet: tải HSQ của đơn → AI (Gemini) tìm trang GCN → tách file '<số GCN>-GCN.pdf'.

Lõi xử lý: tach_gcn_theo_don.py. Khung đăng nhập, nhật ký, Excel kết quả dùng lại từ
day_hsq_theo_mota_flet.DayHsqTheoMoTaApp.
"""

from __future__ import annotations

import os
from typing import Any

import flet as ft

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import tach_gcn_theo_don as tg
import tai_hsq_don as thd
from day_hsq_theo_mota_flet import DayHsqTheoMoTaApp
from tai_hsq_don_flet import APP_DIR, CyberTheme


class TachGcnTheoDonApp(DayHsqTheoMoTaApp):
    WINDOW_TITLE = "MPLIS • Tách GCN theo đơn bằng AI"

    def _build_body(self) -> ft.Control:
        self.ids = self._input("tinhHinhDangKyId", "Mỗi dòng 1 ID (hoặc cách nhau dấu phẩy / khoảng trắng)",
                               multiline=True, min_lines=5, max_lines=8)
        self.out_dir = self._input("Thư mục kết quả", expand=True)
        self.out_dir.value = str(APP_DIR / "tach_gcn_theo_don")
        self.api_key = self._input("Gemini API key", "Mặc định lấy biến môi trường GEMINI_API_KEY",
                                   expand=True, password=True)
        self.api_key.value = os.environ.get("GEMINI_API_KEY", "")
        self.dung_file_da_tai = ft.Checkbox(
            label="Đơn đã tải HSQ (HSQ_goc/<mã đơn>) thì không tải lại", value=False,
            label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=12))
        self.quet_het = ft.Checkbox(
            label="Quét hết mọi file HSQ (mặc định: đủ GCN đúng chủ thì dừng)", value=False,
            label_style=ft.TextStyle(color=CyberTheme.TEXT_LIGHT, size=12))
        self.chay_thu = ft.Checkbox(
            label="Chạy thử phần cập nhật loại HSQ (không ghi MPLIS)", value=True,
            label_style=ft.TextStyle(color=CyberTheme.AMBER, size=12))
        self.excel_text = ft.Text("", color=CyberTheme.GREEN, size=10, selectable=True)
        self.btn_chay = self._button("Tách GCN", ft.Icons.AUTO_AWESOME_ROUNDED, CyberTheme.GREEN,
                                     self._start_tach, filled=True)
        return ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, controls=[
            ft.Container(height=6),
            ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                ft.Text("Tách GCN theo đơn (AI)", color=CyberTheme.TEXT_WHITE, size=13,
                        weight=ft.FontWeight.BOLD),
                self._button("Nạp Excel", ft.Icons.UPLOAD_FILE_ROUNDED, CyberTheme.PURPLE, self._nap_excel_ids),
            ]),
            self.ids,
            ft.Row(spacing=4, controls=[
                self.out_dir, self._icon_btn(ft.Icons.FOLDER_OPEN_ROUNDED, "Chọn thư mục", self._chon_out_dir),
            ]),
            self.api_key,
            self.dung_file_da_tai,
            self.quet_het,
            self.chay_thu,
            self._note(
                "• Mỗi đơn: đọc GCN hiện hành trên MPLIS (số phát hành, số vào sổ, chủ sử dụng). File HSQ "
                "có chữ 'GCN' trong tên xử lý trước nhất, rồi file loại GCN / 'giấy chứng nhận', 'sổ đỏ'...; "
                "đơn đăng ký / trích lục / biên bản... để cuối; tải từng "
                "file về HSQ_goc/<mã đơn>, đủ GCN đúng chủ là dừng (file còn lại không tải, không gọi AI).\n"
                "• Mỗi file: lọc trang GCN bằng ảnh nhỏ (rẻ) → chỉ trang GCN gửi ảnh nét kèm danh sách GCN của "
                "đơn để AI chỉ ra trang nào thuộc GCN nào.\n"
                "• Mỗi GCN → GCN/<mã đơn>/<SỐ GCN>-GCN.pdf. Đơn nhiều GCN → mỗi GCN 1 file.\n"
                "• Cùng 1 GCN có nhiều bản → AI đọc chủ hiện tại từng bản (trang 1; đã đính chính / biến động "
                "thì trang 4) và chọn bản trùng chủ sử dụng trong đơn. Không bản nào đúng chủ → "
                "GCN_can_kiem_tra/<mã đơn>/ (không đưa vào thư mục GCN).\n"
                "• Trang GCN không thuộc đơn (GCN cũ, thửa khác) → GCN_khong_khop/<mã đơn>/ để kiểm tra tay.\n"
                "• Đơn CHƯA CÓ GCN: không tải, không gọi AI; chỉ ghi nhận thành phần HSQ có chuỗi 'DON' trong tên "
                "(có dấu hay không, viết dính cũng lấy: 'Đơn đăng ký', 'DONDANGKY') là loại Đơn đăng ký + hồ sơ "
                "gốc (gửi đủ thành phần, không đổi file; kiểm tra lại sau khi gửi, MPLIS giảm thành phần → dừng "
                "toàn bộ). Bật 'Chạy thử' thì chỉ ghi kế hoạch vào Excel. Đơn CÓ GCN không bị cập nhật loại.\n"
                "• Trong lúc chờ AI đọc: luồng nền tải trước file kế tiếp + chi tiết và vài file đầu của 2 đơn sau "
                "(đơn chưa có GCN không tải file). Nhật ký vẫn in lần lượt từng đơn.\n"
                "• Kết quả AI lưu cache: chạy lại cùng file không tốn thêm phí.\n"
                "• Thư mục GCN dùng ngay cho tool 'Đẩy HSQ theo mô tả + file -GCN'.\n"
                "• Phần tách GCN chỉ ĐỌC MPLIS. Excel kết quả ghi vào thư mục kết quả."
            ),
            ft.Row(spacing=8, controls=[self.btn_chay]),
            self.excel_text,
        ])

    async def init_ui(self) -> None:
        await super().init_ui()
        self.log_list.controls.clear()
        self._append_log("Tách GCN theo đơn — nhập mã đơn rồi bấm 'Tách GCN'.", "title")

    def _set_running(self, running: bool) -> None:
        self.btn_chay.disabled = running
        self.btn_chay.opacity = 0.4 if running else 1
        self.btn_stop.disabled = not running
        self.btn_stop.opacity = 1 if running else 0.4

    def _start_tach(self, _=None) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        ids = thd.tach_danh_sach_id(self.ids.value or "")
        if not ids:
            self._toast("Chưa có tinhHinhDangKyId hợp lệ.", error=True)
            return
        api_key = (self.api_key.value or "").strip()
        if not api_key:
            self._toast("Nhập Gemini API key.", error=True)
            return
        out_dir = (self.out_dir.value or "").strip() or str(APP_DIR / "tach_gcn_theo_don")
        chay_thu = bool(self.chay_thu.value)

        def bat_dau():
            self._run_thread(self._worker, nguon, ids, out_dir, api_key, bool(self.dung_file_da_tai.value),
                             bool(self.quet_het.value), chay_thu)

        if chay_thu:
            bat_dau()
            return
        self._hoi_xac_nhan(
            "Xác nhận ghi MPLIS",
            f"{len(ids)} đơn. Đơn nào CHƯA CÓ GCN: thành phần HSQ có chuỗi 'DON' trong tên sẽ được GHI THẬT lên "
            "MPLIS thành loại Đơn đăng ký + hồ sơ gốc (đã tắt Chạy thử).\n\nNên chạy thử trước để xem Excel. "
            "Tiếp tục?",
            "Ghi MPLIS",
            bat_dau,
        )

    def _worker(self, nguon: dict[str, Any], ids: list[int], out_dir: str, api_key: str,  # type: ignore[override]
                dung_file_da_tai: bool, quet_het: bool, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = tg.chay(
                client, ids, out_dir, api_key,
                dung_file_da_tai=dung_file_da_tai,
                quet_het=quet_het,
                chay_thu=chay_thu,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self._emit("progress", done, total, label),
            )
            self._emit("excel_kq", excel)
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except day.DungKhanCap:
            ket_thuc = "ĐÃ DỪNG KHẨN CẤP"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self._emit("done", ket_thuc)


async def flet_main(page: ft.Page) -> None:
    app = TachGcnTheoDonApp(page)
    await app.init_ui()


def main() -> None:
    ft.run(flet_main)


if __name__ == "__main__":
    main()
