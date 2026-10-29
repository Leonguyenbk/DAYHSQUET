# -*- coding: utf-8 -*-
"""Giao diện PySide6: đẩy file HSQ theo mô tả, mỗi đơn 1 thư mục con đặt tên theo mã đơn.

<thư mục gốc>/<madon>/<mô tả>.pdf → thay đúng thành phần HSQ có mô tả trùng tên file (day_hsq_theo_mota.py).
Khung đăng nhập (Chrome / token + cookie), giải SliderCaptcha trực tiếp và nhật ký dùng lại từ
tai_hsq_don_pyside6.TaiHsqPySide6Window.
"""

from __future__ import annotations

import os
import sys
from typing import Any

from PySide6.QtCore import Qt, Slot
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

import cap_nhat_don_hsq as core
import day_hsq_ky_so as day
import day_hsq_theo_mota as dm
import tai_hsq_don as thd
from tai_hsq_don_pyside6 import APP_DIR, TaiHsqPySide6Window


class DayHsqTheoMoTaPySide6Window(TaiHsqPySide6Window):
    WINDOW_TITLE = "MPLIS • Đẩy HSQ theo mô tả (thư mục mã đơn) — PySide6"

    def __init__(self) -> None:
        self.excel_kq = ""
        super().__init__()
        self.bridge.excelKqSignal.connect(self._on_excel_kq)
        self.log_edit.clear()
        self._append_log("Đẩy HSQ theo mô tả — bấm 'Kiểm tra (chạy thử)' trước khi đẩy thật.")

    def _build_body(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(10)

        title = QLabel("📁 Đẩy HSQ theo mô tả")
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        row_dir = QHBoxLayout()
        self.inp_out_dir = QLineEdit(str(APP_DIR / "hsq_tai_ve"))
        self.inp_out_dir.setPlaceholderText("Thư mục gốc (chứa các thư mục <mã đơn>)")
        btn_dir = QPushButton("📂 Chọn thư mục")
        btn_dir.clicked.connect(self._chon_out_dir)
        row_dir.addWidget(self.inp_out_dir, 1)
        row_dir.addWidget(btn_dir)
        layout.addLayout(row_dir)

        note = QLabel(
            "• Mỗi đơn 1 thư mục con, tên bắt đầu bằng mã đơn: 13456477 hoặc 13456477_Nguyen Van A.\n"
            "• Tên file = mô tả HSQ trên MPLIS: 'Đơn đăng ký.pdf', 'Giấy chứng nhận CP 099228.pdf'. "
            "Không phân biệt dấu / hoa thường / ký tự đặc biệt; được thêm '_signed'; file tải bằng "
            "TaiDayHSQ ('03_Đơn đăng ký.pdf') cũng được.\n"
            "• Mô tả trùng nhau trong 1 đơn → cần tiền tố số thứ tự 'NN_' (như file TaiDayHSQ tải về).\n"
            "• Chỉ THAY file thành phần đã có; giữ nguyên mô tả, loại HSQ, GCN và các thành phần khác.\n"
            "• File có mô tả chứa 'GCN' + số phát hành GCN hiện hành của đơn (vd 24221_3_96_GCN_DI 923410) "
            "→ ghi nhận loại Giấy chứng nhận + gắn GCN đó.\n"
            "• Thêm '-GCN' cuối tên file → thay file + ghi nhận Giấy chứng nhận, gắn GCN có SỐ nằm trong "
            "tên file / mô tả.\n"
            "• Không tạo HSQ mới: đơn chưa có HSQ / file không khớp mô tả nào → báo trong Excel.\n"
            "• Sau khi gửi kiểm tra lại; MPLIS bị giảm thành phần → dừng toàn bộ.\n"
            "• Excel kết quả ghi vào thư mục gốc."
        )
        note.setObjectName("mutedText")
        note.setWordWrap(True)
        layout.addWidget(note)

        row_btns = QHBoxLayout()
        self.btn_quet = QPushButton("📦 Quét thư mục")
        self.btn_quet.setObjectName("btnPurple")
        self.btn_quet.clicked.connect(self._quet)

        self.btn_kiem_tra = QPushButton("🔍 Kiểm tra (chạy thử)")
        self.btn_kiem_tra.setObjectName("btnBlue")
        self.btn_kiem_tra.clicked.connect(lambda: self._start(chay_thu=True))

        self.btn_day = QPushButton("☁ Đẩy lên MPLIS")
        self.btn_day.setObjectName("btnSuccess")
        self.btn_day.clicked.connect(lambda: self._start(chay_thu=False))

        row_btns.addWidget(self.btn_quet)
        row_btns.addWidget(self.btn_kiem_tra)
        row_btns.addWidget(self.btn_day)
        row_btns.addStretch()
        layout.addLayout(row_btns)

        self.lbl_excel_kq = QLabel("")
        self.lbl_excel_kq.setObjectName("monoGreen")
        self.lbl_excel_kq.setWordWrap(True)
        self.lbl_excel_kq.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.lbl_excel_kq)
        layout.addStretch()

        return panel

    @Slot(str)
    def _on_excel_kq(self, path: str) -> None:
        self.excel_kq = path
        self.lbl_excel_kq.setText(f"Excel kết quả: {path}")

    def _set_running(self, running: bool) -> None:
        for btn in (self.btn_quet, self.btn_kiem_tra, self.btn_day):
            btn.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _mo_excel(self) -> None:
        if not self.excel_kq or not os.path.isfile(self.excel_kq):
            QMessageBox.warning(self, "Chưa có Excel", "Chưa có Excel kết quả.")
            return
        os.startfile(self.excel_kq)

    def _goc(self) -> str | None:
        goc = self.inp_out_dir.text().strip()
        if not os.path.isdir(goc):
            QMessageBox.warning(self, "Thư mục không hợp lệ", "Hãy chọn thư mục gốc hợp lệ.")
            return None
        return goc

    def _quet(self) -> None:
        """Chỉ đọc thư mục (không cần đăng nhập): liệt kê đơn + file sẽ xử lý."""
        goc = self._goc()
        if not goc:
            return
        theo_don, khong_nhan = dm.quet_thu_muc_goc(goc)
        self._append_log(f"Thư mục gốc: {goc}")
        for madon, pdfs in theo_don.items():
            self._append_log(
                f"Đơn {madon}: {len(pdfs)} file — "
                + ", ".join(os.path.basename(p) for p in pdfs[:6])
                + (" ..." if len(pdfs) > 6 else "")
            )
        if khong_nhan:
            self._append_log(
                f"⚠ Bỏ qua thư mục có PDF nhưng không nằm trong thư mục mã đơn nào ({len(khong_nhan)}): "
                f"{', '.join(khong_nhan[:10])}{' ...' if len(khong_nhan) > 10 else ''}"
            )
        self._append_log(f"Tổng: {len(theo_don)} đơn | {sum(len(v) for v in theo_don.values())} file PDF")

    def _start(self, chay_thu: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        goc = self._goc()
        if not goc:
            return

        if not chay_thu:
            so_don = len(dm.quet_thu_muc_goc(goc)[0])
            ans = QMessageBox.question(
                self,
                "Xác nhận đẩy HSQ",
                f"Thay file HSQ theo mô tả cho {so_don} đơn trong:\n{goc}\n\n"
                "Thao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra (chạy thử)' trước để xem file nào thay vào đâu. Tiếp tục?",
            )
            if ans != QMessageBox.StandardButton.Yes:
                return
        self._run_thread(self._worker, nguon, goc, chay_thu)

    def _worker(self, nguon: dict[str, Any], goc: str, chay_thu: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = dm.chay(
                client,
                goc,
                chay_thu,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self.bridge.progressSignal.emit(done, total, label),
            )
            if excel:
                self.bridge.excelKqSignal.emit(excel)
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except day.DungKhanCap as exc:
            thd.log(f"❌ {exc}")
            ket_thuc = "ĐÃ DỪNG KHẨN CẤP"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self.bridge.doneSignal.emit(ket_thuc)


def main() -> int:
    app = QApplication(sys.argv)
    win = DayHsqTheoMoTaPySide6Window()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
