# -*- coding: utf-8 -*-
"""Giao diện PySide6: đẩy HSQ theo mô tả (thư mục mã đơn) — file đuôi '-GCN' gán cho GCN có số trong tên file.

<thư mục gốc>/<madon>/<mô tả>.pdf, <...>-GCN.pdf (day_hsq_mota_gan_gcn.py).
Khung đăng nhập, giải SliderCaptcha trực tiếp (QPainter + tự động kéo), nhật ký và quét thư mục
dùng lại từ day_hsq_theo_mota_pyside6.DayHsqTheoMoTaPySide6Window.
"""

from __future__ import annotations

import sys
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
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
import day_hsq_mota_gan_gcn as gg
import day_hsq_theo_mota as dm
import tai_hsq_don as thd
from day_hsq_theo_mota_pyside6 import DayHsqTheoMoTaPySide6Window
from tai_hsq_don_pyside6 import APP_DIR


class DayHsqMoTaGanGcnPySide6Window(DayHsqTheoMoTaPySide6Window):
    WINDOW_TITLE = "MPLIS • Đẩy HSQ theo mô tả + file -GCN theo số GCN (PySide6)"

    def __init__(self) -> None:
        super().__init__()
        self.log_edit.clear()
        self._append_log("Đẩy HSQ theo mô tả + file -GCN — bấm 'Kiểm tra (chạy thử)' trước khi đẩy thật.")

    def _build_body(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(10)

        title = QLabel("📁 Đẩy HSQ theo mô tả + file -GCN")
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

        self.chk_giay_to = QCheckBox("HSQ có đẩy GCN: chuyển mọi thành phần khác của HSQ đó sang loại Giấy tờ")
        self.chk_giay_to.setChecked(True)
        self.chk_giay_to.setStyleSheet("color: #fbbf24;")
        layout.addWidget(self.chk_giay_to)

        self.chk_bo_qua_da_xong = QCheckBox(
            f"Bỏ qua đơn đã đẩy xong (nhớ trong {gg.FILE_TIEN_DO} ở thư mục gốc)"
        )
        self.chk_bo_qua_da_xong.setChecked(True)
        layout.addWidget(self.chk_bo_qua_da_xong)

        note = QLabel(
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

    def _start(self, chay_thu: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        goc = self._goc()
        if not goc:
            return

        giay_to = self.chk_giay_to.isChecked()
        bo_qua = self.chk_bo_qua_da_xong.isChecked()

        if not chay_thu:
            theo_don = dm.quet_thu_muc_goc(goc)[0]
            so_don = len(theo_don)
            if bo_qua:
                tien_do = gg.doc_tien_do(goc)
                so_don -= sum(gg.don_da_xong(tien_do, goc, m, p) for m, p in theo_don.items())
            ans = QMessageBox.question(
                self,
                "Xác nhận đẩy HSQ",
                f"Đẩy HSQ cho {so_don} đơn{' chưa xong' if bo_qua else ''} trong:\n{goc}\n\n"
                + ("HSQ có đẩy GCN: mọi thành phần khác sẽ chuyển sang Giấy tờ.\n\n" if giay_to else "")
                + "Thao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra (chạy thử)' trước để xem file nào thay thành phần nào. Tiếp tục?",
            )
            if ans != QMessageBox.StandardButton.Yes:
                return

        self._run_thread(self._worker, nguon, goc, chay_thu, giay_to, bo_qua)

    def _worker(
        self,
        nguon: dict[str, Any],
        goc: str,
        chay_thu: bool,
        giay_to: bool = False,
        bo_qua_da_xong: bool = True,
    ) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel = gg.chay(
                client,
                goc,
                chay_thu,
                nen_dung=lambda: self.stop_flag,
                giay_to=giay_to,
                bao_tien_do=lambda done, total, label: self.bridge.progressSignal.emit(done, total, label),
                bo_qua_da_xong=bo_qua_da_xong,
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
    win = DayHsqMoTaGanGcnPySide6Window()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
