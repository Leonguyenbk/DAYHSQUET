# -*- coding: utf-8 -*-
"""Giao diện PySide6: đẩy file GCN theo danh sách mã đơn trong Excel (day_gcn_theo_excel.py).

Excel cột 'madon' → GCN hiện hành của đơn → file PDF tên đúng số phát hành trong folder (kể cả folder con)
→ đẩy làm thành phần GCN, các thành phần khác của HSQ đó chuyển Giấy tờ.
Khung đăng nhập, giải SliderCaptcha, nhật ký dùng lại từ day_hsq_theo_mota_pyside6.DayHsqTheoMoTaPySide6Window.
"""

from __future__ import annotations

import os
import sys
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

import cap_nhat_don_hsq as core
import day_gcn_theo_excel as ge
import day_hsq_ky_so as day
import tai_hsq_don as thd
from day_hsq_theo_mota_pyside6 import DayHsqTheoMoTaPySide6Window
from tai_hsq_don_pyside6 import APP_DIR


class DayGcnTheoExcelPySide6Window(DayHsqTheoMoTaPySide6Window):
    WINDOW_TITLE = "MPLIS • Đẩy GCN theo Excel mã đơn — thành phần khác sang Giấy tờ (PySide6)"

    def __init__(self) -> None:
        super().__init__()
        self.log_edit.clear()
        self._append_log("Đẩy GCN theo Excel mã đơn — bấm 'Kiểm tra (chạy thử)' trước khi đẩy thật.")

    def _build_body(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(10)

        title = QLabel("📄 Đẩy GCN theo Excel mã đơn")
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        row_excel = QHBoxLayout()
        self.inp_excel_madon = QLineEdit()
        self.inp_excel_madon.setPlaceholderText("File Excel có cột madon")
        btn_excel = QPushButton("📄 Chọn Excel")
        btn_excel.clicked.connect(self._chon_excel_madon)
        row_excel.addWidget(self.inp_excel_madon, 1)
        row_excel.addWidget(btn_excel)
        layout.addLayout(row_excel)

        row_dir = QHBoxLayout()
        self.inp_out_dir = QLineEdit(str(APP_DIR / "hsq_tai_ve"))
        self.inp_out_dir.setPlaceholderText("Folder chứa file GCN (tìm cả folder con)")
        btn_dir = QPushButton("📂 Chọn folder")
        btn_dir.clicked.connect(self._chon_out_dir)
        row_dir.addWidget(self.inp_out_dir, 1)
        row_dir.addWidget(btn_dir)
        layout.addLayout(row_dir)

        self.chk_bo_qua_da_xong = QCheckBox(
            f"Bỏ qua đơn đã đẩy xong (nhớ trong {ge.FILE_TIEN_DO} cạnh file Excel mã đơn)"
        )
        self.chk_bo_qua_da_xong.setChecked(True)
        layout.addWidget(self.chk_bo_qua_da_xong)

        note = QLabel(
            "• Excel: cột 'madon' (không có header thì lấy cột A).\n"
            "• Mỗi đơn lấy các GCN hiện hành, tìm file PDF có TÊN CHỨA SỐ PHÁT HÀNH trong folder (kể cả folder "
            "con): '24221_GCN_DH 770214_signed.pdf', 'DH770214.pdf', 'GCN dh_770214.pdf' đều khớp GCN 'DH 770214'; "
            "'DH 7702145.pdf' thì không.\n"
            "• Nhiều file cùng khớp 1 GCN → chọn 1 file: tên có chữ 'GCN' > đã ký số > đường dẫn xếp trước. "
            "Tên chứa 2 số GCN của đơn → gán cho cả 2.\n"
            "• Đơn CHƯA CÓ GCN hiện hành → không đẩy file: chọn 1 thành phần có sẵn trong HSQ (ưu tiên đã là Đơn "
            "đăng ký, rồi mô tả có chữ 'Đơn', rồi 'Giấy tờ') đổi loại Đơn đăng ký, thành phần khác → Giấy tờ. "
            "Đơn chưa có HSQ → bỏ qua.\n"
            "• Mỗi GCN: thay thành phần đã gắn GCN → thay thành phần có mô tả trùng tên file → thêm thành phần "
            "'Giấy chứng nhận <số>' → đơn chưa có HSQ thì tạo HSQ mới.\n"
            "• HSQ có đẩy GCN: mọi thành phần khác (kể cả Đơn đăng ký) → Giấy tờ, không đổi file.\n"
            "• Sau khi gửi kiểm tra lại; MPLIS bị giảm thành phần → dừng toàn bộ.\n"
            "• Excel kết quả ghi cạnh file Excel mã đơn.\n"
            f"• Đơn đẩy xong không lỗi được nhớ trong {ge.FILE_TIEN_DO}: chạy lại thì bỏ qua, trừ khi file GCN "
            "đã đẩy bị sửa / xoá. Đơn lỗi / GCN thiếu file → lần sau chạy lại. Xoá file đó (hoặc bỏ tích) để "
            "chạy lại từ đầu."
        )
        note.setObjectName("mutedText")
        note.setWordWrap(True)
        layout.addWidget(note)

        row_btns = QHBoxLayout()
        self.btn_quet = QPushButton("📦 Quét Excel + folder")
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

    def _chon_excel_madon(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel có cột madon",
            os.path.dirname(self.inp_excel_madon.text().strip()) or str(APP_DIR), "Excel (*.xlsx *.xlsm)"
        )
        if path:
            self.inp_excel_madon.setText(os.path.normpath(path))

    def _excel_madon(self) -> str | None:
        path = self.inp_excel_madon.text().strip()
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Thiếu Excel", "Hãy chọn file Excel có cột madon.")
            return None
        return path

    def _quet(self) -> None:
        """Chỉ đọc Excel + folder (không cần đăng nhập): số mã đơn, số file PDF."""
        excel = self._excel_madon()
        goc = self._goc() if excel else None
        if not goc:
            return
        try:
            ids = thd.tach_danh_sach_id(thd.doc_id_tu_excel(excel))
        except Exception as exc:
            QMessageBox.critical(self, "Lỗi đọc Excel", f"Không đọc được Excel: {exc}")
            return
        ds_pdf = ge.quet_pdf(goc)
        self._append_log(f"Excel: {excel} → {len(ids)} mã đơn: "
                         + ", ".join(str(i) for i in ids[:10]) + (" ..." if len(ids) > 10 else ""))
        self._append_log(f"Folder: {goc} → {len(ds_pdf)} file PDF (kể cả folder con)")
        self._append_log("Số GCN của từng đơn lấy từ MPLIS khi bấm 'Kiểm tra (chạy thử)'.")

    def _start(self, chay_thu: bool) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        excel = self._excel_madon()
        goc = self._goc() if excel else None
        if not goc:
            return

        bo_qua = self.chk_bo_qua_da_xong.isChecked()
        if not chay_thu:
            ids = thd.tach_danh_sach_id(thd.doc_id_tu_excel(excel))
            so_don = len(ids)
            if bo_qua:
                tien_do = ge.doc_tien_do(excel)
                so_don -= sum(ge.don_da_xong(tien_do, m) for m in ids)
            ans = QMessageBox.question(
                self,
                "Xác nhận đẩy GCN",
                f"Đẩy GCN cho {so_don} đơn{' chưa xong' if bo_qua else ''} trong:\n{excel}\n\nFile tìm trong:\n{goc}\n\n"
                "HSQ có đẩy GCN: mọi thành phần khác sẽ chuyển sang Giấy tờ.\n\n"
                "Thao tác GHI THẬT lên MPLIS. Nên bấm 'Kiểm tra (chạy thử)' trước. Tiếp tục?",
            )
            if ans != QMessageBox.StandardButton.Yes:
                return

        self._run_thread(self._worker, nguon, excel, goc, chay_thu, bo_qua)

    def _worker(self, nguon: dict[str, Any], excel: str, goc: str, chay_thu: bool, bo_qua_da_xong: bool) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            kq = ge.chay(
                client,
                excel,
                goc,
                chay_thu,
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self.bridge.progressSignal.emit(done, total, label),
                bo_qua_da_xong=bo_qua_da_xong,
            )
            if kq:
                self.bridge.excelKqSignal.emit(kq)
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
    win = DayGcnTheoExcelPySide6Window()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
