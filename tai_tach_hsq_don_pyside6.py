# -*- coding: utf-8 -*-
"""Giao diện PySide6: tải toàn bộ HSQ của đơn → AI (Gemini) tách GCN; HSQ không có GCN thì lấy / tách đơn đăng ký.

Lõi xử lý: tai_tach_hsq_don.py. Phiên đăng nhập, giải Captcha, nhật ký dùng lại từ
tai_hsq_don_pyside6.TaiHsqPySide6Window.
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
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

import cap_nhat_don_hsq as core
import lay_thong_tin_thieu_hsq as ltt
import tai_hsq_don as thd
import tai_tach_hsq_don as tth
from tach_gcn.gemini import DEFAULT_MODEL
from tai_hsq_don_pyside6 import APP_DIR, TaiHsqPySide6Window


try:
    from _phien_ban_tai_tach import PHIEN_BAN  # sinh khi build TaiTachHSQ.spec
except ImportError:
    PHIEN_BAN = "dev"


class TaiTachHsqWindow(TaiHsqPySide6Window):
    WINDOW_TITLE = f"MPLIS • Tải & tách HSQ theo đơn — GCN / Đơn đăng ký / Thông tin thiếu • v{PHIEN_BAN}"

    def _build_body(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(8, 10, 8, 8)
        layout.setSpacing(8)

        row_nap = QHBoxLayout()
        lbl_nap = QLabel("Nhập tinhHinhDangKyId hoặc nạp từ Excel có cột madon:")
        lbl_nap.setObjectName("mutedText")
        btn_nap = QPushButton("📥 Nạp Excel")
        btn_nap.setObjectName("btnPurple")
        btn_nap.clicked.connect(self._nap_excel_ids)
        row_nap.addWidget(lbl_nap, 1)
        row_nap.addWidget(btn_nap)
        layout.addLayout(row_nap)

        self.inp_ids = QPlainTextEdit()
        self.inp_ids.setPlaceholderText("Mỗi dòng 1 tinhHinhDangKyId (hoặc cách nhau dấu phẩy / khoảng trắng)")
        layout.addWidget(self.inp_ids, 1)

        row_thieu = QHBoxLayout()
        lbl_thieu = QLabel("Thông tin thiếu thêm (tuỳ chọn — đã tự kiểm tra trên MPLIS thì không cần) — dán thông báo "
                           "hoặc nạp Excel/TXT:")
        lbl_thieu.setObjectName("mutedText")
        lbl_thieu.setWordWrap(True)
        btn_thieu = QPushButton("📥 Nạp file")
        btn_thieu.setObjectName("btnPurple")
        btn_thieu.clicked.connect(self._nap_thong_tin_thieu)
        row_thieu.addWidget(lbl_thieu, 1)
        row_thieu.addWidget(btn_thieu)
        layout.addLayout(row_thieu)

        self.inp_thieu = QPlainTextEdit()
        self.inp_thieu.setPlaceholderText(
            "Tình hình đăng ký 18566127 có thửa đất 4703930_0 có mục đích sử dụng 9205258 không có thời hạn sử dụng\n"
            "(đơn trong danh sách này cũng được tải & tách HSQ)")
        self.inp_thieu.textChanged.connect(self._dem_thong_tin_thieu)
        layout.addWidget(self.inp_thieu, 1)
        self.lbl_thieu = QLabel("")
        self.lbl_thieu.setObjectName("monoCyan")
        layout.addWidget(self.lbl_thieu)

        row_out = QHBoxLayout()
        self.inp_out_dir = QLineEdit(str(APP_DIR / "tai_tach_hsq"))
        btn_out = QPushButton("📂 Chọn")
        btn_out.clicked.connect(self._chon_out_dir)
        row_out.addWidget(self.inp_out_dir, 1)
        row_out.addWidget(btn_out)
        layout.addLayout(row_out)

        row_ai = QHBoxLayout()
        self.inp_api_key = QLineEdit(os.environ.get("GEMINI_API_KEY", ""))
        self.inp_api_key.setPlaceholderText("Gemini API key (mặc định biến môi trường GEMINI_API_KEY)")
        self.inp_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_model = QLineEdit(os.environ.get("GEMINI_MODEL", DEFAULT_MODEL))
        self.inp_model.setPlaceholderText("Model Gemini")
        self.inp_rpm = QLineEdit("30")
        self.inp_rpm.setPlaceholderText("RPM")
        self.inp_rpm.setFixedWidth(52)
        row_ai.addWidget(self.inp_api_key, 3)
        row_ai.addWidget(self.inp_model, 2)
        row_ai.addWidget(self.inp_rpm)
        layout.addLayout(row_ai)

        self.chk_dung_file_da_tai = QCheckBox("Đơn đã tải HSQ rồi thì không tải lại")
        self.chk_quet_het = QCheckBox("AI xét mọi file HSQ (không dừng khi đã đủ GCN đúng chủ)")
        self.chk_khong_loc = QCheckBox("Gửi ảnh nét mọi trang (không lọc trước — tốn hơn nhiều)")
        self.chk_tu_kiem_tra = QCheckBox("Tự kiểm tra đơn đang thiếu thông tin gì trên MPLIS → AI đọc thông tin đó")
        self.chk_tu_kiem_tra.setChecked(True)
        self.chk_tu_kiem_tra.setStyleSheet("color: #4ade80;")
        for chk in (self.chk_tu_kiem_tra, self.chk_dung_file_da_tai, self.chk_quet_het, self.chk_khong_loc):
            layout.addWidget(chk)

        note = QLabel(
            "• Tải TOÀN BỘ file HSQ mọi đơn (kể cả đơn chưa có GCN) → HSQ_goc/<madon>/.\n"
            "• Có trang GCN → tách GCN (GCN/, GCN_can_kiem_tra/, GCN_khong_khop/).\n"
            f"• HSQ không có GCN → tìm đơn đăng ký: file chứa đơn < {tth.GIOI_HAN_KHONG_TACH_DON // tth.MB} MB "
            "chép nguyên, lớn hơn thì tách trang đơn → DON/<madon>/.\n"
            "• Thông tin thiếu (MPLIS tự báo + danh sách dán thêm) → AI đọc trên GCN / đơn của HSQ vừa tải → Excel "
            "ket_qua_thong_tin_thieu_*.\n"
            "• Không ghi gì lên MPLIS."
        )
        note.setObjectName("mutedText")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.btn_chay = QPushButton("▶ Tải & tách")
        self.btn_chay.setObjectName("btnPrimary")
        self.btn_chay.clicked.connect(self._start_tai_tach)
        layout.addWidget(self.btn_chay, 0, Qt.AlignmentFlag.AlignLeft)

        self.lbl_excel_tai = QLabel("")
        self.lbl_excel_tai.setObjectName("monoGreen")
        self.lbl_excel_tai.setWordWrap(True)
        layout.addWidget(self.lbl_excel_tai)
        self.excel_kq = ""
        return body

    # ------------------------------------------------------------------ Chạy / Dừng
    def _set_running(self, running: bool) -> None:
        self.btn_chay.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _on_excel_tai(self, path: str) -> None:
        self.excel_kq = path
        self.lbl_excel_tai.setText(f"Excel kết quả: {path}")

    def _mo_excel(self) -> None:
        if not os.path.isfile(self.excel_kq):
            QMessageBox.warning(self, "Chưa có Excel", "Chưa có file Excel kết quả.")
            return
        os.startfile(self.excel_kq)

    def _nap_thong_tin_thieu(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file thông báo thiếu thông tin", str(APP_DIR), "Excel / văn bản (*.xlsx *.xlsm *.txt *.csv)"
        )
        if not path:
            return
        try:
            text = ltt.doc_file_dau_vao(path)
        except Exception as exc:
            QMessageBox.critical(self, "Lỗi đọc file", f"Không đọc được file: {exc}")
            return
        cu = self.inp_thieu.toPlainText().strip()
        self.inp_thieu.setPlainText(f"{cu}\n{text}".strip())
        self._append_log(f"Nạp thông báo thiếu từ {os.path.basename(path)}.")

    def _dem_thong_tin_thieu(self) -> None:
        muc, bo_qua = ltt.doc_thong_bao(self.inp_thieu.toPlainText())
        if not muc and not bo_qua:
            self.lbl_thieu.setText("")
            return
        self.lbl_thieu.setText(f"{len(muc)} mục thiếu / {len({m.id_don for m in muc})} đơn"
                               + (f" • {len(bo_qua)} mệnh đề không nhận ra" if bo_qua else ""))

    def _start_tai_tach(self) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        ids = thd.tach_danh_sach_id(self.inp_ids.toPlainText())
        muc_thieu, bo_qua = ltt.doc_thong_bao(self.inp_thieu.toPlainText())
        if not ids and not muc_thieu:
            QMessageBox.warning(self, "Thiếu ID", "Chưa có tinhHinhDangKyId hoặc thông báo thiếu thông tin hợp lệ.")
            return
        for menh_de in bo_qua[:10]:
            self._append_log(f"⚠ Không nhận ra: {menh_de}")
        api_key = self.inp_api_key.text().strip()
        if not api_key:
            QMessageBox.warning(self, "Thiếu API key", "Nhập Gemini API key.")
            return
        try:
            rpm = float(self.inp_rpm.text().strip() or "30")
            if rpm <= 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "RPM không hợp lệ", "RPM phải là số > 0.")
            return
        cfg = {
            "ids": ids,
            "out_dir": self.inp_out_dir.text().strip() or str(APP_DIR / "tai_tach_hsq"),
            "api_key": api_key,
            "model": self.inp_model.text().strip() or DEFAULT_MODEL,
            "rpm": rpm,
            "dung_file_da_tai": self.chk_dung_file_da_tai.isChecked(),
            "quet_het": self.chk_quet_het.isChecked(),
            "loc_truoc": not self.chk_khong_loc.isChecked(),
            "muc_thieu": muc_thieu,
            "tu_kiem_tra_thieu": self.chk_tu_kiem_tra.isChecked(),
        }
        self._run_thread(self._worker_tai_tach, nguon, cfg)

    def _worker_tai_tach(self, nguon: dict[str, Any], cfg: dict[str, Any]) -> None:
        ket_thuc = "Hoàn tất"
        try:
            client = self._tao_client(nguon)
            excel, _excel_tai, excel_thieu = tth.chay(
                client,
                cfg["ids"],
                cfg["out_dir"],
                cfg["api_key"],
                cfg["model"],
                cfg["rpm"],
                cfg["dung_file_da_tai"],
                nen_dung=lambda: self.stop_flag,
                bao_tien_do=lambda done, total, label: self.bridge.progressSignal.emit(done, total, label),
                loc_truoc=cfg["loc_truoc"],
                quet_het=cfg["quet_het"],
                muc_thieu=cfg["muc_thieu"],
                tu_kiem_tra_thieu=cfg["tu_kiem_tra_thieu"],
            )
            self.bridge.excelTaiSignal.emit(excel_thieu or excel)
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self.bridge.doneSignal.emit(ket_thuc)


def main() -> int:
    app = QApplication(sys.argv)
    win = TaiTachHsqWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
