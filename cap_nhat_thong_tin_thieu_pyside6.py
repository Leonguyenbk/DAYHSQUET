# -*- coding: utf-8 -*-
"""Giao diện PySide6: cập nhật thông tin thiếu lên MPLIS.

Nạp Excel:
  - chỉ có cột madon (mã đơn, hoặc câu thông báo 'Tình hình đăng ký ... không có ...') → TỰ ĐỘNG: tải HSQ → AI
    đọc đúng mục thiếu → ghi MPLIS các mục chắc chắn (Đã đọc / Mặc định; tick thêm để ghi cả Cần kiểm tra);
  - ket_qua_thong_tin_thieu_* / ket_qua_cap_nhat_thieu_* đã duyệt (cột duyet) → chỉ ghi các dòng duyet = x.
Lõi xử lý: cap_nhat_thong_tin_thieu.py. Phiên đăng nhập, giải Captcha, nhật ký dùng lại từ
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
    QPushButton,
    QVBoxLayout,
    QWidget,
)

import cap_nhat_don_hsq as core
import cap_nhat_thong_tin_thieu as cnt
import tai_hsq_don as thd
from tach_gcn.gemini import DEFAULT_MODEL
from tai_hsq_don_pyside6 import APP_DIR, TaiHsqPySide6Window

try:
    from _phien_ban_cap_nhat_thieu import PHIEN_BAN  # sinh khi build CapNhatThongTinThieu.spec
except ImportError:
    PHIEN_BAN = "dev"


class CapNhatThongTinThieuWindow(TaiHsqPySide6Window):
    WINDOW_TITLE = f"MPLIS • Tự lấy & cập nhật thông tin thiếu • v{PHIEN_BAN}"

    def _build_body(self) -> QWidget:
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(8, 10, 8, 8)
        layout.setSpacing(8)
        self.che_do = ""  # "madon" | "duyet"
        self.excel_kq = ""

        lbl = QLabel("Excel cột madon (mã đơn hoặc câu thông báo thiếu) — hoặc Excel kết quả đã duyệt:")
        lbl.setObjectName("mutedText")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)
        row = QHBoxLayout()
        self.inp_excel = QLineEdit()
        self.inp_excel.setPlaceholderText("Chọn file Excel...")
        self.inp_excel.setReadOnly(True)
        btn = QPushButton("📄 Chọn Excel")
        btn.setObjectName("btnPurple")
        btn.clicked.connect(self._chon_excel)
        row.addWidget(self.inp_excel, 1)
        row.addWidget(btn)
        layout.addLayout(row)

        self.lbl_dem = QLabel("")
        self.lbl_dem.setObjectName("monoCyan")
        self.lbl_dem.setWordWrap(True)
        layout.addWidget(self.lbl_dem)

        row_out = QHBoxLayout()
        self.inp_out_dir = QLineEdit(str(APP_DIR / "cap_nhat_thieu"))
        btn_out = QPushButton("📂 Chọn")
        btn_out.clicked.connect(self._chon_out_dir)
        row_out.addWidget(self.inp_out_dir, 1)
        row_out.addWidget(btn_out)
        layout.addLayout(row_out)

        row_ai = QHBoxLayout()
        self.inp_api_key = QLineEdit(os.environ.get("GEMINI_API_KEY", ""))
        self.inp_api_key.setPlaceholderText("Gemini API key (chế độ tự động)")
        self.inp_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.inp_model = QLineEdit(os.environ.get("GEMINI_MODEL", DEFAULT_MODEL))
        self.inp_rpm = QLineEdit("30")
        self.inp_rpm.setFixedWidth(52)
        row_ai.addWidget(self.inp_api_key, 3)
        row_ai.addWidget(self.inp_model, 2)
        row_ai.addWidget(self.inp_rpm)
        layout.addLayout(row_ai)

        self.chk_chay_thu = QCheckBox("Chạy thử (chỉ ghi kế hoạch vào Excel, không gửi MPLIS)")
        self.chk_chay_thu.setChecked(True)
        self.chk_chay_thu.setStyleSheet("color: #fbbf24;")
        self.chk_can_kiem_tra = QCheckBox("Ghi cả dòng 'Cần kiểm tra' (AI đọc mờ / loại GCN suy luận)")
        self.chk_dung_file_da_tai = QCheckBox("Đơn đã tải HSQ rồi thì không tải lại")
        self.chk_dung_file_da_tai.setChecked(True)
        for chk in (self.chk_chay_thu, self.chk_can_kiem_tra, self.chk_dung_file_da_tai):
            layout.addWidget(chk)

        note = QLabel(
            "• Chế độ tự động: mỗi đơn lấy mục thiếu từ câu thông báo trong Excel (ô chỉ có mã đơn thì hỏi MPLIS) → "
            "tải HSQ → AI đọc trên GCN / đơn đăng ký → ghi MPLIS các mục Đã đọc + Mặc định. Dòng Cần kiểm tra để "
            "lại trong Excel kết quả: đánh x cột duyet rồi nạp lại file đó để ghi tiếp.\n"
            "• Trước khi ghi luôn hỏi lại MPLIS: mục không còn bị báo thiếu thì bỏ qua (không ghi đè). Ghi xong đọc "
            "lại xác nhận. Không ghi gì khi đang Chạy thử."
        )
        note.setObjectName("mutedText")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.btn_chay = QPushButton("☁ Lấy & cập nhật lên MPLIS")
        self.btn_chay.setObjectName("btnSuccess")
        self.btn_chay.clicked.connect(self._xac_nhan)
        layout.addWidget(self.btn_chay, 0, Qt.AlignmentFlag.AlignLeft)
        self.lbl_excel_tai = QLabel("")
        self.lbl_excel_tai.setObjectName("monoGreen")
        self.lbl_excel_tai.setWordWrap(True)
        layout.addWidget(self.lbl_excel_tai)
        layout.addStretch()
        return body

    def _chon_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Chọn Excel", str(APP_DIR), "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        path = os.path.normpath(path)
        try:
            if cnt.la_excel_da_duyet(path):
                rows = cnt.doc_excel(path)
                duyet = [r for r in rows if cnt.la_duyet(r.get("duyet"))]
                xong = sum(r.get("trang_thai_cap_nhat") == "Đã cập nhật" for r in duyet)
                self.che_do = "duyet"
                self.lbl_dem.setText(f"Excel đã duyệt: {len(rows)} dòng • {len(duyet)} đánh x"
                                     + (f" ({xong} đã cập nhật trước)" if xong else "") + " → chỉ ghi dòng đánh x")
            else:
                ids, muc = cnt.doc_dau_vao_madon(path)
                if not ids and not muc:
                    QMessageBox.warning(self, "Không có dữ liệu", "Excel không có mã đơn / câu thông báo thiếu nào.")
                    return
                self.che_do = "madon"
                self.lbl_dem.setText(f"Tự động: {len({m.id_don for m in muc}) + len(ids)} đơn • {len(muc)} mục thiếu "
                                     f"từ câu thông báo" + (f" • {len(ids)} đơn chỉ có mã (hỏi MPLIS)" if ids else ""))
        except Exception as exc:
            QMessageBox.critical(self, "Lỗi đọc Excel", str(exc))
            return
        self.inp_excel.setText(path)

    # ------------------------------------------------------------------ Chạy / Dừng
    def _set_running(self, running: bool) -> None:
        self.btn_chay.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _on_excel_tai(self, path: str) -> None:
        self.excel_kq = path
        self.lbl_excel_tai.setText(f"Excel kết quả: {path}")

    def _mo_excel(self) -> None:
        path = self.excel_kq or self.inp_excel.text().strip()
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Chưa có Excel", "Chưa có file Excel.")
            return
        os.startfile(path)

    def _xac_nhan(self) -> None:
        if self.is_running:
            return
        nguon = self._nguon_session()
        if not nguon:
            return
        path = self.inp_excel.text().strip()
        if not os.path.isfile(path) or not self.che_do:
            QMessageBox.warning(self, "Thiếu file Excel", "Chọn file Excel trước.")
            return
        cfg: dict[str, Any] = {"path": path, "chay_thu": self.chk_chay_thu.isChecked(), "che_do": self.che_do}
        if self.che_do == "madon":
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
            cfg.update(api_key=api_key, model=self.inp_model.text().strip() or DEFAULT_MODEL, rpm=rpm,
                       out_dir=self.inp_out_dir.text().strip() or str(APP_DIR / "cap_nhat_thieu"),
                       ghi_ca=self.chk_can_kiem_tra.isChecked(), dung_file=self.chk_dung_file_da_tai.isChecked())
        if not cfg["chay_thu"]:
            pham_vi = ("các dòng đánh x cột duyet" if self.che_do == "duyet" else
                       "giá trị AI đọc được (Đã đọc + Mặc định" + (" + Cần kiểm tra" if cfg.get("ghi_ca") else "") + ")")
            ans = QMessageBox.question(self, "Xác nhận ghi MPLIS",
                                       f"Ghi THẬT {pham_vi} lên MPLIS theo:\n{path}\n\nTiếp tục?")
            if ans != QMessageBox.StandardButton.Yes:
                return
        self._run_thread(self._worker, nguon, cfg)

    def _worker(self, nguon: dict[str, Any], cfg: dict[str, Any]) -> None:
        ket_thuc = "Hoàn tất"
        tien_do = lambda done, total, label: self.bridge.progressSignal.emit(done, total, label)  # noqa: E731
        try:
            client = self._tao_client(nguon)
            if cfg["che_do"] == "duyet":
                cnt.chay(client, cfg["path"], cfg["chay_thu"], nen_dung=lambda: self.stop_flag, bao_tien_do=tien_do)
                self.bridge.excelTaiSignal.emit(cfg["path"])
            else:
                ids, muc = cnt.doc_dau_vao_madon(cfg["path"])
                excel = cnt.chay_tu_madon(client, ids, cfg["out_dir"], cfg["api_key"], cfg["model"], cfg["rpm"],
                                          cfg["chay_thu"], cfg["ghi_ca"], cfg["dung_file"], muc_dau_vao=muc,
                                          nen_dung=lambda: self.stop_flag, bao_tien_do=tien_do)
                self.bridge.excelTaiSignal.emit(excel)
            if self.stop_flag:
                ket_thuc = "Đã dừng"
        except PermissionError:
            thd.log("❌ Không ghi được Excel — đóng file Excel rồi chạy lại.")
            ket_thuc = "Lỗi"
        except Exception as exc:
            thd.log(f"❌ {core.rut_gon_text(exc, 300)}")
            ket_thuc = "Lỗi"
        finally:
            self.bridge.doneSignal.emit(ket_thuc)


def main() -> int:
    app = QApplication(sys.argv)
    win = CapNhatThongTinThieuWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
