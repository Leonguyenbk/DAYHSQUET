# -*- coding: utf-8 -*-
"""Bản riêng: tra thông tin thiếu trực tiếp bằng tinhHinhDangKyId."""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

import lay_thong_tin_thieu_hsq as ltt
from cap_nhat_thong_tin_thieu_pyside6 import CapNhatThongTinThieuWindow

try:
    from _phien_ban_cap_nhat_thieu_theo_id import PHIEN_BAN
except ImportError:
    PHIEN_BAN = "dev"


class CapNhatThongTinThieuTheoIdWindow(CapNhatThongTinThieuWindow):
    WINDOW_TITLE = f"MPLIS • Cập nhật thông tin thiếu theo tinhHinhDangKyId • v{PHIEN_BAN}"


def main() -> int:
    # Chỉ tác động trong tiến trình của bản EXE riêng này; bản cũ vẫn giữ luồng cũ.
    ltt.dat_tra_cuu_truc_tiep_theo_id(True)
    app = QApplication(sys.argv)
    win = CapNhatThongTinThieuTheoIdWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
