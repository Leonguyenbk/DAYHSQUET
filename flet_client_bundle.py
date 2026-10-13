# -*- coding: utf-8 -*-
"""
Đóng gói sẵn Flet desktop client vào exe (dùng trong file .spec của PyInstaller).

Không có client đi kèm, exe chạy lần đầu trên máy khác sẽ tải flet-windows.zip từ GitHub — máy có
proxy / chặn mạng / thiếu chứng chỉ gốc báo lỗi "SSL: CERTIFICATE_VERIFY_FAILED". flet_desktop tìm
archive tại flet_desktop/app/flet-windows.zip trước khi tải, nên đóng zip từ client đã cache trên máy
build (~/.flet/client/flet-desktop-*-<version>) và đưa vào đúng chỗ đó.
"""

from __future__ import annotations

import hashlib
import os
import zipfile
from pathlib import Path

THU_MUC_BUILD = Path(__file__).resolve().parent / "build" / "flet_client"


def _thu_muc_cache() -> Path:
    import flet_desktop.version

    goc = Path.home() / ".flet" / "client"
    ver = flet_desktop.version.version
    ung_vien = [p for p in goc.glob(f"flet-desktop-*-{ver}") if (p / "flet" / "flet.exe").is_file()]
    if not ung_vien:
        raise SystemExit(
            f"Chưa có Flet client {ver} trong {goc} — chạy thử 1 tool flet bằng Python (có mạng) để nó tải về, "
            "rồi build lại."
        )
    return ung_vien[0]


def datas_flet_client() -> list[tuple[str, str]]:
    """[(flet-windows.zip, 'flet_desktop/app'), (.sha256, ...)] để cộng vào `datas` của Analysis."""
    cache = _thu_muc_cache()
    THU_MUC_BUILD.mkdir(parents=True, exist_ok=True)
    zip_path = THU_MUC_BUILD / "flet-windows.zip"
    ten_nguon = THU_MUC_BUILD / "nguon.txt"
    if not zip_path.exists() or not ten_nguon.exists() or ten_nguon.read_text(encoding="utf-8") != str(cache):
        tam = zip_path.with_suffix(".tmp")
        with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in sorted((cache / "flet").rglob("*")):
                if p.is_file():
                    zf.write(p, p.relative_to(cache).as_posix())
        os.replace(tam, zip_path)
        ten_nguon.write_text(str(cache), encoding="utf-8")
    # Sidecar "<sha256> <size>" để exe khỏi băm lại zip mỗi lần mở (flet_desktop.__get_archive_fingerprint).
    h = hashlib.sha256()
    with open(zip_path, "rb") as f:
        while chunk := f.read(1024 * 1024):
            h.update(chunk)
    sidecar = zip_path.with_name(zip_path.name + ".sha256")
    sidecar.write_text(f"{h.hexdigest()} {zip_path.stat().st_size}")
    return [(str(zip_path), "flet_desktop/app"), (str(sidecar), "flet_desktop/app")]
