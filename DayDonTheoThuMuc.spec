# -*- mode: python ; coding: utf-8 -*-
# Build: python -m PyInstaller DayDonTheoThuMuc.spec --noconfirm
# Phiên bản = ngày build (YYYY.MM.DD): ghi vào tên exe; exe cùng ngày đang mở → thêm giờ phút.
import os
import time

PHIEN_BAN = time.strftime("%Y.%m.%d")
_cu = os.path.join(SPECPATH, "dist", f"DayDonTheoThuMuc_v{PHIEN_BAN}.exe")
if os.path.exists(_cu):
    try:
        open(_cu, "ab").close()
    except PermissionError:
        PHIEN_BAN = time.strftime("%Y.%m.%d.%H%M")
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

import sys
sys.path.insert(0, SPECPATH)
from flet_client_bundle import datas_flet_client

datas = collect_data_files("flet") + datas_flet_client()  # client đi kèm: máy khác không phải tải từ GitHub

hiddenimports = ['selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chromium.webdriver', 'selenium.webdriver.chromium.options', 'selenium.webdriver.chromium.service', 'selenium.webdriver.remote.webdriver', 'selenium.webdriver.common.keys', 'selenium.webdriver.common.by', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions']
hiddenimports += collect_submodules("openpyxl")

a = Analysis(
    ['day_don_theo_thu_muc_flet.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=f'DayDonTheoThuMuc_v{PHIEN_BAN}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app.ico'],
)
