# -*- mode: python ; coding: utf-8 -*-
# Build: python -m PyInstaller TaiTachHSQ.spec --noconfirm
# Phiên bản = ngày build (YYYY.MM.DD): ghi vào tên exe + tiêu đề cửa sổ (_phien_ban_tai_tach.py).
import os
import time

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

PHIEN_BAN = time.strftime("%Y.%m.%d")
# exe cùng ngày đang mở (bị khoá) → không ghi đè, thêm giờ phút vào phiên bản
_cu = os.path.join(SPECPATH, "dist", f"TaiTachHSQ_v{PHIEN_BAN}.exe")
if os.path.exists(_cu):
    try:
        open(_cu, "ab").close()
    except PermissionError:
        PHIEN_BAN = time.strftime("%Y.%m.%d.%H%M")
with open(os.path.join(SPECPATH, "_phien_ban_tai_tach.py"), "w", encoding="utf-8") as fp:
    fp.write(f'# Tự sinh khi build TaiTachHSQ.spec — không sửa tay.\nPHIEN_BAN = "{PHIEN_BAN}"\n')

hiddenimports = ['selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chromium.webdriver', 'selenium.webdriver.chromium.options', 'selenium.webdriver.chromium.service', 'selenium.webdriver.remote.webdriver', 'selenium.webdriver.common.keys', 'selenium.webdriver.common.by', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions']
hiddenimports += collect_submodules("openpyxl") + collect_submodules("tach_gcn") + ["_phien_ban_tai_tach"]
hiddenimports += collect_submodules("google.genai", filter=lambda name: ".tests" not in name)

a = Analysis(
    ['tai_tach_hsq_don_pyside6.py'],
    pathex=[SPECPATH, os.path.join(SPECPATH, 'src')],
    binaries=[],
    datas=collect_data_files("pypdfium2") + collect_data_files("pypdfium2_raw"),
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['flet', 'flet_desktop', 'matplotlib', 'PyQt5', 'PyQt6'],  # tkinter: cap_nhat_don_hsq cần
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
    name=f'TaiTachHSQ_v{PHIEN_BAN}',
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
