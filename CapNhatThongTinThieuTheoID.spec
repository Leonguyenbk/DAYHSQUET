# -*- mode: python ; coding: utf-8 -*-
# Build: python -m PyInstaller CapNhatThongTinThieuTheoID.spec --noconfirm
import os
import time

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

PHIEN_BAN = time.strftime("%Y.%m.%d")
_cu = os.path.join(SPECPATH, "dist", f"CapNhatThongTinThieuTheoID_v{PHIEN_BAN}.exe")
if os.path.exists(_cu):
    try:
        open(_cu, "ab").close()
    except PermissionError:
        PHIEN_BAN = time.strftime("%Y.%m.%d.%H%M")
with open(os.path.join(SPECPATH, "_phien_ban_cap_nhat_thieu_theo_id.py"), "w", encoding="utf-8") as fp:
    fp.write(
        "# Tự sinh khi build CapNhatThongTinThieuTheoID.spec — không sửa tay.\n"
        f'PHIEN_BAN = "{PHIEN_BAN}"\n'
    )

hiddenimports = [
    'selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.options',
    'selenium.webdriver.chrome.service', 'selenium.webdriver.chromium.webdriver',
    'selenium.webdriver.chromium.options', 'selenium.webdriver.chromium.service',
    'selenium.webdriver.remote.webdriver', 'selenium.webdriver.common.keys',
    'selenium.webdriver.common.by', 'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',
]
hiddenimports += collect_submodules("openpyxl") + collect_submodules("tach_gcn")
hiddenimports += ["_phien_ban_cap_nhat_thieu_theo_id"]
hiddenimports += collect_submodules("google.genai", filter=lambda name: ".tests" not in name)

a = Analysis(
    ['cap_nhat_thong_tin_thieu_theo_id_pyside6.py'],
    pathex=[SPECPATH, os.path.join(SPECPATH, 'src')],
    binaries=[],
    datas=collect_data_files("pypdfium2") + collect_data_files("pypdfium2_raw"),
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['flet', 'flet_desktop', 'matplotlib', 'PyQt5', 'PyQt6'],
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
    name=f'CapNhatThongTinThieuTheoID_v{PHIEN_BAN}',
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
