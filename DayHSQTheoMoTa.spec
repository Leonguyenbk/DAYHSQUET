# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

import sys
sys.path.insert(0, SPECPATH)
from flet_client_bundle import datas_flet_client

datas = collect_data_files("flet") + datas_flet_client()  # client đi kèm: máy khác không phải tải từ GitHub

hiddenimports = ['selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chromium.webdriver', 'selenium.webdriver.chromium.options', 'selenium.webdriver.chromium.service', 'selenium.webdriver.remote.webdriver', 'selenium.webdriver.common.keys', 'selenium.webdriver.common.by', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions']
hiddenimports += collect_submodules("openpyxl")

a = Analysis(
    ['day_hsq_theo_mota_flet.py'],
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
    name='DayHSQTheoMoTa',
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
