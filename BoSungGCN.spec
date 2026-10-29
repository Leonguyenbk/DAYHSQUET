# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['bo_sung_gcn_hsq_flet.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chromium.webdriver', 'selenium.webdriver.chromium.options', 'selenium.webdriver.chromium.service', 'selenium.webdriver.remote.webdriver', 'selenium.webdriver.common.keys', 'selenium.webdriver.common.by', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions', 'openpyxl'],
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
    name='BoSungGCN',
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
