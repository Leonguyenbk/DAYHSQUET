# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['tai_hsq_don_flet.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chromium.webdriver', 'selenium.webdriver.chromium.options', 'selenium.webdriver.chromium.service', 'selenium.webdriver.remote.webdriver', 'selenium.webdriver.common.keys', 'selenium.webdriver.common.by', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions'],
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
    name='TaiDayHSQ',
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
    version='C:\\Users\\Admin\\AppData\\Local\\Temp\\8a93df33-2c1e-4ab8-97aa-94c71c57ddf9',
    icon=['app.ico'],
)
