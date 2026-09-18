# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.hooks import copy_metadata

datas = [('D:/Python/DAYHSQUET/gcn_file_finder/models/paddleocr', 'models/paddleocr'), ('D:/Python/DAYHSQUET/gcn_file_finder/config.example.json', '.')]
binaries = []
hiddenimports = []
datas += copy_metadata('imagesize')
datas += copy_metadata('pyclipper')
datas += copy_metadata('pypdfium2')
datas += copy_metadata('python-bidi')
datas += copy_metadata('shapely')
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('paddleocr')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('paddlex')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('paddle')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('cv2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['D:/Python/DAYHSQUET/gcn_file_finder/app.py'],
    pathex=['D:/Python/DAYHSQUET'],
    binaries=binaries,
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
    [],
    exclude_binaries=True,
    name='GCNFileFinder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GCNFileFinder',
)
