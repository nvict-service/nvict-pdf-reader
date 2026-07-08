# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [
    ('icons', 'icons'),
    ('favicon.ico', '.'),
    ('pdf_file_icon.ico', '.'),
]
binaries = []
hiddenimports = ['win32print', 'win32ui', 'win32con', 'windnd']
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PIL')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['NVict_Reader.py'],
    pathex=[],
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

# One-folder (onedir) build: de app-bestanden staan als losse map naast de exe
# en worden NIET bij elke start uit één exe uitgepakt. Dit versnelt de opstart
# aanzienlijk (met name de eerste PDF). De installer bundelt de hele map.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NVict_Reader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPX uit: snellere opstart (geen DLL-decompressie) en minder antivirus-vertraging
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['favicon.ico'],
    version='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='NVict_Reader',
)
