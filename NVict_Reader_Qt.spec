# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller-spec voor de PySide6-editie (test-build, fase 8).

Nog geen installer/signing/release-pipeline - alleen een standalone
onedir-build om lokaal te draaien en te testen, analoog aan
NVict_Reader.spec maar zonder de pywin32/PIL-afhankelijkheden die de
tkinter-versie nodig had (print gaat via QPrinter, geen PIL meer nodig
sinds fase 1).
"""
from PyInstaller.utils.hooks import collect_all

datas = [
    ('icons', 'icons'),
    ('favicon.ico', '.'),
]
binaries = []
hiddenimports = []
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


# PyMuPDF's optionele tabel-export (fitz.table) trekt pandas/pyarrow/lxml/
# openpyxl mee binnen via collect_all('fitz'), maar die worden nooit
# aangeroepen (geen find_tables()/to_pandas() in deze app) - runtime-check
# bevestigt dat `import fitz` ze niet eager laadt. Uitsluiten scheelt
# ~250MB zonder functionaliteit te verliezen.
excludes = [
    'pandas', 'pyarrow', 'openpyxl', 'lxml', 'IPython', 'matplotlib', 'numba', 'scipy',
    # De Qt-app rendert al zonder PIL/numpy (fitz-pixmap -> QImage direct,
    # zie pdf_view.py); `import fitz` laadt ze niet eager (runtime-check),
    # dus ze zijn alleen bereikbaar via ongebruikte fitz-hulpfuncties.
    'PIL', 'numpy',
]

a = Analysis(
    ['NVict_Reader_Qt.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# One-folder (onedir) build, zelfde reden als bij de tkinter-app: snellere
# opstart dan een onefile-build die zichzelf bij elke start moet uitpakken.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NVict_Reader_Qt',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['favicon.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='NVict_Reader_Qt',
)
