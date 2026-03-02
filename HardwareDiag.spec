# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('flet')
try:
    datas += collect_data_files('flet_desktop')
except Exception:
    pass

benchmarks_json = os.path.join(SPECPATH, 'benchmarks.json')
if os.path.isfile(benchmarks_json):
    datas.append((benchmarks_json, '.'))

binaries = []
bin_dir = os.path.join(SPECPATH, 'bin')
if os.path.isdir(bin_dir):
    for fname in os.listdir(bin_dir):
        fpath = os.path.join(bin_dir, fname)
        if os.path.isfile(fpath):
            binaries.append((fpath, 'bin'))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=['clr', 'fpdf', 'fpdf.enums', 'fpdf.output', 'pystray', 'PIL', 'pystray._win32'],
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
    name='HardwareDiag',
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
)
