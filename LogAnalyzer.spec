# -*- mode: python ; coding: utf-8 -*-
# NOTE: Use build_exe.bat to build, or run:
#   pyinstaller --clean --distpath "Log Analyzer" LogAnalyzer.spec

import sys
import os

# Get tkinterdnd2 path for including its DLL files
try:
    import tkinterdnd2
    tkdnd_path = os.path.dirname(tkinterdnd2.__file__)
    tkdnd_datas = [(tkdnd_path, 'tkinterdnd2')]
except ImportError:
    tkdnd_datas = []

a = Analysis(
    ['log_analyzer_gui.py'],
    pathex=[],
    binaries=[],
    datas=tkdnd_datas,
    hiddenimports=['tkinterdnd2'],
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
    name='LogAnalyzer',
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
    icon=['sasquatch.ico'],
)
