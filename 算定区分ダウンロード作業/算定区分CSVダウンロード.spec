# -*- mode: python ; coding: utf-8 -*-

import os
import playwright

driver_dir = os.path.join(os.path.dirname(playwright.__file__), 'driver')

a = Analysis(
    ['算定区分CSVダウンロード.py'],
    pathex=[],
    binaries=[],
    datas=[(driver_dir, 'playwright/driver')],
    hiddenimports=[],
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
    name='算定区分CSVダウンロード',
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
