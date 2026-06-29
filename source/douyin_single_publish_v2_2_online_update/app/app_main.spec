# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

a = Analysis(
    ['抖音单图文发布v2.2.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('google4', 'google4'),
        ('douyin_gui_config.json', '.'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'playwright',
        'playwright.sync_api',
        'playwright._impl',
        'tkinter',
        'PIL',
    ],
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
    name='抖音单图文发布v2.2_新增随机图片',
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
    icon='assets\\app_logo.ico',
)
