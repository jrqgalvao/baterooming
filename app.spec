# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('ui/menu_ui.html', 'ui'),
        ('ui/bate_rooming_ui.html', 'ui'),
        ('ui/match_nomes_ui.html', 'ui'),
        ('assets/app_icon.ico', 'assets'),
        ('assets/logo_generic_color.png', 'assets'),
        ('assets/logo_generic_white.png', 'assets'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'numpy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='app',
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
    version='app_version_info.txt',
    icon=['assets/app_icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='app',
)

from pathlib import Path
import shutil

shutil.copyfile('app.exe.config', Path(DISTPATH) / 'app' / 'app.exe.config')
