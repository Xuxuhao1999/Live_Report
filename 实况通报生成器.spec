# -*- mode: python ; coding: utf-8 -*-

import glob as _glob

a = Analysis(
    ['report_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.')] + [(_f, '.') for _f in _glob.glob('*.docx')] + [('两直一白.xlsx', '.')],
    hiddenimports=['docx', 'docx.opc', 'docx.oxml', 'lxml'],
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
    name='天气通报生成器',
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
    icon=['icon.ico'],
)
