# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['getExcel.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\20242176\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\fake_useragent', './fake_useragent')],
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
    name='getExcel',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['image\\favicon.ico'],
)
