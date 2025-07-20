# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['radiologist_app.py'],
    pathex=[],
    binaries=[],
    datas=[('admin_settings.json', '.'), ('Reports', 'Reports')],
    hiddenimports=['tkinter', 'tkinter.ttk', 'docx', 'docx.enum.text', 'docx.shared', 'docx.oxml', 'docx.oxml.ns'],
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
    name='RadiologistApp',
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
