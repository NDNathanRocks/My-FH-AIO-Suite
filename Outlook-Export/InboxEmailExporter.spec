# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:/Users/ndsouza/OneDrive - HealthBC/Desktop/myGithub/Outlook-Export/inbox_email_exporter_v1_5_2.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['win32com.client', 'pythoncom', 'pywintypes', 'openpyxl'],
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
    name='InboxEmailExporter',
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
