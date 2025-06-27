# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    # ✅ Include required files inside "data/" and "resources/icons/" folders
    datas=[
        # ✅ Include all required data files
        ("data/company_names.txt", "data"),
        ("data/file_name_portions.txt", "data"),
        ("data/recent_filename_portions.txt", "data"),
        ("data/recent_save_locations.txt", "data"),
        ("data/theme_config.txt", "data"),

        # ✅ Include all required icons
        ("resources/icons/main_application_icon.ico", "resources/icons"),
        ("resources/icons/shortcut_icon.ico", "resources/icons"),
        ("resources/icons/taskbar_icon.ico", "resources/icons"),
    ],
    hiddenimports=['PyQt6', 'PyQt6.QtGui', 'PyQt6.QtWidgets', 'PyQt6.QtCore'],
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
    name='DocHandler',
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
    # ✅ Fix: Ensure the correct icon path is used
    icon="resources/icons/main_application_icon.ico",
)
