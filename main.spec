# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.')],
    hiddenimports=[
        'pandas',
        'pandas.core.arrays.integer',
        'pandas.core.arrays.boolean',
        'numpy',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'scipy', 'matplotlib', 'IPython', 'jupyter',
        'notebook', 'pytest', 'customtkinter', 'openpyxl',
        'PIL', 'tkinter',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# one-dir: EXE 本身不自解壓，避免 SAC / SmartScreen 封鎖
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,   # binaries 放進 COLLECT，不打包進 EXE
    name='JsonEditor',
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
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='JsonEditor',       # 輸出到 dist/JsonEditor/ 資料夾
)
