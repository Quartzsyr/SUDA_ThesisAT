# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files


block_cipher = None

datas = [('assets', 'assets')]
datas += collect_data_files('qfluentwidgets', include_py_files=False)

excludes = [
    'PyQt5',
    'PySide2',
    'PySide6',
    'matplotlib',
    'pandas',
    'scipy',
    'pytest',
    'pyarrow',
    'openpyxl',
    'sqlalchemy',
    'jupyter',
    'IPython',
    'tkinter',
    '_tkinter',
    'pygame',
]

a = Analysis(
    ['app.py'],
    pathex=['E:\\paperwrite'],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ThesisFlow',
    icon='E:\\paperwrite\\icon.ico',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ThesisFlow',
)

