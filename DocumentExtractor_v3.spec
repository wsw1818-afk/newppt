# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

tkinterdnd2_datas = collect_data_files('tkinterdnd2')

a = Analysis(
    ['ppt_extractor_v3.py'],
    pathex=[],
    binaries=[],
    datas=tkinterdnd2_datas,
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32clipboard',
        'pythoncom',
        'pywintypes',
        'pptx',
        'openpyxl',
        'openpyxl.drawing.image',
        'PIL',
        'PIL.Image',
        'docx',
        'pypdf',
        'lxml',
        'tkinterdnd2',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['cryptography', 'cffi', '_cffi_backend'],
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
    name='DocumentExtractor_v3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
