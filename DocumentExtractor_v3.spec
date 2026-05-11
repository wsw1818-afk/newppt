# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['ppt_extractor_v3.py'],
    pathex=[],
    binaries=[],
    datas=[],
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
        'lxml',
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
    name='DocumentExtractor_v3',
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
