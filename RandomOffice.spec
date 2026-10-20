# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

block_cipher = None

a = Analysis(
    ['run_desktop.py'],
    pathex=['.'],
    binaries=[],
    datas=[(str(path), 'randomoffice/static')
           for path in Path('randomoffice/static').iterdir()
           if path.is_file() and path.suffix in ('.html', '.js', '.css', '.json', '.svg', '.png')],
    hiddenimports=[
        'randomoffice',
        'randomoffice.web_office',
        'randomoffice.performance.guard',
        'randomoffice.desktop',
        'randomoffice.ui',
        'randomoffice.word_richtext',
        'docx',
        'openpyxl',
        'pptx',
        'webview',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name='RandomOffice',
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
    icon='randomoffice/static/randomoffice.ico',
)
