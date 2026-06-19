# -*- mode: python ; coding: utf-8 -*-
"""InvoiceRenamer PyInstaller 打包配置

用法: pyinstaller InvoiceRenamer.spec
"""

import os
from config import APP_VERSION

block_cipher = None

a = Analysis(
    ['invoice_renamer_ui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('rules.json', '.'),       # 发票解析规则
        ('icon.ico', '.'),         # 应用图标
    ],
    hiddenimports=[
        'pdfplumber',
        'pdfminer.high_level',
        'customtkinter',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test',
        'unittest',
        'pydoc',
        'doctest',
        'test',
        'PyMuPDF',
        'fitz',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=f'InvoiceRenamer_{APP_VERSION}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='icon.ico',
)
