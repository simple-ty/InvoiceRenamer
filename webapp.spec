# -*- mode: python ; coding: utf-8 -*-
"""Invoice Renamer WebView 版 PyInstaller 打包配置

用法: pyinstaller webapp.spec
"""

import os
from config import APP_VERSION

block_cipher = None

a = Analysis(
    ['webapp.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('rules.json', '.'),          # 发票解析规则
        ('icon.ico', '.'),             # 应用图标
        ('webview/main.html', 'webview'),
        ('webview/css/main.css', 'webview/css'),
        ('webview/js/main.js', 'webview/js'),
    ],
    hiddenimports=[
        'pdfplumber',
        'pdfminer.high_level',
        'openpyxl',
        'webview',
        'webview.http',
        'tkinter',
        'tkinter.filedialog',
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
        'customtkinter',
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
    name=f'InvoiceRenamer_WebView_{APP_VERSION}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=r"%LOCALAPPDATA%\InvoiceRenamer\runtime",
    console=False,
    icon='icon.ico',
)
