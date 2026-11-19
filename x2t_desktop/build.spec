# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect all CustomTkinter assets, themes, and modules
ctk_datas, ctk_binaries, ctk_hiddenimports = collect_all('customtkinter')

datas = [
    ('templates', 'templates'),
    ('../sample_data', 'sample_data'),
] + ctk_datas

hiddenimports = [
    'pandas',
    'openpyxl',
    'requests',
    'customtkinter',
    'PIL',
    'PIL._imagingtk',
    'PIL.ImageTk',
    'xml.etree.ElementTree',
] + ctk_hiddenimports

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=ctk_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    name='X2T_Converter',
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
