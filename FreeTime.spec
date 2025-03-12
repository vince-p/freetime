# command to compile program
# pyinstaller --clean FreeTime.spec

# -*- mode: python ; coding: utf-8 -*-

import os

# Define paths relative to spec file location
SPEC_ROOT = os.path.dirname(os.path.abspath('__file__'))
SOURCE_PATH = os.path.join(SPEC_ROOT, 'source')
DIST_PATH = os.path.join(SPEC_ROOT, 'dist')

block_cipher = None

a = Analysis(
    [os.path.join(SOURCE_PATH, 'freetime.py')],  # Main script in source directory
    pathex=[SOURCE_PATH],
    binaries=[],
    datas=[(os.path.join(SOURCE_PATH, 'icon.png'), '.')],  # Icon in source directory
    hiddenimports=['PIL._tkinter_finder', 'pynput', 'pynput.keyboard', 'pynput.mouse'],
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
    name='Freetime',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(SOURCE_PATH, 'icon.png'),  # Icon in source directory
    distpath=DIST_PATH  # Specify dist directory
)

# macOS specific
app = BUNDLE(
    exe,
    name='Freetime.app',
    icon=os.path.join(SOURCE_PATH, 'icon.png'),  # Icon in source directory
    bundle_identifier='com.vince.freetime',
    info_plist={
        'LSUIElement': True,  # Makes the app a background app on macOS
    },
)