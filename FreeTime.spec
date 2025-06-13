# command to compile program
# pyinstaller --clean FreeTime.spec
# double check this does not delete icon.png

# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import platform
import certifi
import subprocess
import tempfile
import plistlib
import time
from PIL import Image
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Define paths
SPEC_ROOT = os.path.dirname(os.path.abspath(sys.argv[0]))
SOURCE_PATH = os.path.join(SPEC_ROOT, 'source')
DIST_PATH = os.path.join(SPEC_ROOT, 'dist')

# PNG icon path
icon_png = os.path.join(SOURCE_PATH, 'icon.png')

def create_proper_icns(png_path, output_path):
    """Create a properly formatted ICNS file with all required resolutions"""
    try:
        # Create temporary iconset directory
        temp_dir = tempfile.mkdtemp()
        iconset_dir = os.path.join(temp_dir, 'icon.iconset')
        os.makedirs(iconset_dir)

        # Create all required sizes (Apple's recommended sizes)
        img = Image.open(png_path)
        sizes = [
            (16, 1), (16, 2),  # 16x16 and 32x32
            (32, 1), (32, 2),  # 32x32 and 64x64
            (128, 1), (128, 2),  # 128x128 and 256x256
            (256, 1), (256, 2),  # 256x256 and 512x512
            (512, 1), (512, 2)   # 512x512 and 1024x1024
        ]

        for size, scale in sizes:
            actual_size = size * scale
            filename = f"icon_{size}x{size}{'' if scale == 1 else '@2x'}.png"
            resized = img.resize((actual_size, actual_size), Image.LANCZOS)
            resized.save(os.path.join(iconset_dir, filename))

        # Convert to ICNS using iconutil
        subprocess.run([
            'iconutil', '-c', 'icns', iconset_dir, '-o', output_path
        ], check=True)
        return True
    except Exception as e:
        print(f"ICNS creation failed: {e}")
        return False

block_cipher = None

a = Analysis(
    [os.path.join(SOURCE_PATH, 'freetime.py')],
    pathex=[SOURCE_PATH],
    binaries=[],
    datas=[
        (icon_png, '.'),
        (certifi.where(), '.'),
        *collect_data_files('pystray', subdir='_darwin')
    ],
    hiddenimports=[
        'pystray._darwin',
        'objc',
        'AppKit',
        'Foundation',
        'Quartz',
        'CoreFoundation',
        'PIL',
        'PIL._tkinter_finder',
        'pynput.keyboard._darwin',
        'pynput.mouse._darwin'
    ],
    hookspath=['.'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

if platform.system() == 'Darwin':
    a.binaries = [x for x in a.binaries if not x[0].startswith('libCoreFoundation')]

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
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    #codesign_identity=None,
    entitlements_file=None,
    icon=icon_png,
    distpath=DIST_PATH,
    codesign_identity='-',  # Ad-hoc signing
    #entitlements_file=None
)

app = BUNDLE(
    exe,
    name='Freetime.app',
    icon=None,  # Will be set in post-build
    bundle_identifier='com.vince.freetime',
    info_plist={
        'LSUIElement': True,
        'NSAppleEventsUsageDescription': 'Needs access to send Apple Events',
        'NSInputMonitoringUsageDescription': 'Needs to monitor keyboard input',
        'CFBundleDisplayName': 'Freetime',
        'CFBundleName': 'Freetime',
        'CFBundleExecutable': 'Freetime',
        'CFBundleVersion': '0.8',
        'CFBundleShortVersionString': '0.8',
        'NSHumanReadableCopyright': 'Copyright © 2025 Vince Polito',
        'CFBundleIconFile': 'icon.icns'  # Explicitly set here
    },
)

# Updated post-build function to include DMG creation
def _post_build():
    """Post-build steps to sign the app and create a DMG."""
    app_path = os.path.join(DIST_PATH, 'Freetime.app')
    app_name = os.path.basename(app_path)
    dmg_name = 'Freetime.dmg'
    dmg_path = os.path.join(DIST_PATH, dmg_name)

    # --- Part 1: Icon and App Signing (from your original script) ---
    try:
        # Create proper ICNS file
        resources_path = os.path.join(app_path, 'Contents', 'Resources')
        os.makedirs(resources_path, exist_ok=True)
        icns_path = os.path.join(resources_path, 'icon.icns')
        if create_proper_icns(icon_png, icns_path):
            print("Successfully created ICNS file")

        # Ad-hoc sign the app and remove quarantine
        print(f"Signing and preparing '{app_name}' for Gatekeeper...")
        subprocess.run(['xattr', '-cr', app_path], check=True)
        subprocess.run(['codesign', '--force', '--deep', '--sign', '-', app_path], check=True)
        print("App is now ad-hoc signed.")

        # Force Finder to update the icon
        subprocess.run(['touch', app_path], check=True)
        subprocess.run(['killall', 'Finder'], check=True)

    except Exception as e:
        print(f"Error during app signing or icon creation: {e}")
        return # Stop if signing fails

    # --- Part 2: DMG Creation (newly added) ---
    try:
        print(f"\nCreating DMG: {dmg_name}...")

        # Check if DMG already exists and remove it
        if os.path.exists(dmg_path):
            os.remove(dmg_path)

        # Create the command to build the DMG
        # This uses a modern format (ULFO) which is compressed and read-only
        command = [
            'hdiutil', 'create',
            '-volname', 'Freetime',         # The name of the volume when mounted
            '-srcfolder', app_path,         # The source to package
            '-ov',                          # Overwrite existing file
            '-format', 'ULFO',              # Compressed, read-only format
            dmg_path                        # The output path for the .dmg file
        ]

        # Run the command
        subprocess.run(command, check=True)

        print("----------------------------------------------------")
        print(f"Successfully created DMG: {dmg_path}")
        print("----------------------------------------------------")

    except Exception as e:
        print(f"Error during DMG creation: {e}")


if platform.system() == 'Darwin': # Only register for macOS
    import atexit
    atexit.register(_post_build)
    print("Registered macOS post-build hook.")
else:
    print("Skipping macOS post-build hook on non-Darwin platform.")