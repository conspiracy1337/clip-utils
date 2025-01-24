# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Main.py'],
    pathex=[],  # Add your project path if needed (e.g., r'C:\path\to\project')
    binaries=[],
    datas=[
        ('ffmpeg/*', 'ffmpeg'),
        ('appicon.png', '.'),
        ('done.wav', '.')  # Added this line
    ],
    hiddenimports=[],
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
    name='ClipUtils',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Already specifies noconsole
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='appicon.png',
)