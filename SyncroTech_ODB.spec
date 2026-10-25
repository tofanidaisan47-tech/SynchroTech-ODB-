# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# ── Clean Datas Collection ───────────────────────────────────────────
datas = [
    ('password_config.json', '.'),
]

# Collect only source code, templates, and assets (ignore venv, git, build, tests)
for sub in ['datalogger_repo', 'synchro_app']:
    for root, dirs, files in os.walk(sub):
        dirs[:] = [
            d for d in dirs
            if d not in ['.venv', '.git', '__pycache__', 'build', 'dist', 'tests', '.pytest_cache', '_internal']
        ]
        for f in files:
            if not f.endswith('.pyc') and not f.endswith('.exe'):
                src = os.path.join(root, f)
                rel_dir = root
                datas.append((src, rel_dir))

binaries = []
hiddenimports = [
    'serial',
    'serial.tools',
    'serial.tools.list_ports',
    'pyvisa',
    'pyvisa_py',
    'zeroconf',
    'ifaddr',
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.styles.colors',
    'openpyxl.worksheet',
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',
    'winsound',
    'psutil',
    'yaml',
]

# Collect all hooks for core complex packages
for pkg in ['customtkinter', 'pyvisa', 'pyvisa_py', 'openpyxl', 'zeroconf']:
    try:
        ret = collect_all(pkg)
        datas += ret[0]
        binaries += ret[1]
        hiddenimports += ret[2]
    except Exception as e:
        print(f"Warning collecting {pkg}: {e}")

a = Analysis(
    ['launcher.py'],
    pathex=['.', 'datalogger_repo', 'synchro_app'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['datalogger_repo.tests', 'tests', 'synchro_app.tests'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SyncroTech ODB',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='datalogger_repo/logo.ico' if Path('datalogger_repo/logo.ico').exists() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SyncroTech_ODB',
)
