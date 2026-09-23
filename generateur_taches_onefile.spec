# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, copy_metadata

highspy_datas = collect_data_files('highspy')
highspy_binaries = collect_dynamic_libs('highspy')

extras_datas = collect_data_files('highspy_extras')
extras_binaries = collect_dynamic_libs('highspy_extras')
extras_metadata = copy_metadata('highspy_extras')

a = Analysis(
    ['src/gui.py'],
    pathex=[],
    binaries=highspy_binaries + extras_binaries,
    datas=[('src/config_default.yaml', '.')] + highspy_datas + extras_datas + extras_metadata,
    hiddenimports=['highspy', 'highspy_extras'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas.io.formats.style', 'pandas.io.clipboard', 'sqlite3', 'unittest', 'doctest',
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='generateur_taches',
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