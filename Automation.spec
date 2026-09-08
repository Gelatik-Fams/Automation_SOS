# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets'), ('VERSION', '.')],
    hiddenimports=['customtkinter', 'PIL', 'PIL._tkinter_finder', 'openpyxl', 'openpyxl.chart', 'openpyxl.chart.label', 'openpyxl.formatting.rule', 'openpyxl.styles'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'IPython'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Automation',
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
    icon=['assets/app.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SOS Dashboard Automation',
)
app = BUNDLE(
    coll,
    name='Automation.app',
    icon='assets/app.ico',
    bundle_identifier=None,
)
