# -*- mode: python ; coding: utf-8 -*-
"""
SalesForecastGUI.spec
======================
PyInstaller build spec for the NVISH Sales Forecast Automation desktop
app. Produces a single, windowed (no console) .exe.

Build:
    pyinstaller SalesForecastGUI.spec

Output:
    dist/NVISH Sales Forecast Automation.exe

Notes
-----
- `datas` bundles the two static brand assets (logo + icon) into an
  "assets/" folder at the root of the frozen app, matching what
  gui/branding.resource_path() expects when `sys._MEIPASS` is set.
- `console=False` gives a proper windowed app (no black terminal window
  popping up alongside the GUI).
- The backend modules (config.py, excel_reader.py, aggregator.py,
  comment_mapper.py, historical_lookup.py, summary_writer.py,
  validator.py) are plain local imports from gui/runner.py, so
  PyInstaller's default import analysis picks them up automatically --
  nothing extra needs to be listed for them here.
"""
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hidden_imports = collect_submodules("ttkbootstrap")
# PIL.ImageTk resolves the Tk photo-image bridge via this module at
# runtime; PyInstaller's static import analysis doesn't always catch it
# because the import happens conditionally inside a try/except, so it
# must be listed explicitly or the frozen app crashes the first time it
# tries to render any image (e.g. the header logo, or ttkbootstrap's own
# combobox arrow icon, which is drawn via PIL internally).
hidden_imports += ["PIL._tkinter_finder"]

a = Analysis(
    ["gui_main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("gui/assets/nvish_logo.png", "assets"),
        ("gui/assets/nvish_icon.ico", "assets"),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="NVISH Sales Forecast Automation",
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
    icon="gui/assets/nvish_icon.ico",
)
