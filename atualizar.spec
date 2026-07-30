# -*- mode: python ; coding: utf-8 -*-
# Empacota o AtualizarDashboard (gera dashboard.html + Excel) num unico arquivo.
# Uso: pyinstaller --clean --noconfirm atualizar.spec
import os

# icone opcional a partir do logo (o build_exe.bat gera data/logo.ico)
icone = "data/logo.ico" if os.path.exists("data/logo.ico") else None

a = Analysis(
    ["atualizar.py"],
    pathex=[SPECPATH],                          # acha os modulos locais (forecast_accuracy etc.)
    binaries=[],
    datas=[("dashboard_template.html", ".")],   # template embutido no .exe
    hiddenimports=["forecast_accuracy", "dashboard_build", "openpyxl.cell._writer"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "tkinter", "PyQt5", "PySide2", "scipy", "IPython"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="AtualizarDashboard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,        # mostra a reconciliacao (DIF 0 OK) ao rodar
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icone,
)
