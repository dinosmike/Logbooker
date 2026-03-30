# -*- mode: python ; coding: utf-8 -*-
# Сборка: pyinstaller Logbooker.spec (из каталога проекта).
# Иконки должны быть в datas, иначе в EXE не попадут в _MEIPASS — окно и панель задач без иконки.

from pathlib import Path

_spec_dir = Path(SPECPATH).resolve().parent
_datas = []
for _name in ("iconMY.png", "iconMY.ico", "app_icon.png", "app_icon.ico"):
    _p = _spec_dir / _name
    if _p.is_file():
        _datas.append((str(_p), "."))

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=_datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

_ico = _spec_dir / "iconMY.ico"
if not _ico.is_file():
    _ico = _spec_dir / "app_icon.ico"
_exe_icon = str(_ico) if _ico.is_file() else None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Logbooker",
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
    icon=_exe_icon,
)
