# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build for the desktop app.

    uv sync --all-groups
    uv run pyinstaller pharmparser.spec

It produces two single-file binaries: `pharmparser` (the GUI, no console — the
rotating log file beside it is the only record of a run) and `pharmparser-cli`
(headless, with a console). The CLI target needs no Tk, so running it is a
complete check that the packaged import graph resolves, on any platform.

The React production build is collected as package data. The entry points live in
`packaging/` because PyInstaller runs its entry script as `__main__`.
"""

from PyInstaller.utils.hooks import copy_metadata

# The dist-info goes in so importlib.metadata can still answer __version__ inside
# the frozen binary — which is what the updater compares against GitHub Releases.
metadata = copy_metadata("pharmparser")

datas = [
    *metadata,
    ("config.json.example", "."),
    ("src/pharmparser/web/static", "pharmparser/web/static"),
]

hiddenimports = [
    # Reached only through lazy imports, so static analysis misses them.
    "pharmparser.export.vba.injector",
    "pharmparser.web.api",
    "pharmparser.web.desktop",
    "uvicorn.logging",
    "uvicorn.loops.auto",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets.auto",
]

a = Analysis(
    ["packaging/gui_entry.py"],
    pathex=["src"],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    # Nothing here imports these; excluding them keeps the binary from bloating.
    excludes=["pandas", "numpy", "matplotlib", "pytest", "mypy", "ruff", "tkinter"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="pharmparser",
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
)


# The headless twin. No Tk, so this one is runnable wherever it is built, which is
# what makes "does the packaged app still import?" answerable in CI.
cli_analysis = Analysis(
    ["packaging/cli_entry.py"],
    pathex=["src"],
    binaries=[],
    datas=[*metadata, ("config.json.example", ".")],
    hiddenimports=["pharmparser.export.vba.injector"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["pandas", "numpy", "matplotlib", "pytest", "mypy", "ruff", "tkinter"],
    noarchive=False,
)

cli_exe = EXE(
    PYZ(cli_analysis.pure),
    cli_analysis.scripts,
    cli_analysis.binaries,
    cli_analysis.datas,
    [],
    name="pharmparser-cli",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
