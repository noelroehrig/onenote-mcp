# PyInstaller spec for the standalone onenote-mcp.exe.
# Build from the repo root: python -m PyInstaller packaging/onenote-mcp.spec
#
# - onefile + console: one asset to download; stdio transport needs a console exe.
# - mcp metadata: mcp.server.fastmcp calls importlib.metadata.version("mcp") on import.
# - comtypes needs nothing extra: hook-comtypes.client (pyinstaller-hooks-contrib)
#   covers its dynamic imports, and comtypes itself writes generated typelib
#   wrappers to %TEMP%\comtypes_cache when frozen.
# - upx=False: UPX-packed executables trigger more antivirus false positives.

import os

from PyInstaller.utils.hooks import copy_metadata

a = Analysis(
    [os.path.join(SPECPATH, "entry.py")],
    datas=copy_metadata("mcp"),
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name="onenote-mcp",
    console=True,
    upx=False,
)
