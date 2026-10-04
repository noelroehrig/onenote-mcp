"""PyInstaller entry point for the standalone onenote-mcp.exe."""

import os
import sys

from onenote_mcp.server import main

if __name__ == "__main__":
    main()
    # mcp's stdio transport closes sys.stdout's buffer on exit. The PyInstaller
    # bootloader flushes sys.stdout and sys.__stdout__ after this script returns,
    # which would print a "closed file" traceback on every shutdown.
    sys.stdout = sys.__stdout__ = open(os.devnull, "w")
