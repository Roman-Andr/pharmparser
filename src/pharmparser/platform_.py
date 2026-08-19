"""Platform capability probes, so the rest of the code need not test sys.platform."""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger(__name__)


def is_windows() -> bool:
    return sys.platform == "win32"


def supports_excel_macros() -> bool:
    """Whether Excel itself can be driven over COM here.

    This is *not* what the macro buttons need any more — the ``.xlsm`` is built in
    Python and works everywhere. It only reports whether the opt-in ``--use-excel``
    path is available, which means Windows with pywin32 installed.
    """
    if not is_windows():
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        logger.warning("pywin32 is not installed; Excel cannot be driven over COM here")
        return False
    return True


def open_file(path: Path) -> None:
    """Open a file with the system's default application."""
    if is_windows():
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)
