"""Driving Excel over COM to turn an ``.xlsx`` into a macro-enabled ``.xlsm``.

This is the only Windows-only part of the export. ``pythoncom`` and ``win32com``
are imported inside :func:`excel_application`, not at module scope, so the package
stays importable — and the rest of it testable — on Linux and macOS.

Bug B12 lived here: the old code started a fresh Excel process *per sheet*, saved
through a chain of ``0data.xlsm``, ``1data.xlsm``, … temp files in the working
directory (deleting ``-1data.xlsm`` on the first pass), and renamed the last one
over the target. Now one Excel session handles the whole workbook inside a
temporary directory, and the result is moved into place with a single
:func:`os.replace`.
"""

from __future__ import annotations

import logging
import os
from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from .button import Button

logger = logging.getLogger(__name__)

XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52
"""``XlFileFormat.xlOpenXMLWorkbookMacroEnabled`` — the .xlsm format."""

VB_COMPONENT_STANDARD_MODULE = 1
"""``vbext_ComponentType.vbext_ct_StdModule``."""


@contextmanager
def excel_application() -> Iterator[Any]:
    """A scoped, invisible Excel instance, quit even if the body raises."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            yield excel
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()


def _close_if_open(excel: Any, path: Path) -> None:
    """Drop a stale copy of the target the user still has open, so it can be replaced."""
    try:
        for workbook in excel.Workbooks:
            if Path(workbook.FullName) == path:
                workbook.Close(SaveChanges=False)
                break
    except Exception:
        logger.debug("Could not check for an open copy of %s", path, exc_info=True)


def inject(source: Path, target: Path, sheets: Mapping[str, Sequence[Button]], replaces: Path | None = None) -> Path:
    """Draw ``sheets``' buttons into ``source`` and save the result as ``target``.

    Every button on a sheet is drawn before any macro source is read, so the
    save/restore geometry each button registers actually reaches the module (B1).
    """
    with excel_application() as excel:
        if replaces is not None:
            _close_if_open(excel, replaces)
        workbook = excel.Workbooks.Open(str(source))
        try:
            for sheet_name, buttons in sheets.items():
                worksheet = workbook.Sheets(sheet_name)
                for button in buttons:
                    button.create(worksheet)
                module = workbook.VBProject.VBComponents.Add(VB_COMPONENT_STANDARD_MODULE)
                module.CodeModule.AddFromString(
                    "\n\n".join(button.macro.code.strip() for button in buttons)
                )
            workbook.SaveAs(str(target), FileFormat=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
        finally:
            workbook.Close(SaveChanges=False)
    logger.info("Injected macro buttons into %d sheet(s)", len(sheets))
    return target


def replace_atomically(built: Path, target: Path) -> Path:
    """Move the finished workbook over the target in one step."""
    os.replace(built, target)
    return target
