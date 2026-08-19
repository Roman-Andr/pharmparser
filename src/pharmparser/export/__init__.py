"""Turning a scraped price table into a workbook.

The layer splits three ways: :mod:`~pharmparser.export.grids` decides what every
sheet contains (pure), :mod:`~pharmparser.export.xlsx_writer` puts that into an
``.xlsx`` with openpyxl (cross-platform), and :mod:`~pharmparser.export.vba` adds
the sort/filter buttons.

That last step used to require driving Excel over COM, which is what kept the
macro-enabled report off Linux and out of CI. It is now ordinary OOXML packaging:
the VBA project is compiled by :mod:`~pharmparser.export.vba.ovba` and stapled on
by :mod:`~pharmparser.export.vba.xlsm`, so the ``.xlsm`` builds anywhere. The COM
injector is kept for Windows users who prefer Excel itself to write the file.
"""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path

from ..config import ExportSettings
from ..domain import PriceTable
from ..platform_ import supports_excel_macros
from .grids import Cell, Grid, build_analysis_grid, build_data_grid, build_grids
from .protocols import Exporter
from .xlsx_writer import build_workbook, render, write_grids

logger = logging.getLogger(__name__)


def write_workbook(settings: ExportSettings, table: PriceTable, path: Path) -> Path:
    """Write the plain ``.xlsx`` — the whole report minus the macro buttons."""
    return write_grids(build_grids(settings, table), path, settings.title)


def export_with_macros(
    settings: ExportSettings,
    table: PriceTable,
    path: Path | None = None,
    use_excel: bool = False,
) -> Path:
    """Write the macro-enabled ``.xlsm``.

    Pure Python by default, so this runs on any platform. Pass ``use_excel`` to go
    through the COM injector instead, which needs Windows, Excel, and the Trust
    Center's "Trust access to the VBA project object model".

    The workbook is built inside a temporary directory next to the target and moved
    over it in one step: the reader either sees the previous report or the new one,
    never a half-written chain of ``0data.xlsm``, ``1data.xlsm``, … left behind in
    the working directory (B12).
    """
    from .vba import buttons_for, replace_atomically

    target = Path(path or settings.macro_file_name).absolute()
    target.parent.mkdir(parents=True, exist_ok=True)
    grids = build_grids(settings, table)
    sheets = {grid.title: buttons for grid in grids if (buttons := buttons_for(grid))}

    # Same directory as the target so the final replace stays on one filesystem.
    with tempfile.TemporaryDirectory(dir=target.parent, prefix=".pharmparser-") as scratch:
        plain = write_grids(grids, Path(scratch) / "report.xlsx", settings.title)
        built = Path(scratch) / "report.xlsm"
        if use_excel:
            from .vba import inject

            inject(plain, built, sheets, replaces=target)
        else:
            _package_macros(plain, built, grids, sheets)
        replace_atomically(built, target)

    logger.info("Wrote %s", target)
    return target


def _package_macros(plain: Path, built: Path, grids, sheets) -> Path:
    """Compile the macros and staple them onto ``plain`` without Excel."""
    from .vba.ovba import build_project
    from .vba.source import MODULE_NAME, module_source
    from .vba.xlsm import ButtonSpec, package

    titles = [grid.title for grid in grids]
    project = build_project({MODULE_NAME: module_source(sheets)}, titles)
    specs = {
        title: [ButtonSpec(button.cell_address, button.caption, button.macro.name) for button in buttons]
        for title, buttons in sheets.items()
    }
    return package(plain, built, project, specs, titles)


class XlsxExporter:
    """Plain ``.xlsx``. Works on every platform and is what CI exercises."""

    suffix = ".xlsx"

    def default_path(self, settings: ExportSettings) -> Path:
        return Path(settings.file_name)

    def export(self, settings: ExportSettings, table: PriceTable, path: Path | None = None) -> Path:
        return write_workbook(settings, table, path or self.default_path(settings))


class MacroExporter:
    """Macro-enabled ``.xlsm`` with the VBA sort/filter buttons.

    Builds anywhere; ``use_excel`` opts into the Windows-only COM injector.
    """

    suffix = ".xlsm"

    def __init__(self, use_excel: bool = False) -> None:
        self.use_excel = use_excel

    def default_path(self, settings: ExportSettings) -> Path:
        return Path(settings.macro_file_name)

    def export(self, settings: ExportSettings, table: PriceTable, path: Path | None = None) -> Path:
        return export_with_macros(
            settings, table, path or self.default_path(settings), use_excel=self.use_excel
        )


def select_exporter(macros: bool = True, use_excel: bool = False) -> Exporter:
    """The backend to use here.

    The macro path no longer depends on the platform, so this only falls back when
    macros were not asked for, or when the caller specifically wanted Excel to do
    the writing and Excel is not available.
    """
    if not macros:
        return XlsxExporter()
    if use_excel and not supports_excel_macros():
        logger.warning(
            "Excel cannot be driven on this machine; building the .xlsm in Python instead"
        )
        return MacroExporter()
    return MacroExporter(use_excel=use_excel)


__all__ = [
    "Cell",
    "Exporter",
    "Grid",
    "MacroExporter",
    "XlsxExporter",
    "build_analysis_grid",
    "build_data_grid",
    "build_grids",
    "build_workbook",
    "export_with_macros",
    "render",
    "select_exporter",
    "write_grids",
    "write_workbook",
]
