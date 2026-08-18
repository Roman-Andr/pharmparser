"""Turning a scraped price table into a workbook.

The layer splits three ways: :mod:`~pharmparser.export.grids` decides what every
sheet contains (pure), :mod:`~pharmparser.export.xlsx_writer` puts that into an
``.xlsx`` with openpyxl (cross-platform), and :mod:`~pharmparser.export.vba` adds
the sort/filter buttons by driving Excel over COM (Windows only, and imported
lazily so the rest stays importable everywhere).
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


def export_with_macros(settings: ExportSettings, table: PriceTable, path: Path | None = None) -> Path:
    """Write the macro-enabled ``.xlsm``. Requires Excel, so Windows only.

    The workbook is built and injected inside a temporary directory next to the
    target, then moved over it in one step: the reader either sees the previous
    report or the new one, never a half-written chain of ``0data.xlsm``,
    ``1data.xlsm``, … left behind in the working directory (B12).
    """
    from .vba import buttons_for, inject, replace_atomically

    target = Path(path or settings.macro_file_name).absolute()
    target.parent.mkdir(parents=True, exist_ok=True)
    grids = build_grids(settings, table)
    sheets = {grid.title: buttons for grid in grids if (buttons := buttons_for(grid))}

    # Same directory as the target so the final replace stays on one filesystem.
    with tempfile.TemporaryDirectory(dir=target.parent, prefix=".pharmparser-") as scratch:
        plain = write_grids(grids, Path(scratch) / "report.xlsx", settings.title)
        built = inject(plain, Path(scratch) / "report.xlsm", sheets, replaces=target)
        replace_atomically(built, target)

    logger.info("Wrote %s", target)
    return target


class XlsxExporter:
    """Plain ``.xlsx``. Works on every platform and is what CI exercises."""

    suffix = ".xlsx"

    def default_path(self, settings: ExportSettings) -> Path:
        return Path(settings.file_name)

    def export(self, settings: ExportSettings, table: PriceTable, path: Path | None = None) -> Path:
        return write_workbook(settings, table, path or self.default_path(settings))


class MacroExporter:
    """Macro-enabled ``.xlsm`` with the VBA sort/filter buttons. Needs Excel."""

    suffix = ".xlsm"

    def default_path(self, settings: ExportSettings) -> Path:
        return Path(settings.macro_file_name)

    def export(self, settings: ExportSettings, table: PriceTable, path: Path | None = None) -> Path:
        return export_with_macros(settings, table, path or self.default_path(settings))


def select_exporter(macros: bool = True) -> Exporter:
    """The best backend available here.

    Falls back to the plain ``.xlsx`` — with a warning rather than a failure — when
    the macro buttons were wanted but Excel cannot be driven on this machine.
    """
    if not macros:
        return XlsxExporter()
    if supports_excel_macros():
        return MacroExporter()
    logger.warning(
        "The macro buttons need Excel driven over COM, which is Windows only; "
        "writing a plain .xlsx instead"
    )
    return XlsxExporter()


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
