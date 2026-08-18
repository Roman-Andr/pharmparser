"""One place that turns a scraped price table into a workbook."""

from __future__ import annotations

import logging
from pathlib import Path

from openpyxl import Workbook

from .config import ExportSettings
from .domain import PriceTable, absolute_difference, percentage_difference
from .excel.formatters import AnalysisFormatter, BaseFormatter, DataFormatter

logger = logging.getLogger(__name__)

DATA_SHEET = "Данные"
PERCENT_SHEET = "Проценты"
ANALYSIS_SHEET = "Анализ"


def build_formatters(settings: ExportSettings, table: PriceTable) -> list[tuple[BaseFormatter, str]]:
    return [
        (DataFormatter(settings, table, absolute_difference), DATA_SHEET),
        (DataFormatter(settings, table, percentage_difference), PERCENT_SHEET),
        (AnalysisFormatter(settings, table), ANALYSIS_SHEET),
    ]


def write_workbook(settings: ExportSettings, table: PriceTable, path: Path) -> Path:
    """Write the plain ``.xlsx``.

    This is the whole report minus the macro buttons, and needs neither Excel nor
    Windows — which is what makes the CLI and the integration tests possible.
    """
    workbook = Workbook()
    workbook.remove(workbook.active)
    for formatter, title in build_formatters(settings, table):
        formatter.format(workbook.create_sheet(title))
    workbook.save(path)
    logger.info("Wrote %s", path)
    return path


def export_with_macros(settings: ExportSettings, table: PriceTable) -> str:
    """Write the macro-enabled ``.xlsm``. Requires Excel, so Windows only."""
    from .excel import Spreadsheet

    return Spreadsheet(table, settings, build_formatters(settings, table)).export()
