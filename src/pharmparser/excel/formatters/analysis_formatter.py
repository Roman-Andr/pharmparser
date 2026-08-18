from openpyxl.worksheet.worksheet import Worksheet

from ...domain import summarise
from .base_formatter import BaseFormatter

Cell = str | float | None

BREAKDOWN_COLUMNS = 5


class AnalysisFormatter(BaseFormatter):
    __slots__ = ()

    def format(self, ws: Worksheet) -> None:
        summary = summarise(self.table)
        self._set_column_widths(
            ws, {1: self.settings.colWidth}, self.settings.cellWidth, BREAKDOWN_COLUMNS
        )
        breakdown: list[list[Cell]] = [
            [stats.pharmacy.name, stats.assortment, stats.dearer, stats.cheaper, stats.unique]
            for stats in summary.competitors
        ]
        grid: list[list[Cell]] = [
            [summary.reference.name],
            ["Асортимент", summary.assortment],
            ["Средний асортимент конкурентов", summary.mean_competitor_assortment],
            ["Позиций ниже всех", summary.cheapest_everywhere],
            ["Уникальных позиций", summary.unique_items],
            ["", "Асортимент", "Дороже", "Дешевле", "Уникальных"],
            *breakdown,
        ]
        for row in grid:
            ws.append(row)
