"""Pure sheet builders.

A :class:`Grid` is everything a sheet contains — its cells *and* the layout the
writer needs to style them — described without touching openpyxl. That is the
point: the report's rules used to live inside worksheet-writing loops, so the only
way to test them was to build a workbook. Here they are ordinary functions
returning ordinary data, and :mod:`pharmparser.export.xlsx_writer` is the single
place that knows about Excel.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field

from ..config import DATA_SHEET, PERCENT_SHEET, ExportSettings
from ..domain import DifferenceFn, PriceTable, absolute_difference, comparison_rows, percentage_difference, summarise

Cell = str | float | None
"""One cell's value. ``None`` is written as a genuinely empty cell."""

HEADER_OFFSET = 2
"""Blank spacer rows above the header, left free for the macro buttons."""

NOT_STOCKED = "Нет"
"""Displayed where a pharmacy does not stock an item.

Only a rendering concern: the domain represents absence as ``None`` so it never
travels inside otherwise-numeric data (B9).
"""

MIN_STYLED_COLUMNS = 26
"""Style at least A..Z so sheets keep their familiar width even when narrow.

Bug B10 was the reverse of this: the formatters iterated ``string.ascii_uppercase``
and therefore stopped styling at Z, silently dropping column widths and conditional
formatting past ~13 pharmacies.
"""

BREAKDOWN_COLUMNS = 5
"""Columns used by the per-competitor table on the analysis sheet."""


@dataclass(frozen=True, slots=True)
class Grid:
    """A whole sheet: its cells plus how they should be laid out."""

    title: str
    rows: tuple[tuple[Cell, ...], ...]
    width: int
    """Number of columns actually carrying content."""
    column_widths: Mapping[int, float] = field(default_factory=dict)
    """Explicit widths by 1-based column index; every other column gets the default."""
    default_column_width: float = 15
    header_row: int | None = None
    """1-based row holding the column headers, or ``None`` for a sheet without one.

    Doubles as the marker for "this is a data sheet": those get an autofilter, and
    they are the sheets the VBA sort/filter buttons are attached to.
    """
    difference_columns: tuple[int, ...] = ()
    """1-based indices of the "Разница" columns, which are colour-scaled and sortable."""
    below_colour: str | None = None
    """Fill for a negative difference — the competitor undercuts the reference."""
    above_colour: str | None = None
    """Fill for a positive difference — the competitor is dearer."""

    @property
    def last_row(self) -> int:
        return len(self.rows)

    @property
    def first_data_row(self) -> int | None:
        return None if self.header_row is None else self.header_row + 1


def _difference_columns(width: int) -> tuple[int, ...]:
    return tuple(range(4, width + 1, 2))


def _data_header(table: PriceTable) -> tuple[Cell, ...]:
    header: list[Cell] = ["Название", table.reference.name]
    for competitor in table.competitors:
        header += [competitor.name, "Разница"]
    return tuple(header)


def _data_rows(table: PriceTable, difference: DifferenceFn) -> list[tuple[Cell, ...]]:
    rows: list[tuple[Cell, ...]] = []
    for row in comparison_rows(table, difference):
        cells: list[Cell] = [row.item, row.prices[0] if row.prices[0] is not None else NOT_STOCKED]
        for price, delta in zip(row.prices[1:], row.differences, strict=True):
            cells.append(price if price is not None else NOT_STOCKED)
            # Left blank when undefined, so it stays distinguishable from a
            # genuine difference of zero (B9).
            cells.append(delta)
        rows.append(tuple(cells))
    return rows


def build_data_grid(
    settings: ExportSettings, table: PriceTable, difference: DifferenceFn, title: str
) -> Grid:
    """A price sheet: one row per item, a price and a difference per competitor."""
    # "Название" + the reference price, then a price and a difference per competitor.
    width = 2 + 2 * len(table.competitors)
    differences = _difference_columns(width)
    widths: dict[int, float] = {1: settings.col_width}
    for column in differences:
        widths[column] = settings.diff_width
    return Grid(
        title=title,
        rows=(
            *((),) * HEADER_OFFSET,
            _data_header(table),
            *_data_rows(table, difference),
        ),
        width=width,
        column_widths=widths,
        default_column_width=settings.cell_width,
        header_row=HEADER_OFFSET + 1,
        difference_columns=differences,
        below_colour=settings.red,
        above_colour=settings.green,
    )


def build_analysis_grid(settings: ExportSettings, table: PriceTable) -> Grid:
    """The summary sheet. Named after ``settings.title`` (B15)."""
    summary = summarise(table)
    breakdown: list[tuple[Cell, ...]] = [
        (stats.pharmacy.name, stats.assortment, stats.dearer, stats.cheaper, stats.unique)
        for stats in summary.competitors
    ]
    return Grid(
        title=settings.title,
        rows=(
            (summary.reference.name,),
            ("Асортимент", summary.assortment),
            ("Средний асортимент конкурентов", summary.mean_competitor_assortment),
            ("Позиций ниже всех", summary.cheapest_everywhere),
            ("Уникальных позиций", summary.unique_items),
            ("", "Асортимент", "Дороже", "Дешевле", "Уникальных"),
            *breakdown,
        ),
        width=BREAKDOWN_COLUMNS,
        column_widths={1: settings.col_width},
        default_column_width=settings.cell_width,
    )


def build_grids(settings: ExportSettings, table: PriceTable) -> list[Grid]:
    """Every sheet of the report, in workbook order."""
    return [
        build_data_grid(settings, table, absolute_difference, DATA_SHEET),
        build_data_grid(settings, table, percentage_difference, PERCENT_SHEET),
        build_analysis_grid(settings, table),
    ]
