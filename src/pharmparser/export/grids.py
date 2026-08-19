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

ANALYSIS_COLUMNS = 9
"""Columns used by the redesigned analysis dashboard."""


@dataclass(frozen=True, slots=True)
class AnalysisPresentation:
    """Semantic layout metadata for the styled analysis dashboard."""

    merged_ranges: tuple[tuple[int, int, int, int], ...]
    section_rows: tuple[int, ...]
    metric_rows: tuple[int, ...]
    table_header_row: int
    table_first_row: int
    table_last_row: int
    note_rows: tuple[int, ...]


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
    analysis_presentation: AnalysisPresentation | None = None
    """Present only for the dashboard sheet; interpreted by the workbook writer."""

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


def build_data_grid(settings: ExportSettings, table: PriceTable, difference: DifferenceFn, title: str) -> Grid:
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
    """Build the presentation-ready market dashboard."""
    summary = summarise(table)
    breakdown: list[tuple[Cell, ...]] = [
        (
            stats.pharmacy.name,
            stats.assortment,
            stats.shared,
            stats.dearer,
            stats.cheaper,
            stats.equal,
            stats.unique,
            _round_metric(stats.mean_price),
            _round_metric(stats.mean_difference_percent),
        )
        for stats in summary.competitors
    ]
    table_header_row = 14
    table_first_row = table_header_row + 1
    table_last_row = table_header_row + len(breakdown)
    note_row = table_last_row + 2
    return Grid(
        title=settings.title,
        rows=(
            ("АНАЛИЗ ЦЕН И АССОРТИМЕНТА",),
            (
                f"Базовая аптека: {summary.reference.name}",
                None,
                None,
                None,
                None,
                None,
                "Конкурентов",
                None,
                len(summary.competitors),
            ),
            (),
            ("КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ",),
            (
                "Ассортимент",
                summary.assortment,
                "Дешевле всех",
                summary.cheapest_everywhere,
                "Уникальные товары",
                summary.unique_items,
                "Средняя цена, BYN",
                _round_metric(summary.mean_price),
            ),
            (
                "Позиций на рынке",
                summary.market_assortment,
                "Средний ассортимент конкурентов",
                _round_metric(summary.mean_competitor_assortment),
                "Общих с рынком",
                summary.shared_market_items,
                "Только у конкурентов",
                summary.competitor_only_items,
            ),
            (),
            ("ЦЕНОВАЯ ПОЗИЦИЯ",),
            (
                "Сравнимых пар цен",
                summary.comparisons,
                "У нас дешевле",
                summary.reference_cheaper,
                "У нас дороже",
                summary.reference_dearer,
                "Цена совпадает",
                summary.equal_prices,
            ),
            (
                "Средняя разница, BYN",
                _round_metric(summary.mean_difference),
                "Средняя разница, %",
                _round_metric(summary.mean_difference_percent),
                "Доля выгодных сравнений, %",
                _round_metric(summary.advantageous_share),
            ),
            ("Положительная разница означает, что цена конкурента выше.",),
            (),
            ("СРАВНЕНИЕ С КОНКУРЕНТАМИ",),
            (
                "Аптека",
                "Ассортимент",
                "Общие позиции",
                "У нас дешевле",
                "У нас дороже",
                "Цена равна",
                "Только у конкурента",
                "Средняя цена, BYN",
                "Разница, %",
            ),
            *breakdown,
            (),
            (
                "«Дешевле всех» совпадает с Apply Filters: цена есть у всех конкурентов и каждая разница > 0. "
                "«Уникальные товары» есть только у базовой аптеки.",
            ),
        ),
        width=ANALYSIS_COLUMNS,
        column_widths={1: 30, 2: 15, 3: 22, 4: 17, 5: 17, 6: 15, 7: 22, 8: 20, 9: 16},
        default_column_width=17,
        analysis_presentation=AnalysisPresentation(
            merged_ranges=(
                (1, 1, 1, 9),
                (2, 1, 2, 6),
                (2, 7, 2, 8),
                (4, 1, 4, 9),
                (8, 1, 8, 9),
                (11, 1, 11, 9),
                (13, 1, 13, 9),
                (note_row, 1, note_row, 9),
            ),
            section_rows=(4, 8, 13),
            metric_rows=(5, 6, 9, 10),
            table_header_row=table_header_row,
            table_first_row=table_first_row,
            table_last_row=table_last_row,
            note_rows=(11, note_row),
        ),
    )


def _round_metric(value: float) -> float:
    return round(value, 2)


def build_grids(settings: ExportSettings, table: PriceTable) -> list[Grid]:
    """Every sheet of the report, in workbook order."""
    return [
        build_data_grid(settings, table, absolute_difference, DATA_SHEET),
        build_data_grid(settings, table, percentage_difference, PERCENT_SHEET),
        build_analysis_grid(settings, table),
    ]
