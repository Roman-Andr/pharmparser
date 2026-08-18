"""The one place that knows openpyxl.

Everything it writes comes from a :class:`~pharmparser.export.grids.Grid`, so the
report's content and layout are decided — and tested — before Excel is involved.
"""

from __future__ import annotations

import logging
from collections.abc import Sequence
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting import Rule
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .grids import MIN_STYLED_COLUMNS, Grid

logger = logging.getLogger(__name__)


def _apply_column_widths(grid: Grid, ws: Worksheet) -> None:
    """Widths by 1-based column index, indexed with ``get_column_letter``.

    Bug B10 was doing this over ``string.ascii_uppercase``, which silently stopped
    at column Z — around 13 pharmacies.
    """
    for column in range(1, max(grid.width, MIN_STYLED_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(column)].width = grid.column_widths.get(
            column, grid.default_column_width
        )


def _apply_conditional_formatting(grid: Grid, ws: Worksheet) -> None:
    if grid.first_data_row is None or not (grid.below_colour and grid.above_colour):
        return
    below = DifferentialStyle(fill=PatternFill(bgColor=grid.below_colour))
    above = DifferentialStyle(fill=PatternFill(bgColor=grid.above_colour))
    for column in grid.difference_columns:
        letter = get_column_letter(column)
        cells = f"{letter}{grid.first_data_row}:{letter}{grid.last_row}"
        ws.conditional_formatting.add(cells, Rule("cellIs", operator="lessThan", formula=["0"], dxf=below))
        ws.conditional_formatting.add(cells, Rule("cellIs", operator="greaterThan", formula=["0"], dxf=above))


def render(grid: Grid, ws: Worksheet) -> None:
    """Write one grid into an existing worksheet."""
    _apply_column_widths(grid, ws)
    _apply_conditional_formatting(grid, ws)
    if grid.header_row is not None:
        ws.auto_filter.ref = f"A{grid.header_row}:{get_column_letter(grid.width)}{grid.last_row}"
    for row in grid.rows:
        ws.append(list(row))


def build_workbook(grids: Sequence[Grid], title: str | None = None) -> Workbook:
    """An in-memory workbook holding every grid, one sheet each."""
    workbook = Workbook()
    workbook.remove(workbook.active)
    if title:
        workbook.properties.title = title
    for grid in grids:
        render(grid, workbook.create_sheet(grid.title))
    return workbook


def write_grids(grids: Sequence[Grid], path: Path, title: str | None = None) -> Path:
    """Save every grid to ``path`` as a plain ``.xlsx``.

    Needs neither Excel nor Windows, which is what makes the CLI and the
    integration tests possible.
    """
    build_workbook(grids, title).save(path)
    logger.info("Wrote %s", path)
    return path
