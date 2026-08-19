"""Which buttons a data sheet gets. Pure — no Excel needed to decide this."""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from ..grids import Grid
from .button import Button, rgb
from .criteria import FilterCriteria, SortOrder
from .macros import ApplyFiltersMacro, RemoveFiltersMacro, SortMacro

APPLY_COLOUR = rgb(18, 230, 89)
REMOVE_COLOUR = rgb(230, 64, 18)


def buttons_for(grid: Grid) -> list[Button]:
    """Filter buttons in column A, plus an up/down sort pair over each difference column.

    Both live in the two spacer rows the data grid leaves above its header.
    """
    if grid.header_row is None:
        return []
    buttons = [
        Button(
            "A1",
            "Apply Filters",
            ApplyFiltersMacro(grid.width, FilterCriteria.GREATER_THAN_ZERO, grid.title),
            back_color=APPLY_COLOUR,
            fore_color=APPLY_COLOUR,
        ),
        Button(
            "A2",
            "Remove Filters",
            RemoveFiltersMacro(grid.width, grid.title),
            back_color=REMOVE_COLOUR,
            fore_color=REMOVE_COLOUR,
        ),
    ]
    for column in grid.difference_columns:
        letter = get_column_letter(column)
        buttons += [
            Button(f"{letter}1", "↑", SortMacro(letter, grid.width, SortOrder.DESCENDING, grid.title)),
            Button(f"{letter}2", "↓", SortMacro(letter, grid.width, SortOrder.ASCENDING, grid.title)),
        ]
    return buttons
