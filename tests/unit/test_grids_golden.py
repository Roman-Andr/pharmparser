"""Golden-file check on the built grids.

The grids are the whole report's content and layout, so pinning them as data
catches any unintended change to a sheet — a shifted column, a lost row, a
restyled width — without opening a workbook. Regenerate deliberately, and read
the diff: it is the diff a user would see in Excel.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pharmparser.config import ExportSettings
from pharmparser.domain import PriceTable
from pharmparser.export import Grid, build_grids

GOLDEN = Path(__file__).parent.parent / "fixtures" / "golden_grids.json"


def as_json(grid: Grid) -> dict[str, Any]:
    return {
        "title": grid.title,
        "width": grid.width,
        "header_row": grid.header_row,
        "difference_columns": list(grid.difference_columns),
        "column_widths": {str(key): value for key, value in sorted(grid.column_widths.items())},
        "default_column_width": grid.default_column_width,
        "below_colour": grid.below_colour,
        "above_colour": grid.above_colour,
        "rows": [list(row) for row in grid.rows],
    }


def test_grids_match_the_golden_file(settings: ExportSettings, table: PriceTable) -> None:
    expected = json.loads(GOLDEN.read_text(encoding="utf-8"))
    actual = [as_json(grid) for grid in build_grids(settings, table)]
    assert actual == expected


def test_the_golden_file_covers_every_sheet() -> None:
    expected = json.loads(GOLDEN.read_text(encoding="utf-8"))
    assert [sheet["title"] for sheet in expected] == ["Данные", "Проценты", "Test"]
