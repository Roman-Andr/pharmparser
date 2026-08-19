"""Sheet-content tests for the pure grid builders.

The grids carry everything a sheet will contain, so these assert on plain data —
no workbook, no Excel. The metrics themselves are covered in
tests/unit/test_domain_analysis.py; here it is layout that is under test.
"""

import pytest

from pharmparser.config import ExportSettings
from pharmparser.domain import Pharmacy, PriceTable, absolute_difference, percentage_difference
from pharmparser.export.grids import (
    DATA_SHEET,
    HEADER_OFFSET,
    MIN_STYLED_COLUMNS,
    PERCENT_SHEET,
    build_analysis_grid,
    build_data_grid,
    build_grids,
)


def data_rows(settings: ExportSettings, table: PriceTable, difference) -> list[list]:
    return [list(row) for row in build_data_grid(settings, table, difference, DATA_SHEET).rows]


def test_header_row_alternates_pharmacy_and_difference(settings: ExportSettings, table: PriceTable) -> None:
    rows = data_rows(settings, table, absolute_difference)
    assert rows[HEADER_OFFSET] == ["Название", "Аптека 1", "Аптека 2", "Разница", "Аптека 3", "Разница"]


def test_rows_are_sorted_case_insensitively(settings: ExportSettings, table: PriceTable) -> None:
    rows = data_rows(settings, table, absolute_difference)
    assert [row[0] for row in rows[3:]] == [
        "Аспирин, 100мг",
        "Ибупрофен, 200мг",
        "Парацетамол, 500мг",
        "Цитрамон, 10шт",
    ]


def test_differences_are_signed_against_the_reference_pharmacy(settings: ExportSettings, table: PriceTable) -> None:
    rows = data_rows(settings, table, absolute_difference)
    assert rows[3] == ["Аспирин, 100мг", 5.0, 6.5, 1.5, 7.0, 2.0]
    assert rows[5][:4] == ["Парацетамол, 500мг", 3.0, 2.5, -0.5]


def test_percentage_grid_is_relative(settings: ExportSettings, table: PriceTable) -> None:
    assert data_rows(settings, table, percentage_difference)[3][3] == 30.0


def test_missing_items_render_as_a_label_with_a_blank_difference(settings: ExportSettings, table: PriceTable) -> None:
    """Regression for B9.

    A pharmacy that does not stock an item shows "Нет", and its difference column is
    left empty rather than 0 — a 0 would be indistinguishable from prices that
    genuinely match.
    """
    rows = data_rows(settings, table, absolute_difference)
    assert rows[4] == ["Ибупрофен, 200мг", "Нет", "Нет", None, 4.0, None]
    assert rows[6] == ["Цитрамон, 10шт", 2.0, "Нет", None, "Нет", None]


def test_data_grid_layout(settings: ExportSettings, table: PriceTable) -> None:
    grid = build_data_grid(settings, table, absolute_difference, DATA_SHEET)
    assert grid.title == DATA_SHEET
    assert grid.width == 6
    assert grid.header_row == 3
    assert grid.first_data_row == 4
    assert grid.last_row == 7
    assert grid.difference_columns == (4, 6)
    assert grid.column_widths == {1: settings.col_width, 4: settings.diff_width, 6: settings.diff_width}
    assert grid.default_column_width == settings.cell_width


def wide_table(pharmacy_count: int) -> PriceTable:
    return PriceTable.build(
        (Pharmacy(id=str(i), name=f"Аптека {i}"), {"Аспирин": 5.0 + i}) for i in range(pharmacy_count)
    )


@pytest.mark.parametrize("pharmacy_count", [3, 13, 20])
def test_difference_columns_run_past_column_z(pharmacy_count: int, settings: ExportSettings) -> None:
    """Regression for B10.

    The formatters used to walk ``string.ascii_uppercase``, so column widths and the
    conditional formatting stopped at Z — i.e. at roughly 13 pharmacies. The grid
    describes columns by index, so there is no alphabet to run out of.
    """
    grid = build_data_grid(settings, wide_table(pharmacy_count), absolute_difference, DATA_SHEET)
    total_columns = 2 + 2 * (pharmacy_count - 1)
    assert grid.width == total_columns
    assert grid.difference_columns == tuple(range(4, total_columns + 1, 2))
    assert all(grid.column_widths[column] == settings.diff_width for column in grid.difference_columns)


def test_single_pharmacy_grid_has_no_difference_columns(settings: ExportSettings) -> None:
    grid = build_data_grid(settings, wide_table(1), absolute_difference, DATA_SHEET)
    assert grid.width == 2
    assert grid.difference_columns == ()
    assert list(grid.rows[HEADER_OFFSET]) == ["Название", "Аптека 0"]


def test_analysis_grid_reports_the_headline_metrics(settings: ExportSettings, table: PriceTable) -> None:
    rows = [list(row) for row in build_analysis_grid(settings, table).rows]
    assert rows[0] == ["АНАЛИЗ ЦЕН И АССОРТИМЕНТА"]
    assert rows[1] == [
        "Базовая аптека: Аптека 1",
        None,
        None,
        None,
        None,
        None,
        "Конкурентов",
        None,
        2,
    ]
    assert rows[4] == ["Ассортимент", 3, "Средняя цена, BYN", 3.33, "Дешевле всех", 1, "Только у нас", 1]
    assert rows[5] == [
        "Позиций на рынке",
        4,
        "Средний ассортимент конкурентов",
        2,
        "Общих с рынком",
        2,
        "Только у конкурентов",
        1,
    ]
    assert rows[8] == [
        "Сравнимых пар цен",
        3,
        "У нас дешевле",
        2,
        "У нас дороже",
        1,
        "Цена совпадает",
        0,
    ]
    assert rows[9] == [
        "Средняя разница, BYN",
        1,
        "Средняя разница, %",
        17.78,
        "Доля выгодных сравнений, %",
        66.67,
    ]
    assert rows[13] == [
        "Аптека",
        "Ассортимент",
        "Общие позиции",
        "У нас дешевле",
        "У нас дороже",
        "Цена равна",
        "Только у конкурента",
        "Средняя цена, BYN",
        "Разница, %",
    ]
    assert rows[14] == ["Аптека 2", 2, 2, 1, 1, 0, 0, 4.5, 6.67]
    assert rows[15] == ["Аптека 3", 2, 1, 1, 0, 0, 1, 5.5, 40]


def test_analysis_grid_has_no_header_row_and_so_no_autofilter(settings: ExportSettings, table: PriceTable) -> None:
    grid = build_analysis_grid(settings, table)
    assert grid.header_row is None
    assert grid.first_data_row is None
    assert grid.difference_columns == ()
    assert grid.analysis_presentation is not None
    assert grid.analysis_presentation.table_header_row == 14
    assert grid.analysis_presentation.table_first_row == 15
    assert grid.analysis_presentation.table_last_row == 16


def test_analysis_sheet_is_named_after_settings_title(table: PriceTable) -> None:
    """B15: ``settings.title`` was loaded and validated but never read by anything."""
    settings = ExportSettings(title="Сводка")
    assert build_analysis_grid(settings, table).title == "Сводка"
    assert [grid.title for grid in build_grids(settings, table)] == [DATA_SHEET, PERCENT_SHEET, "Сводка"]


def test_min_styled_columns_is_the_old_alphabet_width() -> None:
    assert MIN_STYLED_COLUMNS == 26
