"""Characterisation tests for DataFormatter.

Pins the exact sheet grid so the phase 1/3 rewrite can be shown to preserve
behaviour, and documents bug B9 (docs/REFACTOR_PLAN.md): the "Нет" sentinel
string is written into the same columns as floats, so price columns are a mixed
type and diff columns silently fall back to 0 rather than being left blank.
"""

from openpyxl import Workbook

from pharmparser.excel.formatters import DataFormatter
from pharmparser.utils import DataType, Settings


def build_rows(settings: Settings, data: DataType, formatting) -> list[list]:
    formatter = DataFormatter(settings, data, list(data.keys()), formatting)
    ws = Workbook().active
    formatter.format(ws)
    return [list(row) for row in ws.iter_rows(values_only=True)]


def absolute(p1: float, p2: float) -> float:
    return p2 - p1


def test_header_row_alternates_pharmacy_and_difference(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table, absolute)
    # Two blank spacer rows precede the header (offset = 2).
    assert rows[2] == ["Название", "Аптека 1", "Аптека 2", "Разница", "Аптека 3", "Разница"]


def test_rows_are_sorted_case_insensitively(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table, absolute)
    assert [row[0] for row in rows[3:]] == [
        "Аспирин, 100мг",
        "Ибупрофен, 200мг",
        "Парацетамол, 500мг",
        "Цитрамон, 10шт",
    ]


def test_differences_are_signed_against_the_reference_pharmacy(
    settings: Settings, price_table: DataType
) -> None:
    rows = build_rows(settings, price_table, absolute)
    # Аспирин: 5.00 reference, 6.50 and 7.00 elsewhere.
    assert rows[3] == ["Аспирин, 100мг", 5.0, 6.5, 1.5, 7.0, 2.0]
    # Парацетамол is cheaper at Аптека 2, so the difference is negative.
    assert rows[5][:4] == ["Парацетамол, 500мг", 3.0, 2.5, -0.5]


def test_percentage_formatting_is_relative(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table, lambda p1, p2: (p2 - p1) / p1 * 100)
    # 5.00 -> 6.50 is +30%.
    assert rows[3][3] == 30.0


def test_missing_items_use_the_sentinel_and_a_zero_difference(
    settings: Settings, price_table: DataType
) -> None:
    """B9: absence is encoded as the string "Нет" with a 0 difference.

    A zero difference is indistinguishable from a genuine price match, and the
    column ends up holding both str and float. Phase 1 replaces this with a real
    optional in the domain model.
    """
    rows = build_rows(settings, price_table, absolute)
    assert rows[4] == ["Ибупрофен, 200мг", "Нет", "Нет", 0, 4.0, 0]
    assert rows[6] == ["Цитрамон, 10шт", 2.0, "Нет", 0, "Нет", 0]


def test_autofilter_spans_the_data_block(settings: Settings, price_table: DataType) -> None:
    formatter = DataFormatter(settings, price_table, list(price_table.keys()), absolute)
    ws = Workbook().active
    formatter.format(ws)
    assert ws.auto_filter.ref == "A3:F7"
