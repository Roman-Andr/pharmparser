"""Characterisation tests for AnalysisFormatter.

Bug B2 in docs/REFACTOR_PLAN.md lives here: the "Позиций ниже всех" metric uses
``float('-inf')`` as the missing-item sentinel, so ``price < -inf`` is always
False and any item a competitor does not stock silently drops out of the count.

``test_cheapest_everywhere_ignores_missing_items`` is the failing test that pins
that bug; it is skipped until phase 1 moves the metric into a pure domain
function and fixes it.
"""

import pytest
from openpyxl import Workbook

from pharmparser.excel.formatters import AnalysisFormatter
from pharmparser.utils import DataType, Settings


def build_rows(settings: Settings, data: DataType) -> list[list]:
    formatter = AnalysisFormatter(settings, data, list(data.keys()))
    ws = Workbook().active
    formatter.format(ws)
    return [list(row) for row in ws.iter_rows(values_only=True)]


def test_assortment_counts(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table)
    assert rows[0][0] == "Аптека 1"
    assert rows[1] == ["Асортимент", 3, None, None, None]


def test_competitor_mean_is_arithmetic_mean(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table)
    # Competitors stock 2 and 2 items.
    assert rows[2][1] == pytest.approx(2.0)


def test_competitor_mean_survives_a_single_pharmacy(settings: Settings) -> None:
    """Regression: with no competitors the mean input is empty.

    numpy.mean returned nan here; statistics.mean raises. The formatter must
    yield a plain 0 instead of either.
    """
    rows = build_rows(settings, {"Аптека 1": {"Аспирин, 100мг": 5.00}})
    assert rows[2][1] == 0


def test_unique_positions(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table)
    # Only "Цитрамон, 10шт" is stocked by the reference pharmacy alone.
    assert rows[4] == ["Уникальных позиций", 1, None, None, None]


def test_per_competitor_breakdown(settings: Settings, price_table: DataType) -> None:
    rows = build_rows(settings, price_table)
    assert rows[5] == ["", "Асортимент", "Дороже", "Дешевле", "Уникальных"]
    # Аптека 2: 2 items; аспирин dearer there, парацетамол cheaper; no unique items.
    assert rows[6] == ["Аптека 2", 2, 1, 1, 0]
    # Аптека 3: 2 items; аспирин dearer there; ибупрофен is unique to it.
    assert rows[7] == ["Аптека 3", 2, 1, 0, 1]


@pytest.mark.xfail(
    reason="B2: float('-inf') sentinel makes the metric always 0 when any competitor "
    "lacks the item; fixed in phase 1",
    strict=True,
)
def test_cheapest_everywhere_ignores_missing_items(settings: Settings, price_table: DataType) -> None:
    """"Позиций ниже всех" should count items cheaper than every competitor that stocks them.

    Парацетамол is 3.00 at the reference and 2.50 at Аптека 2, so it does not
    count. Цитрамон is stocked only by the reference, so it counts vacuously.
    Аспирин at 5.00 is cheaper than both 6.50 and 7.00, so it counts. Expected: 2.
    """
    rows = build_rows(settings, price_table)
    assert rows[3] == ["Позиций ниже всех", 2, None, None, None]
