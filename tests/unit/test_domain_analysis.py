"""Tests for the pure comparison and analysis rules."""

import pytest

from pharmparser.domain import (
    Pharmacy,
    PriceTable,
    absolute_difference,
    comparison_rows,
    count_cheapest_everywhere,
    count_unique_items,
    percentage_difference,
    summarise,
)


def test_absolute_difference_is_signed_towards_the_competitor() -> None:
    assert absolute_difference(5.0, 6.5) == 1.5
    assert absolute_difference(5.0, 3.0) == -2.0


def test_percentage_difference_is_relative_to_the_reference() -> None:
    assert percentage_difference(5.0, 6.5) == pytest.approx(30.0)
    assert percentage_difference(4.0, 2.0) == pytest.approx(-50.0)


def test_comparison_rows_cover_every_item_in_order(table: PriceTable) -> None:
    rows = comparison_rows(table, absolute_difference)
    assert [row.item for row in rows] == table.item_names()


def test_comparison_row_prices_are_parallel_to_pharmacies(table: PriceTable) -> None:
    aspirin = comparison_rows(table, absolute_difference)[0]
    assert aspirin.item == "Аспирин, 100мг"
    assert aspirin.prices == (5.00, 6.50, 7.00)
    assert aspirin.differences == (1.50, 2.00)


def test_missing_prices_give_an_undefined_difference(table: PriceTable) -> None:
    """B9: absent is None, not 0 — a zero difference means the prices actually matched."""
    rows = {row.item: row for row in comparison_rows(table, absolute_difference)}

    ibuprofen = rows["Ибупрофен, 200мг"]  # only Аптека 3 stocks it
    assert ibuprofen.prices == (None, None, 4.00)
    assert ibuprofen.differences == (None, None)

    paracetamol = rows["Парацетамол, 500мг"]  # Аптека 3 does not stock it
    assert paracetamol.prices == (3.00, 2.50, None)
    assert paracetamol.differences == (-0.50, None)


def test_differences_are_rounded_to_two_decimals() -> None:
    table = PriceTable.build([(Pharmacy("1", "A"), {"x": 3.00}), (Pharmacy("2", "B"), {"x": 4.00})])
    (row,) = comparison_rows(table, percentage_difference)
    assert row.differences == (33.33,)


def test_cheapest_everywhere_ignores_missing_offers_but_not_every_offer(table: PriceTable) -> None:
    """An exclusive item is not also presented as a price win.

    Аспирин (5.00 vs 6.50 and 7.00) is cheapest. Цитрамон has no comparable
    competitor price and belongs only in the separate unique-items metric.
    """
    assert count_cheapest_everywhere(table) == 1


def test_cheapest_everywhere_requires_a_strict_win() -> None:
    table = PriceTable.build([(Pharmacy("1", "A"), {"x": 5.0}), (Pharmacy("2", "B"), {"x": 5.0})])
    assert count_cheapest_everywhere(table) == 0


def test_unique_items_counts_only_the_reference(table: PriceTable) -> None:
    assert count_unique_items(table) == 1  # Цитрамон


def test_a_single_pharmacy_cannot_claim_market_wins_or_unique_items() -> None:
    table = PriceTable.build([(Pharmacy("1", "A"), {"x": 1.0})])

    assert count_cheapest_everywhere(table) == 0
    assert count_unique_items(table) == 0


def test_summary_reports_the_reference_and_competitor_assortments(table: PriceTable) -> None:
    summary = summarise(table)
    assert summary.reference.name == "Аптека 1"
    assert summary.assortment == 3
    assert summary.market_assortment == 4
    assert summary.mean_price == pytest.approx(10 / 3)
    assert summary.mean_competitor_assortment == pytest.approx(2.0)
    assert summary.cheapest_everywhere == 1
    assert summary.unique_items == 1
    assert summary.shared_market_items == 2
    assert summary.competitor_only_items == 1


def test_summary_reports_price_position_from_comparable_pairs(table: PriceTable) -> None:
    summary = summarise(table)

    assert summary.comparisons == 3
    assert summary.reference_cheaper == 2
    assert summary.reference_dearer == 1
    assert summary.equal_prices == 0
    assert summary.mean_difference == pytest.approx(1.0)
    assert summary.mean_difference_percent == pytest.approx((30 + 40 - 16.6666667) / 3)
    assert summary.advantageous_share == pytest.approx(200 / 3)


def test_summary_competitor_breakdown(table: PriceTable) -> None:
    second, third = summarise(table).competitors

    assert second.pharmacy.name == "Аптека 2"
    assert (
        second.assortment,
        second.shared,
        second.dearer,
        second.cheaper,
        second.equal,
        second.unique,
    ) == (2, 2, 1, 1, 0, 0)
    assert second.mean_price == pytest.approx(4.5)
    assert second.mean_difference == pytest.approx(0.5)
    assert second.mean_difference_percent == pytest.approx((30 - 100 / 6) / 2)

    assert third.pharmacy.name == "Аптека 3"
    assert (
        third.assortment,
        third.shared,
        third.dearer,
        third.cheaper,
        third.equal,
        third.unique,
    ) == (2, 1, 1, 0, 0, 1)


def test_equal_prices_are_a_separate_comparison_outcome() -> None:
    table = PriceTable.build(
        [
            (Pharmacy("1", "A"), {"same": 5.0, "only-a": 2.0}),
            (Pharmacy("2", "B"), {"same": 5.0, "only-b": 7.0}),
        ]
    )

    summary = summarise(table)
    assert summary.comparisons == 1
    assert (summary.reference_cheaper, summary.reference_dearer, summary.equal_prices) == (0, 0, 1)
    assert summary.unique_items == 1
    assert summary.competitor_only_items == 1
    assert summary.cheapest_everywhere == 0


def test_summary_survives_a_single_pharmacy() -> None:
    """statistics.mean raises on an empty sequence where numpy returned nan."""
    table = PriceTable.build([(Pharmacy("1", "A"), {"x": 1.0})])
    summary = summarise(table)
    assert summary.mean_competitor_assortment == 0
    assert summary.competitors == ()
    assert summary.cheapest_everywhere == 0
    assert summary.unique_items == 0
