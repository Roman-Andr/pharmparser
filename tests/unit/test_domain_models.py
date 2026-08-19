"""Tests for the pure domain model."""

import pytest

from pharmparser.domain import Pharmacy, PriceTable


def test_build_preserves_pharmacy_order(table: PriceTable) -> None:
    assert [p.name for p in table.pharmacies] == ["Аптека 1", "Аптека 2", "Аптека 3"]
    assert table.reference.name == "Аптека 1"
    assert [p.name for p in table.competitors] == ["Аптека 2", "Аптека 3"]


def test_price_of_returns_none_when_not_stocked(table: PriceTable) -> None:
    assert table.price_of(table.reference, "Аспирин, 100мг") == 5.00
    assert table.price_of(table.reference, "Ибупрофен, 200мг") is None


def test_item_names_are_the_union_sorted_case_insensitively(table: PriceTable) -> None:
    assert table.item_names() == [
        "Аспирин, 100мг",
        "Ибупрофен, 200мг",
        "Парацетамол, 500мг",
        "Цитрамон, 10шт",
    ]


def test_item_names_sorting_ignores_case() -> None:
    table = PriceTable.build(
        [(Pharmacy(id="1", name="A"), {"banana": 1.0, "Apple": 2.0, "cherry": 3.0})]
    )
    assert table.item_names() == ["Apple", "banana", "cherry"]


def test_assortment_counts_stocked_items(table: PriceTable) -> None:
    assert table.assortment(table.reference) == 3
    assert table.assortment(table.competitors[0]) == 2


def test_empty_table_is_rejected() -> None:
    with pytest.raises(ValueError, match="at least one pharmacy"):
        PriceTable(pharmacies=(), prices={})


def test_duplicate_pharmacy_ids_are_rejected() -> None:
    with pytest.raises(ValueError, match="duplicate pharmacy ids"):
        PriceTable.build(
            [(Pharmacy(id="1", name="A"), {}), (Pharmacy(id="1", name="B"), {})]
        )


def test_pharmacies_may_share_a_display_name_when_ids_differ() -> None:
    """B14: identity is the id, so two branches of one chain are representable."""
    table = PriceTable.build(
        [
            (Pharmacy(id="1", name="Аптека"), {"Аспирин": 5.0}),
            (Pharmacy(id="2", name="Аптека"), {"Аспирин": 6.0}),
        ]
    )
    assert table.assortment(table.pharmacies[0]) == 1
    assert table.price_of(table.pharmacies[1], "Аспирин") == 6.0


def test_legacy_adapter_refuses_duplicate_names() -> None:
    """B14: the name-keyed structure cannot represent duplicates, so it must not pretend to."""
    with pytest.raises(ValueError, match="must be unique"):
        PriceTable.from_mapping(["Аптека", "Аптека"], {"Аптека": {"Аспирин": 5.0}})


def test_legacy_adapter_tolerates_a_pharmacy_with_no_prices() -> None:
    table = PriceTable.from_mapping(["A", "B"], {"A": {"x": 1.0}})
    assert table.assortment(table.competitors[0]) == 0


def test_pharmacy_without_a_prices_entry_is_rejected() -> None:
    with pytest.raises(ValueError, match="no prices supplied"):
        PriceTable(pharmacies=(Pharmacy(id="1", name="A"),), prices={})
