"""Price comparison and market analysis.

Pure functions over :class:`~pharmparser.domain.models.PriceTable`. These carry
the rules that used to be embedded in worksheet-writing loops, which is what made
them untestable.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from statistics import mean

from .models import Pharmacy, PriceTable

DifferenceFn = Callable[[float, float], float]
"""Reference price and competitor price -> the number shown in a "Разница" column."""


def _round2(value: float) -> float:
    """Round to two decimals the way the original formatter did."""
    return float(format(value, ".2f"))


def absolute_difference(reference: float, other: float) -> float:
    """How much dearer the competitor is, in roubles."""
    return other - reference


def percentage_difference(reference: float, other: float) -> float:
    """How much dearer the competitor is, as a percentage of the reference price."""
    return (other - reference) / reference * 100


@dataclass(frozen=True, slots=True)
class ComparisonRow:
    """One item across every pharmacy.

    ``prices`` is parallel to ``PriceTable.pharmacies``; ``differences`` is parallel
    to ``PriceTable.competitors``. ``None`` means "not stocked" — for a difference it
    means the comparison is undefined, which is distinct from a difference of zero
    (B9: the old code wrote 0 for both).
    """

    item: str
    prices: tuple[float | None, ...]
    differences: tuple[float | None, ...]


@dataclass(frozen=True, slots=True)
class CompetitorStats:
    pharmacy: Pharmacy
    assortment: int
    dearer: int
    """Items the reference stocks more cheaply than this competitor."""
    cheaper: int
    """Items this competitor stocks more cheaply than the reference."""
    unique: int
    """Items this competitor stocks and the reference does not."""


@dataclass(frozen=True, slots=True)
class MarketSummary:
    reference: Pharmacy
    assortment: int
    mean_competitor_assortment: float
    cheapest_everywhere: int
    unique_items: int
    competitors: tuple[CompetitorStats, ...]


def comparison_rows(table: PriceTable, difference: DifferenceFn) -> list[ComparisonRow]:
    """One row per item, sorted case-insensitively by item name."""
    reference = table.reference
    rows: list[ComparisonRow] = []
    for item in table.item_names():
        reference_price = table.price_of(reference, item)
        prices = tuple(table.price_of(pharmacy, item) for pharmacy in table.pharmacies)
        differences = tuple(
            None
            if reference_price is None or (competitor_price := table.price_of(competitor, item)) is None
            else _round2(difference(reference_price, competitor_price))
            for competitor in table.competitors
        )
        rows.append(ComparisonRow(item=item, prices=prices, differences=differences))
    return rows


def count_cheapest_everywhere(table: PriceTable) -> int:
    """Items the reference stocks strictly below every competitor that also stocks them.

    Competitors that do not stock the item are ignored rather than counted as
    infinitely cheap — that inversion was bug B2, which silently forced this metric
    to 0 whenever any competitor lacked the item.
    """
    reference = table.reference
    total = 0
    for item, price in table.prices_for(reference).items():
        competitor_prices = [
            competitor_price
            for competitor in table.competitors
            if (competitor_price := table.price_of(competitor, item)) is not None
        ]
        if all(price < competitor_price for competitor_price in competitor_prices):
            total += 1
    return total


def count_unique_items(table: PriceTable) -> int:
    """Items only the reference stocks."""
    reference = table.reference
    return sum(
        1
        for item in table.prices_for(reference)
        if all(table.price_of(competitor, item) is None for competitor in table.competitors)
    )


def competitor_stats(table: PriceTable, competitor: Pharmacy) -> CompetitorStats:
    reference = table.reference
    reference_prices = table.prices_for(reference)
    competitor_prices = table.prices_for(competitor)
    shared = [(price, competitor_prices[item]) for item, price in reference_prices.items() if item in competitor_prices]
    return CompetitorStats(
        pharmacy=competitor,
        assortment=len(competitor_prices),
        dearer=sum(1 for price, other in shared if price < other),
        cheaper=sum(1 for price, other in shared if price > other),
        unique=sum(1 for item in competitor_prices if item not in reference_prices),
    )


def summarise(table: PriceTable) -> MarketSummary:
    """Everything the "Анализ" sheet reports."""
    competitor_sizes = [table.assortment(competitor) for competitor in table.competitors]
    return MarketSummary(
        reference=table.reference,
        assortment=table.assortment(table.reference),
        mean_competitor_assortment=mean(competitor_sizes) if competitor_sizes else 0,
        cheapest_everywhere=count_cheapest_everywhere(table),
        unique_items=count_unique_items(table),
        competitors=tuple(competitor_stats(table, competitor) for competitor in table.competitors),
    )
