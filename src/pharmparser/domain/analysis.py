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
    shared: int
    """Items stocked by both the reference and this competitor."""
    dearer: int
    """Items the reference stocks more cheaply than this competitor."""
    cheaper: int
    """Items this competitor stocks more cheaply than the reference."""
    equal: int
    """Items whose prices are equal."""
    unique: int
    """Items this competitor stocks and the reference does not."""
    mean_price: float
    mean_difference: float
    """Mean competitor minus reference price for shared items, in roubles."""
    mean_difference_percent: float
    """Mean competitor minus reference price for shared items, in percent."""


@dataclass(frozen=True, slots=True)
class MarketSummary:
    reference: Pharmacy
    assortment: int
    market_assortment: int
    mean_price: float
    mean_competitor_assortment: float
    cheapest_everywhere: int
    unique_items: int
    shared_market_items: int
    competitor_only_items: int
    comparisons: int
    reference_cheaper: int
    reference_dearer: int
    equal_prices: int
    mean_difference: float
    mean_difference_percent: float
    advantageous_share: float
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
    """Items that pass ``Apply Filters`` on every displayed difference column.

    Excel combines those filters with AND, so a missing competitor price excludes
    the row. The comparison uses the rounded values written to the workbook to keep
    the dashboard count identical to the visible filtered result.
    """
    return sum(
        1
        for row in comparison_rows(table, absolute_difference)
        if row.differences and all(difference is not None and difference > 0 for difference in row.differences)
    )


def count_unique_items(table: PriceTable) -> int:
    """Items only the reference stocks."""
    if not table.competitors:
        return 0
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
    percentage_differences = [percentage_difference(price, other) for price, other in shared if price]
    return CompetitorStats(
        pharmacy=competitor,
        assortment=len(competitor_prices),
        shared=len(shared),
        dearer=sum(1 for price, other in shared if price < other),
        cheaper=sum(1 for price, other in shared if price > other),
        equal=sum(1 for price, other in shared if price == other),
        unique=sum(1 for item in competitor_prices if item not in reference_prices),
        mean_price=mean(competitor_prices.values()) if competitor_prices else 0,
        mean_difference=mean(other - price for price, other in shared) if shared else 0,
        mean_difference_percent=mean(percentage_differences) if percentage_differences else 0,
    )


def summarise(table: PriceTable) -> MarketSummary:
    """Everything the "Анализ" sheet reports."""
    reference_prices = table.prices_for(table.reference)
    competitor_sizes = [table.assortment(competitor) for competitor in table.competitors]
    stats = tuple(competitor_stats(table, competitor) for competitor in table.competitors)
    competitor_items = {item for competitor in table.competitors for item in table.prices_for(competitor)}
    comparisons = sum(entry.shared for entry in stats)
    reference_cheaper = sum(entry.dearer for entry in stats)
    reference_dearer = sum(entry.cheaper for entry in stats)
    equal_prices = sum(entry.equal for entry in stats)
    differences = [
        other - price
        for competitor in table.competitors
        for item, price in reference_prices.items()
        if (other := table.price_of(competitor, item)) is not None
    ]
    percentage_differences = [
        percentage_difference(price, other)
        for competitor in table.competitors
        for item, price in reference_prices.items()
        if price and (other := table.price_of(competitor, item)) is not None
    ]
    return MarketSummary(
        reference=table.reference,
        assortment=table.assortment(table.reference),
        market_assortment=len(set(reference_prices) | competitor_items),
        mean_price=mean(reference_prices.values()) if reference_prices else 0,
        mean_competitor_assortment=mean(competitor_sizes) if competitor_sizes else 0,
        cheapest_everywhere=count_cheapest_everywhere(table),
        unique_items=count_unique_items(table),
        shared_market_items=len(set(reference_prices) & competitor_items),
        competitor_only_items=len(competitor_items - set(reference_prices)),
        comparisons=comparisons,
        reference_cheaper=reference_cheaper,
        reference_dearer=reference_dearer,
        equal_prices=equal_prices,
        mean_difference=mean(differences) if differences else 0,
        mean_difference_percent=mean(percentage_differences) if percentage_differences else 0,
        advantageous_share=reference_cheaper / comparisons * 100 if comparisons else 0,
        competitors=stats,
    )
