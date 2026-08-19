"""Core domain model.

Deliberately free of I/O and of every framework the app uses: no openpyxl, no
UI toolkit, no COM, no network. Everything here is directly unit-testable.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Pharmacy:
    """A pharmacy whose prices are being compared.

    ``id`` is the stable identity (the numeric id from the tabletka.by URL) and is
    what the price table is keyed by. ``name`` is only ever a display label, so two
    pharmacies may legitimately share one — see B14 in docs/REFACTOR_PLAN.md, where
    conflating the two misaligned the exported sheet.
    """

    id: str
    name: str


@dataclass(frozen=True, slots=True)
class PriceTable:
    """Prices for a set of pharmacies.

    The first pharmacy is the *reference*: every comparison in
    :mod:`pharmparser.domain.analysis` is expressed relative to it.
    """

    pharmacies: tuple[Pharmacy, ...]
    prices: Mapping[str, Mapping[str, float]]
    """Pharmacy id -> item name -> price."""

    def __post_init__(self) -> None:
        if not self.pharmacies:
            raise ValueError("a price table needs at least one pharmacy")
        ids = [pharmacy.id for pharmacy in self.pharmacies]
        duplicates = {pid for pid in ids if ids.count(pid) > 1}
        if duplicates:
            raise ValueError(f"duplicate pharmacy ids: {sorted(duplicates)}")
        missing = [pid for pid in ids if pid not in self.prices]
        if missing:
            raise ValueError(f"no prices supplied for pharmacy ids: {missing}")

    @classmethod
    def build(cls, entries: Iterable[tuple[Pharmacy, Mapping[str, float]]]) -> PriceTable:
        pharmacies: list[Pharmacy] = []
        prices: dict[str, Mapping[str, float]] = {}
        for pharmacy, item_prices in entries:
            pharmacies.append(pharmacy)
            prices[pharmacy.id] = dict(item_prices)
        return cls(tuple(pharmacies), prices)

    @classmethod
    def from_mapping(cls, names: Sequence[str], data: Mapping[str, Mapping[str, float]]) -> PriceTable:
        """Adapter for the legacy name-keyed structure produced by ``ParserEngine``.

        Because the legacy structure keys prices by display name, two pharmacies
        sharing a name are already indistinguishable by the time the data reaches
        here. Rather than silently dropping one (B14), this refuses the input.
        Phase 2 threads the real pharmacy ids through the scraper and this adapter
        goes away.
        """
        if len(set(names)) != len(names):
            duplicates = sorted({name for name in names if list(names).count(name) > 1})
            raise ValueError(
                "pharmacy names must be unique until the scraper carries real ids "
                f"(duplicates: {duplicates})"
            )
        return cls.build((Pharmacy(id=name, name=name), data.get(name, {})) for name in names)

    @property
    def reference(self) -> Pharmacy:
        """The pharmacy every other one is compared against."""
        return self.pharmacies[0]

    @property
    def competitors(self) -> tuple[Pharmacy, ...]:
        return self.pharmacies[1:]

    def prices_for(self, pharmacy: Pharmacy) -> Mapping[str, float]:
        return self.prices.get(pharmacy.id, {})

    def price_of(self, pharmacy: Pharmacy, item: str) -> float | None:
        """The price of ``item``, or ``None`` when the pharmacy does not stock it.

        ``None`` replaces the old "Нет" sentinel string, which used to travel inside
        otherwise-numeric columns (B9).
        """
        return self.prices_for(pharmacy).get(item)

    def assortment(self, pharmacy: Pharmacy) -> int:
        return len(self.prices_for(pharmacy))

    def item_names(self) -> list[str]:
        """Every item stocked by any pharmacy, sorted case-insensitively."""
        names: set[str] = set()
        for pharmacy in self.pharmacies:
            names.update(self.prices_for(pharmacy))
        return sorted(names, key=str.lower)
