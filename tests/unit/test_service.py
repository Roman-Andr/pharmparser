"""Tests for the concurrent fan-out, using a fake PriceSource."""

import asyncio
from collections.abc import Mapping

import pytest

from pharmparser.config import PharmacyEntry
from pharmparser.scraping import NoPharmaciesError, ScrapeError, collect


class FakeSource:
    """A PriceSource that answers from a dict and records concurrency."""

    def __init__(self, prices: dict[str, dict[str, float]], fail: set[str] | None = None) -> None:
        self._prices = prices
        self._fail = fail or set()
        self.in_flight = 0
        self.peak_in_flight = 0

    async def prices_for(self, entry: PharmacyEntry) -> Mapping[str, float]:
        self.in_flight += 1
        self.peak_in_flight = max(self.peak_in_flight, self.in_flight)
        try:
            await asyncio.sleep(0)
            if entry.name in self._fail:
                raise ScrapeError(f"boom for {entry.name}")
            return self._prices.get(entry.name, {})
        finally:
            self.in_flight -= 1


def make_entries(*names: str) -> list[PharmacyEntry]:
    return [
        PharmacyEntry(name=name, url=f"https://tabletka.by/pharmacies/{i}")
        for i, name in enumerate(names, start=1)
    ]


async def test_builds_a_price_table_in_entry_order() -> None:
    source = FakeSource({"A": {"x": 1.0}, "B": {"x": 2.0}})
    table = await collect(source, make_entries("A", "B"))

    assert [p.name for p in table.pharmacies] == ["A", "B"]
    assert table.reference.name == "A"
    assert table.price_of(table.competitors[0], "x") == 2.0


async def test_pharmacy_ids_come_from_the_url() -> None:
    """B14: identity is the URL id, so display names need not be unique to the domain."""
    table = await collect(FakeSource({}), make_entries("A", "B"))
    assert [p.id for p in table.pharmacies] == ["1", "2"]


async def test_blank_rows_are_skipped_not_fatal() -> None:
    """B4: Pool(len(codes)) raised ValueError on an empty list."""
    entries = [*make_entries("A"), PharmacyEntry()]
    table = await collect(FakeSource({"A": {"x": 1.0}}), entries)
    assert len(table.pharmacies) == 1


async def test_a_profile_with_nothing_to_scrape_is_reported_clearly() -> None:
    with pytest.raises(NoPharmaciesError, match="add one before parsing"):
        await collect(FakeSource({}), [])


async def test_a_profile_of_only_blank_rows_is_reported_clearly() -> None:
    with pytest.raises(NoPharmaciesError):
        await collect(FakeSource({}), [PharmacyEntry(), PharmacyEntry()])


async def test_failures_are_collected_and_named() -> None:
    """B6: every failure is surfaced, not just the first, and each names its pharmacy."""
    source = FakeSource({"A": {"x": 1.0}}, fail={"B", "C"})
    with pytest.raises(ScrapeError) as caught:
        await collect(source, make_entries("A", "B", "C"))

    message = str(caught.value)
    assert "B: boom for B" in message
    assert "C: boom for C" in message


async def test_concurrency_is_bounded() -> None:
    source = FakeSource({})
    await collect(source, make_entries(*[f"P{i}" for i in range(20)]), concurrency=4)
    assert source.peak_in_flight <= 4


async def test_concurrency_of_zero_is_clamped_not_deadlocked() -> None:
    source = FakeSource({})
    table = await collect(source, make_entries("A", "B"), concurrency=0)
    assert len(table.pharmacies) == 2
    assert source.peak_in_flight == 1
