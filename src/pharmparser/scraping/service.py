"""Fanning out across the pharmacies in a profile."""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Mapping, Sequence

from ..config import PharmacyEntry, RequestConfig
from ..domain import Pharmacy, PriceTable
from .client import ClientSessionFactory, ScrapeError
from .protocols import PriceSource

logger = logging.getLogger(__name__)

DEFAULT_CONCURRENCY = 8
"""Simultaneous in-flight pharmacies.

The old implementation opened one OS *process* per pharmacy via
``multiprocessing.Pool(len(codes))`` for work that is purely I/O bound — and
crashed outright on an empty profile, because ``Pool(0)`` is an error (B4).
"""


class NoPharmaciesError(Exception):
    """Raised when a profile has nothing to scrape."""


async def collect(
    source: PriceSource,
    entries: Sequence[PharmacyEntry],
    *,
    concurrency: int = DEFAULT_CONCURRENCY,
) -> PriceTable:
    """Scrape every entry concurrently and assemble a :class:`PriceTable`.

    Failures are reported together rather than silently swallowed: the previous
    ``ParserEngine.errors`` list was never appended to, so the error path had
    nothing to show (B6).
    """
    scrapable = [entry for entry in entries if entry.is_complete]
    if not scrapable:
        raise NoPharmaciesError(
            "this profile has no pharmacies with both a name and a URL — add one before parsing"
        )

    semaphore = asyncio.Semaphore(max(1, concurrency))

    async def run(entry: PharmacyEntry) -> Mapping[str, float]:
        async with semaphore:
            return await source.prices_for(entry)

    results = await asyncio.gather(*(run(entry) for entry in scrapable), return_exceptions=True)

    failures = [
        f"{entry.name}: {result}"
        for entry, result in zip(scrapable, results, strict=True)
        if isinstance(result, BaseException)
    ]
    if failures:
        raise ScrapeError("Some pharmacies could not be parsed:\n  " + "\n  ".join(failures))

    return PriceTable.build(
        (Pharmacy(id=entry.pharmacy_id, name=entry.name), prices)
        for entry, prices in zip(scrapable, results, strict=True)
        if not isinstance(prices, BaseException)
    )


async def scrape_profile(
    request: RequestConfig,
    entries: Sequence[PharmacyEntry],
    *,
    concurrency: int = DEFAULT_CONCURRENCY,
) -> PriceTable:
    """Open a session and scrape one profile's worth of pharmacies."""
    async with ClientSessionFactory(request) as client:
        return await collect(client, entries, concurrency=concurrency)
