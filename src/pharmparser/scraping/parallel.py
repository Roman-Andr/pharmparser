"""Running the HTML parse off the event loop.

Measured on a nine-pharmacy profile against the live endpoint: downloading every
pharmacy concurrently takes 4.4s, but parsing costs about 1.2s of CPU each. That
work is synchronous, so it ran one pharmacy at a time inside the event loop and set
the floor for the whole scrape at ~15s. Spread over four cores it is ~9s.

A process pool rather than threads: lxml holds the GIL during XPath, and threads
measured *slower* (0.48x on four cores).

Two things keep the pool from costing more than it saves:

* It is created lazily, on the first page big enough to be worth shipping to a
  worker. A small page parses faster than it pickles, and a profile of small pages
  should not pay for workers at all.
* It is proved with a trivial task before anything depends on it. Python 3.14
  starts workers with ``forkserver`` on Linux, which re-imports ``__main__``, so an
  embedding script with no ``if __name__ == "__main__"`` guard — or a frozen build
  with no ``freeze_support()`` — takes the workers down. Proving it once turns that
  into a single warning instead of one failure per pharmacy, and pays the workers'
  startup cost before the first page arrives.
"""

from __future__ import annotations

import asyncio
import logging
import os
from collections.abc import Iterator, Sequence
from concurrent.futures import BrokenExecutor, ProcessPoolExecutor
from contextlib import contextmanager

from .parser import parse_prices

logger = logging.getLogger(__name__)

MAX_PARSE_WORKERS = 4
"""More workers than this stopped helping and started costing: 8 measured slower
than 4 on a four-core box, and each worker peaks around 400 MB on a big pharmacy."""

MIN_PARALLEL_BYTES = 512 * 1024
"""Below this a pharmacy's pages parse faster than they pickle. Real pages run to
20-25 MB, so this only spares small and synthetic ones the round trip."""

WARMUP_TIMEOUT = 30.0
"""Long enough for a cold ``spawn`` worker to import the package on a slow machine."""

_UNAVAILABLE = "Parse workers unavailable (%s); parsing in-process, which is slower"


def worker_count(pharmacies: int) -> int:
    """How many parse workers are worth starting for this much work."""
    if pharmacies < 2:
        return 0  # nothing to overlap with; the pool would only cost startup
    return max(0, min(pharmacies, os.cpu_count() or 1, MAX_PARSE_WORKERS))


def _ping() -> bool:
    """The warm-up task. Top level so every start method can import it."""
    return True


class ParsePool:
    """Parses a pharmacy's pages, on worker processes once that starts paying off."""

    __slots__ = ("_executor", "_started", "_workers")

    def __init__(self, workers: int) -> None:
        self._workers = workers
        self._executor: ProcessPoolExecutor | None = None
        self._started = False

    @property
    def started(self) -> bool:
        """Whether worker processes are actually running."""
        return self._executor is not None

    def _ensure(self) -> ProcessPoolExecutor | None:
        """Start the workers on first real demand, or give up on them for good."""
        if self._started:
            return self._executor
        self._started = True

        try:
            executor = ProcessPoolExecutor(max_workers=self._workers)
            executor.submit(_ping).result(timeout=WARMUP_TIMEOUT)
        except Exception as error:  # any failure here simply means "no pool"
            logger.warning(_UNAVAILABLE, error)
            return None

        logger.debug("Parsing on %d worker process(es)", self._workers)
        self._executor = executor
        return executor

    async def parse(self, pages: Sequence[str]) -> dict[str, float]:
        if self._workers < 2 or sum(map(len, pages)) < MIN_PARALLEL_BYTES:
            return parse_prices(pages)

        executor = self._ensure()
        if executor is None:
            return parse_prices(pages)

        loop = asyncio.get_running_loop()
        try:
            return await loop.run_in_executor(executor, parse_prices, list(pages))
        except (BrokenExecutor, OSError, EOFError) as error:
            # A worker killed mid-run, or a pipe torn down under us.
            logger.warning(_UNAVAILABLE, error)
            return parse_prices(pages)

    def shutdown(self) -> None:
        if self._executor is not None:
            self._executor.shutdown(wait=False, cancel_futures=True)
            self._executor = None


@contextmanager
def parse_pool(pharmacies: int) -> Iterator[ParsePool]:
    """A parse pool sized for this much work, shut down on the way out."""
    pool = ParsePool(worker_count(pharmacies))
    try:
        yield pool
    finally:
        pool.shutdown()
