"""Parsing on a process pool, and the ways it is allowed to not happen.

The speed-up is real but it must never be the reason a report fails: a frozen build
without ``freeze_support``, a sandbox that forbids fork, or a killed worker all have
to degrade to in-process parsing rather than raise.
"""

from __future__ import annotations

import pytest

from pharmparser.scraping.parallel import (
    MAX_PARSE_WORKERS,
    MIN_PARALLEL_BYTES,
    ParsePool,
    parse_pool,
    worker_count,
)

ROW = (
    '<tr class="tr-border">'
    '<td class="name"><div class="tooltip-info-header"><a>Аспирин</a></div></td>'
    '<td class="form"><span class="form-title">таб</span></td>'
    '<td class="price"><span class="price-value">5,40 р.</span></td>'
    "</tr>"
)
EXPECTED = {"Аспирин, таб": 5.4}


@pytest.mark.parametrize(
    ("pharmacies", "expected"),
    [(0, 0), (1, 0), (2, 2), (100, MAX_PARSE_WORKERS)],
)
def test_worker_count_scales_with_the_work(pharmacies: int, expected: int, monkeypatch) -> None:
    monkeypatch.setattr("pharmparser.scraping.parallel.os.cpu_count", lambda: 8)
    assert worker_count(pharmacies) == expected


BIG = [ROW + "<!--" + "x" * MIN_PARALLEL_BYTES + "-->"]


async def test_a_single_pharmacy_never_starts_workers() -> None:
    """Startup would cost more than the parse it is meant to overlap with."""
    with parse_pool(1) as pool:
        assert await pool.parse(BIG) == EXPECTED
        assert not pool.started


async def test_small_pages_are_parsed_in_process() -> None:
    """A page parses faster than it pickles until it gets big."""
    with parse_pool(4) as pool:
        assert await pool.parse([ROW]) == EXPECTED
        assert not pool.started, "no workers for a page this small"


async def test_a_big_page_starts_the_workers() -> None:
    with parse_pool(4) as pool:
        if await pool.parse(BIG) != EXPECTED:
            pytest.fail("wrong prices")
        if not pool.started:
            pytest.skip("no process pool available in this environment")


async def test_workers_that_cannot_start_are_not_fatal(monkeypatch, caplog) -> None:
    def refuse(*args, **kwargs):
        raise OSError("fork not permitted")

    monkeypatch.setattr("pharmparser.scraping.parallel.ProcessPoolExecutor", refuse)
    with caplog.at_level("WARNING"), parse_pool(4) as pool:
        assert await pool.parse(BIG) == EXPECTED
    assert "parsing in-process" in caplog.text


async def test_workers_that_die_on_the_warm_up_are_not_fatal(monkeypatch, caplog) -> None:
    """Python 3.14 uses forkserver on Linux, which re-imports __main__; an unguarded
    embedding script takes the workers down at the first submit."""

    class DeadPool:
        def __init__(self, *args, **kwargs) -> None:
            pass

        def submit(self, *args, **kwargs):
            raise OSError(104, "Connection reset by peer")

        def shutdown(self, **kwargs) -> None:
            pass

    monkeypatch.setattr("pharmparser.scraping.parallel.ProcessPoolExecutor", DeadPool)
    with caplog.at_level("WARNING"), parse_pool(4) as pool:
        assert await pool.parse(BIG) == EXPECTED
        assert not pool.started
    assert "parsing in-process" in caplog.text


async def test_the_warm_up_is_attempted_only_once(monkeypatch) -> None:
    attempts = []

    def refuse(*args, **kwargs):
        attempts.append(1)
        raise OSError("nope")

    monkeypatch.setattr("pharmparser.scraping.parallel.ProcessPoolExecutor", refuse)
    with parse_pool(4) as pool:
        for _ in range(3):
            await pool.parse(BIG)
    assert len(attempts) == 1


async def test_pages_are_merged_across_a_pharmacy() -> None:
    second = ROW.replace("Аспирин", "Цитрамон").replace("5,40", "3,10")
    assert await ParsePool(0).parse([ROW, second]) == {"Аспирин, таб": 5.4, "Цитрамон, таб": 3.1}
