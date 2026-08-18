"""Tests for the per-profile scrape cache."""

from pathlib import Path

import pytest

from pharmparser.cache import read_table, write_table
from pharmparser.config import cache_path
from pharmparser.domain import PriceTable


def test_round_trips_a_price_table(tmp_path: Path, table: PriceTable) -> None:
    path = tmp_path / "cache.json"
    write_table(table, path)
    restored = read_table(path)

    assert [p.name for p in restored.pharmacies] == [p.name for p in table.pharmacies]
    assert [p.id for p in restored.pharmacies] == [p.id for p in table.pharmacies]
    assert restored.prices_for(restored.reference) == table.prices_for(table.reference)


def test_rejects_a_future_cache_version(tmp_path: Path, table: PriceTable) -> None:
    path = tmp_path / "cache.json"
    write_table(table, path)
    path.write_text(path.read_text(encoding="utf-8").replace('"version": 1', '"version": 99'), encoding="utf-8")
    with pytest.raises(ValueError, match="version 99"):
        read_table(path)


def test_rejects_junk(tmp_path: Path) -> None:
    path = tmp_path / "cache.json"
    path.write_text("not json", encoding="utf-8")
    with pytest.raises(ValueError):
        read_table(path)


def test_cache_paths_differ_per_profile() -> None:
    """A7: one shared data.json meant switching profile reused the wrong prices."""
    assert cache_path("Основной") != cache_path("Запасной")


def test_cache_path_is_filesystem_safe() -> None:
    assert "/" not in cache_path("a/b: c").name
