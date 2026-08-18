"""Persisting a scraped price table between runs.

Kept out of the domain layer so :mod:`pharmparser.domain` stays free of I/O, and
out of the UI so it can be tested directly. Caches are per profile and carry a
timestamp — the old implementation wrote one shared ``data.json`` unconditionally,
so switching profile silently reused another profile's prices (A7).
"""

from __future__ import annotations

import json
import logging
from datetime import UTC, datetime
from pathlib import Path

from pydantic import BaseModel, Field

from .domain import Pharmacy, PriceTable

logger = logging.getLogger(__name__)

CACHE_VERSION = 1


class CachedPharmacy(BaseModel):
    id: str
    name: str
    prices: dict[str, float] = Field(default_factory=dict)


class CachedTable(BaseModel):
    version: int = CACHE_VERSION
    scraped_at: datetime
    pharmacies: list[CachedPharmacy]

    def to_table(self) -> PriceTable:
        return PriceTable.build(
            (Pharmacy(id=entry.id, name=entry.name), entry.prices) for entry in self.pharmacies
        )

    @classmethod
    def of(cls, table: PriceTable) -> CachedTable:
        return cls(
            scraped_at=datetime.now(UTC),
            pharmacies=[
                CachedPharmacy(id=p.id, name=p.name, prices=dict(table.prices_for(p)))
                for p in table.pharmacies
            ],
        )


def write_table(table: PriceTable, path: Path) -> None:
    path.write_text(CachedTable.of(table).model_dump_json(indent=2), encoding="utf-8")
    logger.info("Cached %d pharmacies to %s", len(table.pharmacies), path)


def read_table(path: Path) -> PriceTable:
    """Load a cached table. Raises on anything unreadable or of the wrong version."""
    cached = CachedTable.model_validate(json.loads(path.read_text(encoding="utf-8")))
    if cached.version != CACHE_VERSION:
        raise ValueError(f"cache at {path} is version {cached.version}, expected {CACHE_VERSION}")
    return cached.to_table()
