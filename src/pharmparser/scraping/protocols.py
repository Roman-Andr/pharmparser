"""The seam between scraping and the rest of the app."""

from collections.abc import Mapping, Sequence
from typing import Protocol

from ..config import PharmacyEntry


class PriceSource(Protocol):
    """Somewhere prices come from.

    The real implementation talks to tabletka.by; tests substitute a fake, which is
    what makes everything above this layer testable without a network.
    """

    async def prices_for(self, entry: PharmacyEntry) -> Mapping[str, float]:
        """Every price the given pharmacy currently lists."""
        ...


class ProgressCallback(Protocol):
    def __call__(self, completed: int, total: int, entries: Sequence[PharmacyEntry]) -> None: ...
