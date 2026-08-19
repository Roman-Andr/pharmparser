"""The contract every export backend satisfies."""

from __future__ import annotations

from pathlib import Path
from typing import Protocol, runtime_checkable

from ..config import ExportSettings
from ..domain import PriceTable


@runtime_checkable
class Exporter(Protocol):
    """Turns a price table into a workbook on disk.

    Two backends implement it — the plain ``.xlsx`` writer, which runs anywhere,
    and the macro-enabled ``.xlsm`` one, which needs Excel. Callers pick with
    :func:`pharmparser.export.select_exporter` and never branch on the platform
    themselves.
    """

    def default_path(self, settings: ExportSettings) -> Path:
        """Where this backend writes when the caller does not say."""
        ...

    def export(self, settings: ExportSettings, table: PriceTable, path: Path | None = None) -> Path:
        """Write the report and return the path actually produced."""
        ...
