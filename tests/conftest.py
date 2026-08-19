"""Shared pytest fixtures."""

from collections.abc import Iterator
from contextlib import contextmanager

import pytest

from pharmparser.config import ExportSettings
from pharmparser.domain import Pharmacy, PriceTable
from pharmparser.export.vba import injector

from .endpoint import FakeEndpoint, running_endpoint
from .fakes import FakeExcel


@pytest.fixture
def settings() -> ExportSettings:
    """Default export settings, matching config.json.example."""
    return ExportSettings(
        green="19CF1F",
        red="E81737",
        title="Test",
        file_name="data.xlsx",
        col_width=50,
        cell_width=15,
        diff_width=10,
    )


@pytest.fixture
def price_table() -> dict[str, dict[str, float]]:
    """A small three-pharmacy price table in the legacy name-keyed shape.

    ``Аптека 1`` is the reference pharmacy. The data deliberately covers the
    interesting cases: an item every pharmacy stocks, an item only the reference
    stocks, and an item a competitor stocks but the reference does not.
    """
    return {
        "Аптека 1": {"Аспирин, 100мг": 5.00, "Парацетамол, 500мг": 3.00, "Цитрамон, 10шт": 2.00},
        "Аптека 2": {"Аспирин, 100мг": 6.50, "Парацетамол, 500мг": 2.50},
        "Аптека 3": {"Аспирин, 100мг": 7.00, "Ибупрофен, 200мг": 4.00},
    }


@pytest.fixture
def table(price_table: dict[str, dict[str, float]]) -> PriceTable:
    """The same data as a domain :class:`PriceTable`."""
    return PriceTable.build(
        (Pharmacy(id=str(i), name=name), prices)
        for i, (name, prices) in enumerate(price_table.items(), start=1)
    )


@pytest.fixture
def excel_sessions(monkeypatch: pytest.MonkeyPatch) -> list[FakeExcel]:
    """Replaces the Excel COM session with a recorder.

    COM is the only Windows-only part of the .xlsm export, so standing in for it
    lets the whole macro path run in CI — and lets a test assert how many Excel
    processes were started (B12).
    """
    sessions: list[FakeExcel] = []

    @contextmanager
    def fake_application() -> Iterator[FakeExcel]:
        excel = FakeExcel()
        sessions.append(excel)
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            yield excel
        finally:
            excel.Quit()

    monkeypatch.setattr(injector, "excel_application", fake_application)
    return sessions


@pytest.fixture
def endpoint() -> Iterator[FakeEndpoint]:
    """A real HTTP price endpoint on localhost. See tests/endpoint.py."""
    yield from running_endpoint()
