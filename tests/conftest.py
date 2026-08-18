"""Shared pytest fixtures."""

import pytest

from pharmparser.utils import Settings


@pytest.fixture
def settings() -> Settings:
    """Default export settings, matching config.json.example."""
    return Settings(
        green="19CF1F",
        red="E81737",
        title="Test",
        fileName="data.xlsx",
        colWidth=50,
        cellWidth=15,
        diffWidth=10,
    )


@pytest.fixture
def price_table() -> dict[str, dict[str, float]]:
    """A small three-pharmacy price table.

    ``Аптека 1`` is the reference pharmacy. The data deliberately covers the
    interesting cases: an item every pharmacy stocks, an item only the reference
    stocks, and an item a competitor stocks but the reference does not.
    """
    return {
        "Аптека 1": {"Аспирин, 100мг": 5.00, "Парацетамол, 500мг": 3.00, "Цитрамон, 10шт": 2.00},
        "Аптека 2": {"Аспирин, 100мг": 6.50, "Парацетамол, 500мг": 2.50},
        "Аптека 3": {"Аспирин, 100мг": 7.00, "Ибупрофен, 200мг": 4.00},
    }
