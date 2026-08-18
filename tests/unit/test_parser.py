"""Tests for the pure HTML parser."""

from pathlib import Path

import pytest

from pharmparser.scraping import DrugPrice, merge, parse_page, parse_price

FIXTURES = Path(__file__).parent.parent / "fixtures"


def fixture(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


# -- price cleanup -------------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("от 5,00 р.", 5.00),
        ("3.50 р.", 3.50),
        ("от 2 р.", 2.0),
        ("  12,75 р.  ", 12.75),
        ("120", 120.0),
    ],
)
def test_parse_price_reads_the_number(text: str, expected: float) -> None:
    assert parse_price(text) == expected


def test_parse_price_strips_affixes_not_character_sets() -> None:
    """B8: rstrip(" р.")/lstrip("от ") removed *any* of those characters.

    A price whose digits are adjacent to a stripped character used to lose them.
    """
    assert parse_price("от 0,50 р.") == 0.50


def test_parse_price_returns_none_for_junk() -> None:
    assert parse_price("нет в наличии") is None
    assert parse_price("") is None


# -- page parsing --------------------------------------------------------------


def test_parses_every_row_with_name_and_form() -> None:
    assert parse_page(fixture("price_page.html")) == [
        DrugPrice("Аспирин, таблетки 100мг", 5.00),
        DrugPrice("Парацетамол, таблетки 500мг", 3.50),
        DrugPrice("Цитрамон, 10шт", 2.00),
    ]


def test_rows_stay_aligned_when_the_page_drifts() -> None:
    """The B7 regression.

    This page has a promo block carrying a price but no result header, a row with no
    form title, and a row with no price at all. The old code selected names, forms and
    prices document-wide and zipped them, so every later name took the wrong price.
    Each row is now read on its own.
    """
    assert parse_page(fixture("price_page_drifted.html")) == [
        DrugPrice("Аспирин, таблетки 100мг", 5.00),
        DrugPrice("Без формы", 7.25),
    ]


def test_empty_page_is_not_an_error() -> None:
    assert parse_page("") == []
    assert parse_page("<div class='search-result'></div>") == []


def test_unreadable_rows_are_logged(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level("WARNING"):
        parse_page(fixture("price_page_drifted.html"))
    assert "Skipped" in caplog.text


# -- merging -------------------------------------------------------------------


def test_merge_flattens_pages() -> None:
    pages = [[DrugPrice("A", 1.0), DrugPrice("B", 2.0)], [DrugPrice("C", 3.0)]]
    assert merge(pages) == {"A": 1.0, "B": 2.0, "C": 3.0}


def test_merge_keeps_the_last_price_for_a_repeated_name() -> None:
    assert merge([[DrugPrice("A", 1.0)], [DrugPrice("A", 9.0)]]) == {"A": 9.0}


def test_merge_of_nothing_is_empty() -> None:
    assert merge([]) == {}
