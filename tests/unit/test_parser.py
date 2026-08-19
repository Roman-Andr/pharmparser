"""Tests for the pure HTML parser, against responses captured from tabletka.by."""

import json
from pathlib import Path

import pytest
from bs4 import BeautifulSoup

from pharmparser.scraping import DrugPrice, merge, parse_page, parse_price

from ..pages import page, row, simple_page

FIXTURES = Path(__file__).parent.parent / "fixtures"
LIVE_PAGES = ["live_price_page.json", "live_price_page_2.json", "live_price_page_3.json"]


def envelope(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def html_of(name: str) -> str:
    return envelope(name)["data"]


def fixture(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


# -- price cleanup -------------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("10.17 р.", 10.17),   # the shape the live site actually uses
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


# -- real captured pages -------------------------------------------------------


@pytest.mark.parametrize("name", LIVE_PAGES)
def test_every_row_of_a_real_page_is_read(name: str) -> None:
    """The ten rows a ``lim-result=10`` probe returns all parse.

    This is the check the synthetic fixtures could not make. Against the real
    markup the previous selectors read *zero* prices: a result row carries five
    ``div.tooltip-info-header`` elements, so anchoring on that class found the
    name, form, manufacturer and booking cells rather than the results.
    """
    html = html_of(name)
    prices = parse_page(html)
    assert len(prices) == len(BeautifulSoup(html, "lxml").select("tr.tr-border")) == 10
    assert all(price.price > 0 for price in prices)
    assert all(price.name.count(", ") >= 2 for price in prices), "name, form and maker"


def test_a_real_page_parses_to_the_expected_values() -> None:
    prices = parse_page(html_of("live_price_page.json"))
    assert prices[0] == DrugPrice(
        name="9 Месяцев Фолиевая кислота, таблетки покрытые оболочкой 400мкг N30, Валента",
        price=10.17,
    )
    assert prices[1] == DrugPrice(
        name="911 Дегтярное жидкое мыло, жидкое мыло 250мл N1, Твинс Тэк ЗАО",
        price=6.84,
    )
    assert prices[2] == DrugPrice(
        name="911 Теймурова паста, паста 50мл N1, Твинс Тэк ЗАО", price=4.36
    )


@pytest.mark.parametrize("name", LIVE_PAGES)
def test_the_manufacturer_makes_every_label_unique(name: str) -> None:
    """Regression for B16.

    The same drug and pack from two makers sells at two prices. Keying on name and
    form alone silently kept whichever row came last — about 1 % of a real
    pharmacy's list. The manufacturer goes on the end of the label, so the report's
    first column still reads name-first and sorts as it always did.
    """
    labels = [price.name for price in parse_page(html_of(name))]
    assert len(set(labels)) == len(labels)


@pytest.mark.parametrize("name", LIVE_PAGES)
def test_the_envelope_carries_a_count_and_a_status(name: str) -> None:
    body = envelope(name)
    assert body["status"] == 1
    assert body["priceCount"] > len(parse_page(body["data"]))


def test_the_test_markup_matches_the_captured_markup() -> None:
    """The shared template in tests/pages.py must not drift from the real site.

    Everything else in the suite fakes pages with that template, so if it stops
    matching what the site sends, those tests would go on passing over markup no
    parser could read.
    """
    live = BeautifulSoup(html_of("live_price_page.json"), "lxml").select("tr.tr-border")[0]
    fake = BeautifulSoup(simple_page(), "lxml").select("tr.tr-border")[0]
    for selector in (
        "td.name .tooltip-info-header > a",
        "td.form span.form-title",
        "td.price span.price-value",
    ):
        assert live.select_one(selector) is not None, selector
        assert fake.select_one(selector) is not None, selector
    assert len(fake.select("div.tooltip-info-header")) >= 3


# -- page parsing --------------------------------------------------------------


def test_parses_every_row_with_name_form_and_maker() -> None:
    assert parse_page(simple_page("5,00 р.")) == [
        DrugPrice(name="Аспирин, таблетки 100мг, Производитель", price=5.00),
        DrugPrice(name="Цитрамон, таблетки N10, Производитель", price=5.00),
    ]


def test_rows_stay_aligned_when_the_page_drifts() -> None:
    """The B7 regression, over real markup that has been damaged deliberately.

    ``price_page_drifted.html`` is four captured rows plus a promo banner: the
    banner carries a price but no result row, one row has lost its price, one its
    form title and one its name link. The old code selected names, forms and prices
    document-wide and zipped them, so every later name took the wrong price.
    """
    assert parse_page(fixture("price_page_drifted.html")) == [
        DrugPrice(
            name="9 Месяцев Фолиевая кислота, таблетки покрытые оболочкой 400мкг N30, Валента",
            price=10.17,
        ),
        DrugPrice(name="911 Теймурова паста, Твинс Тэк ЗАО", price=4.36),
    ]


def test_a_promo_price_outside_a_result_row_is_ignored() -> None:
    banner = '<div class="promo-banner"><span class="price-value">99,99 р.</span></div>'
    assert 99.99 not in [price.price for price in parse_page(banner + simple_page())]


def test_empty_page_is_not_an_error() -> None:
    assert parse_page("") == []
    assert parse_page("<div class='search-result'></div>") == []


def test_unreadable_rows_are_logged(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level("WARNING"):
        parse_page(fixture("price_page_drifted.html"))
    assert "Skipped" in caplog.text


def test_a_page_of_prices_that_reads_as_empty_is_an_error(caplog: pytest.LogCaptureFixture) -> None:
    """The signal that would have caught this class of bug immediately."""
    with caplog.at_level("ERROR"):
        assert parse_page('<div><span class="price-value">10.17 р.</span></div>') == []
    assert "markup has probably changed" in caplog.text


def test_a_row_without_a_form_title_still_yields_a_price() -> None:
    assert parse_page(page(row("Аспирин", "", "5,00 р.", maker=""))) == [
        DrugPrice(name="Аспирин", price=5.00)
    ]


# -- merging -------------------------------------------------------------------


def test_merge_flattens_pages() -> None:
    pages = [
        [DrugPrice(name="A", price=1.0), DrugPrice(name="B", price=2.0)],
        [DrugPrice(name="C", price=3.0)],
    ]
    assert merge(pages) == {"A": 1.0, "B": 2.0, "C": 3.0}


def test_merge_keeps_the_last_price_for_a_repeated_label() -> None:
    """Only a genuinely identical item — same name, pack and maker — merges now."""
    assert merge([[DrugPrice(name="A", price=1.0)], [DrugPrice(name="A", price=9.0)]]) == {
        "A": 9.0
    }


def test_two_makers_of_one_drug_stay_apart() -> None:
    """Regression for B16, at the label level."""
    from pharmparser.scraping.parser import item_label

    first = item_label("Амлодипин", "таблетки 10мг N30", "Борисовский ЗМП")
    second = item_label("Амлодипин", "таблетки 10мг N30", "Тева")
    assert first != second
    assert first.startswith("Амлодипин, таблетки 10мг N30"), "the name still leads"
    assert merge(
        [[DrugPrice(name=first, price=1.0), DrugPrice(name=second, price=9.0)]]
    ) == {first: 1.0, second: 9.0}


def test_a_label_leaves_out_the_parts_it_does_not_have() -> None:
    from pharmparser.scraping.parser import item_label

    assert item_label("Аспирин", "", "") == "Аспирин"
    assert item_label("Аспирин", "таблетки", "") == "Аспирин, таблетки"
    assert item_label("Аспирин", "", "Bayer") == "Аспирин, Bayer"


def test_merge_of_nothing_is_empty() -> None:
    assert merge([]) == {}
