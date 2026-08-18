"""Turning a price page fragment into prices.

Pure: takes HTML text, returns data. No network, no globals.

The selectors are taken from responses captured from the live endpoint
(``tests/fixtures/live_price_page*.json``). Each result is one ``tr.tr-border``
carrying a name cell, a form cell, a manufacturer cell and a price cell; note that
a single row contains *five* ``div.tooltip-info-header`` elements, one per cell, so
that class alone does not identify a result.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass

from bs4 import BeautifulSoup, Tag

logger = logging.getLogger(__name__)

PRICE_SUFFIX = " р."
PRICE_PREFIX = "от "

_ROW = "tr.tr-border"
_ROW_FALLBACK = "tr"
_NAME = "td.name .tooltip-info-header > a"
_NAME_FALLBACK = "td.name a"
_FORM = "td.form span.form-title"
_FORM_FALLBACK = "span.form-title"
_PRICE = "td.price span.price-value"
_PRICE_FALLBACK = "span.price-value"

_NUMBER = re.compile(r"-?\d+(?:[.,]\d+)?")


@dataclass(frozen=True, slots=True)
class DrugPrice:
    name: str
    price: float


def parse_price(text: str) -> float | None:
    """Read a price out of a cell such as ``"10.17 р."`` or ``"от 12,50 р."``.

    The old implementation used ``.rstrip(" р.").lstrip("от ")``, which strip
    *character sets* rather than affixes and so chewed into the number itself for
    some inputs, then called ``float`` unguarded so one bad cell aborted the whole
    pharmacy (B8).
    """
    cleaned = text.strip().removeprefix(PRICE_PREFIX).removesuffix(PRICE_SUFFIX).strip()
    match = _NUMBER.search(cleaned)
    if match is None:
        logger.warning("Could not read a price from %r", text)
        return None
    return float(match.group().replace(",", "."))


def _text_of(node: Tag | None) -> str:
    return node.text.strip() if node is not None else ""


def _select(row: Tag, selector: str, fallback: str) -> Tag | None:
    """The preferred match, or a looser one if the cell classes have drifted."""
    return row.select_one(selector) or row.select_one(fallback)


def _result_rows(soup: BeautifulSoup) -> list[Tag]:
    """The table rows holding results.

    Scoping to a row is what B7 was really about: the old code selected names, forms
    and prices document-wide and zipped the three lists, so one extra or missing cell
    silently paired every later name with the wrong price.
    """
    rows = soup.select(_ROW)
    if rows:
        return rows
    return [row for row in soup.select(_ROW_FALLBACK) if row.select_one(_PRICE_FALLBACK) is not None]


def parse_page(html: str) -> list[DrugPrice]:
    """Extract every priced result from one page of the price table."""
    soup = BeautifulSoup(html, "lxml")
    rows = _result_rows(soup)
    prices: list[DrugPrice] = []
    skipped = 0

    for row in rows:
        name = _text_of(_select(row, _NAME, _NAME_FALLBACK))
        form = _text_of(_select(row, _FORM, _FORM_FALLBACK))
        price_text = _text_of(_select(row, _PRICE, _PRICE_FALLBACK))

        if not name or not price_text:
            skipped += 1
            continue

        price = parse_price(price_text)
        if price is None:
            skipped += 1
            continue

        prices.append(DrugPrice(name=f"{name}, {form}" if form else name, price=price))

    if skipped:
        logger.warning("Skipped %d unreadable result row(s) out of %d", skipped, len(rows))

    price_cells = len(soup.select(_PRICE_FALLBACK))
    if price_cells and not prices:
        # The page plainly holds prices, so the selectors — not the data — are wrong.
        logger.error(
            "Read no prices from a page holding %d price cell(s); the site's markup has "
            "probably changed. See tests/fixtures/live_price_page.json for the layout "
            "the parser expects.",
            price_cells,
        )
    else:
        logger.debug("Read %d price(s) from %d result row(s)", len(prices), len(rows))
    return prices


def merge(pages: list[list[DrugPrice]]) -> dict[str, float]:
    """Flatten paginated results into one name -> price mapping.

    Rows sharing a name and form — the same drug from different manufacturers — do
    collapse here, keeping the last price seen. See B16 in docs/REFACTOR_PLAN.md.
    """
    merged: dict[str, float] = {}
    for page in pages:
        for entry in page:
            merged[entry.name] = entry.price
    return merged
