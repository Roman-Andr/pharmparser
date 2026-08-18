"""Turning a price page fragment into prices.

Pure: takes HTML text, returns data. No network, no globals.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass

from bs4 import BeautifulSoup, Tag

logger = logging.getLogger(__name__)

PRICE_SUFFIX = " р."
PRICE_PREFIX = "от "

_RESULT_ROW = "div.tooltip-info-header"
_NAME = ":scope > a"
_FORM = "span.form-title"
_PRICE = "span.price-value"

_NUMBER = re.compile(r"-?\d+(?:[.,]\d+)?")

MAX_ROW_DEPTH = 6
"""How far to climb from a result header when looking for its price."""


@dataclass(frozen=True, slots=True)
class DrugPrice:
    name: str
    price: float


def parse_price(text: str) -> float | None:
    """Read a price out of a cell such as ``"от 12,50 р."``.

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


def _row_for(header: Tag, max_levels: int = MAX_ROW_DEPTH) -> Tag:
    """The smallest ancestor of ``header`` that also holds a price.

    Anchoring on the nearest such ancestor keeps each result's name, form and price
    together no matter how the surrounding markup is nested, which is what B7 was
    really about: the old code selected all three across the whole document and
    zipped them positionally.

    The climb stops before any ancestor that contains a second result header, so a
    row with no price of its own cannot reach up and steal a neighbour's — or a
    promo banner's.
    """
    node: Tag = header
    for _ in range(max_levels):
        parent = node.parent
        if not isinstance(parent, Tag) or len(parent.select(_RESULT_ROW)) > 1:
            break
        node = parent
        if node.select_one(_PRICE) is not None:
            return node
    return node


def parse_page(html: str) -> list[DrugPrice]:
    """Extract every priced result from one page of the price table."""
    soup = BeautifulSoup(html, "lxml")
    headers = soup.select(_RESULT_ROW)
    prices: list[DrugPrice] = []
    skipped = 0

    for header in headers:
        row = _row_for(header)
        name = _text_of(header.select_one(_NAME))
        form = _text_of(row.select_one(_FORM))
        price_text = _text_of(row.select_one(_PRICE))

        if not name or not price_text:
            skipped += 1
            continue

        price = parse_price(price_text)
        if price is None:
            skipped += 1
            continue

        prices.append(DrugPrice(name=f"{name}, {form}" if form else name, price=price))

    # Debug rather than a warning: pages legitimately carry prices outside result
    # rows (promo banners), so a mismatch here is a hint, not a fault.
    logger.debug(
        "Read %d price(s) from %d result row(s); page holds %d price cell(s)",
        len(prices),
        len(headers),
        len(soup.select(_PRICE)),
    )
    if skipped:
        logger.warning("Skipped %d unreadable result row(s) out of %d", skipped, len(headers))
    return prices


def merge(pages: list[list[DrugPrice]]) -> dict[str, float]:
    """Flatten paginated results into one name -> price mapping."""
    merged: dict[str, float] = {}
    for page in pages:
        for entry in page:
            merged[entry.name] = entry.price
    return merged
