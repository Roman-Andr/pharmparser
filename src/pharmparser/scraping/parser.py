"""Turning a price page fragment into prices.

Pure: takes HTML text, returns data. No network, no globals.

The selectors are taken from responses captured from the live endpoint
(``tests/fixtures/live_price_page*.json``). Each result is one ``tr.tr-border``
carrying a name cell, a form cell, a manufacturer cell and a price cell; note that
a single row contains *five* ``div.tooltip-info-header`` elements, one per cell, so
that class alone does not identify a result.

An item is identified by all three of name, form and manufacturer. Dropping the
manufacturer was bug B16: the same drug and pack from two makers sells at two
prices, and keying on name and form alone silently kept whichever came last —
about 1 % of rows on a real pharmacy's list.

Selection goes through compiled lxml XPath rather than BeautifulSoup. On a real
4260-row page that is 10.8s -> 1.5s, and parsing was 93 % of a nine-pharmacy run's
wall clock: the work is synchronous, so it also blocked the event loop and stopped
the other pharmacies' requests from overlapping. lxml holds the GIL during XPath,
so a thread pool measured *slower* (0.48x on four cores) — making the parse itself
cheap is what actually helps.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Sequence
from dataclasses import dataclass

from lxml import html as lxml_html
from lxml.etree import XPath, _Element

from ..domain import Product, ProductPrice, parse_money

logger = logging.getLogger(__name__)

PRICE_SUFFIX = " р."
PRICE_PREFIX = "от "

def _has_class(name: str) -> str:
    """XPath predicate matching one class among many, the way a CSS selector does."""
    return f"contains(concat(' ', normalize-space(@class), ' '), ' {name} ')"


_ROW = XPath(f"//tr[{_has_class('tr-border')}]")
_ROW_FALLBACK = XPath(f"//tr[.//span[{_has_class('price-value')}]]")
_NAME = XPath(f".//td[{_has_class('name')}]//*[{_has_class('tooltip-info-header')}]/a")
_NAME_FALLBACK = XPath(f".//td[{_has_class('name')}]//a")
_FORM = XPath(f".//td[{_has_class('form')}]//span[{_has_class('form-title')}]")
_FORM_FALLBACK = XPath(f".//span[{_has_class('form-title')}]")
_MAKER = XPath(f".//td[{_has_class('produce')}]//*[{_has_class('tooltip-info-header')}]//a")
_MAKER_FALLBACK = XPath(f".//td[{_has_class('produce')}]//*[{_has_class('tooltip-info-header')}]")
_PRICE = XPath(f".//td[{_has_class('price')}]//span[{_has_class('price-value')}]")
_PRICE_FALLBACK = XPath(f".//span[{_has_class('price-value')}]")
_ANY_PRICE = XPath(f"//span[{_has_class('price-value')}]")

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


def _text_of(nodes: list[_Element]) -> str:
    """Text of the first match, or "" when nothing matched."""
    return nodes[0].text_content().strip() if nodes else ""


def item_label(name: str, form: str, maker: str) -> str:
    """How one item is labelled in the report — and keyed for comparison.

    Name, pack and manufacturer, comma-separated, with any missing part left out.
    The manufacturer is what makes the label unique (B16); it goes on the end so
    the report's first column still reads name-first and sorts as it always did.
    """
    return ", ".join(part for part in (name, form, maker) if part)


def _select(row: _Element, selector: XPath, fallback: XPath) -> str:
    """Text of the preferred match, or of a looser one if the classes have drifted."""
    return _text_of(selector(row)) or _text_of(fallback(row))


def _result_rows(root: _Element) -> list[_Element]:
    """The table rows holding results.

    Scoping to a row is what B7 was really about: the old code selected names, forms
    and prices document-wide and zipped the three lists, so one extra or missing cell
    silently paired every later name with the wrong price.
    """
    return _ROW(root) or _ROW_FALLBACK(root)


def parse_page(html: str) -> list[DrugPrice]:
    """Extract every priced result from one page of the price table."""
    if not html.strip():
        return []

    root = lxml_html.fromstring(html)
    rows = _result_rows(root)
    prices: list[DrugPrice] = []
    skipped = 0

    for row in rows:
        name = _select(row, _NAME, _NAME_FALLBACK)
        price_text = _select(row, _PRICE, _PRICE_FALLBACK)

        if not name or not price_text:
            skipped += 1
            continue

        price = parse_price(price_text)
        if price is None:
            skipped += 1
            continue

        form = _select(row, _FORM, _FORM_FALLBACK)
        maker = _select(row, _MAKER, _MAKER_FALLBACK)
        prices.append(DrugPrice(name=item_label(name, form, maker), price=price))

    if skipped:
        logger.warning("Skipped %d unreadable result row(s) out of %d", skipped, len(rows))

    price_cells = len(_ANY_PRICE(root))
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


def parse_product_page(html: str) -> list[ProductPrice]:
    """Structured counterpart of :func:`parse_page` for persistence and reports."""
    if not html.strip():
        return []
    root = lxml_html.fromstring(html)
    result: list[ProductPrice] = []
    for row in _result_rows(root):
        name = _select(row, _NAME, _NAME_FALLBACK)
        price_text = _select(row, _PRICE, _PRICE_FALLBACK)
        if not name or not price_text:
            continue
        match = _NUMBER.search(price_text.replace("\xa0", " "))
        if match is None:
            continue
        form = _select(row, _FORM, _FORM_FALLBACK)
        maker = _select(row, _MAKER, _MAKER_FALLBACK)
        result.append(
            ProductPrice(
                Product(name=name, form=form, manufacturer=maker),
                parse_money(match.group().replace(",", ".")),
            )
        )
    return result


def parse_product_prices(pages: Sequence[str]) -> list[ProductPrice]:
    """Merge pages while refusing normalized-key collisions and silent overwrite."""
    merged: dict[str, ProductPrice] = {}
    for page in pages:
        for entry in parse_product_page(page):
            previous = merged.get(entry.product.key)
            if previous is not None and previous.product != entry.product:
                raise ValueError(
                    f"product key collision: {previous.product.label!r} / {entry.product.label!r}"
                )
            merged[entry.product.key] = entry
    return list(merged.values())


def parse_prices(pages: Sequence[str]) -> dict[str, float]:
    """Parse every page of one pharmacy and merge them into label -> price.

    One top-level call doing all of a pharmacy's work, because this is what gets
    handed to a worker process: sending the pages over once and getting a small
    mapping back keeps the inter-process traffic to a minimum.
    """
    return merge([parse_page(page) for page in pages])


def merge(pages: list[list[DrugPrice]]) -> dict[str, float]:
    """Flatten paginated results into one label -> price mapping.

    Labels include the manufacturer, so rows no longer collapse into one another
    (B16). Measured on three real pharmacies' full price lists: 4240, 3547 and 5000
    rows in, and exactly as many out.
    """
    merged: dict[str, float] = {}
    for page in pages:
        for entry in page:
            merged[entry.name] = entry.price
    return merged
