"""Price-page markup for tests, shaped like the live site's.

Mirrors the structure of ``tests/fixtures/live_price_page.json``, which was
captured from tabletka.by: each result is one ``tr.tr-border`` whose name, form,
manufacturer and price each sit in their own ``div.tooltip-info-header``. That
last detail matters — five of those per row is why selecting on the class alone
finds no results at all.

``test_parser.py`` asserts this template still matches the captured page, so a
change on the site shows up as a failing test rather than as an empty report.
"""

from __future__ import annotations

ROW = """\
<tr class="tr-border">
    <td class="name tooltip-info">
        <div class="content-table">
            <div class="text-wrap">
                <div class="tooltip-info-header">
                    <a href="/result/?ls={ls}&region=36">{name}</a>
                </div>
            </div>
        </div>
    </td>
    <td class="form tooltip-info">
        <div class="content-table">
            <div class="tooltip-info-header">
                <div class="text-wrap"><span class="form-title">{form}</span></div>
            </div>
        </div>
    </td>
    <td class="produce tooltip-info">
        <div class="content-table">
            <div class="tooltip-info-header"><span><a href="/search/mnf/?mnf_id=1">{maker}</a></span></div>
        </div>
    </td>
    <td class="price tooltip-info">
        <div class="tooltip-info-header"><span class="price-value">{price}</span></div>
    </td>
</tr>
"""

TABLE = """\
<table class="table-border">
    <thead>
        <tr><th class="name">Наименование</th><th class="form">Форма</th>
            <th class="produce">Производитель</th><th class="price">Цены</th></tr>
    </thead>
    <tbody>
{rows}    </tbody>
</table>
"""


def row(name: str, form: str, price: str, *, ls: int = 1, maker: str = "Производитель") -> str:
    """One result row, exactly as the site renders it."""
    return ROW.format(name=name, form=form, price=price, ls=ls, maker=maker)


def page(*rows: str) -> str:
    """A results table wrapping the given rows."""
    return TABLE.format(rows="".join(rows))


def simple_page(price: str = "5,00 р.") -> str:
    """Two priced results — the default stand-in for a scraped page."""
    return page(
        row("Аспирин", "таблетки 100мг", price, ls=1),
        row("Цитрамон", "таблетки N10", price, ls=2),
    )
