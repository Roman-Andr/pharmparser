# Parser fixtures

## Captured from tabletka.by

`live_price_page.json`, `live_price_page_2.json` and `live_price_page_3.json` are
verbatim responses from `POST /ajax-request/reload-pharmacy-price/`, captured on
2026-08-18 for pharmacies 3563, 381 and 1953 with `lim-result=10`. They hold the
server's whole JSON envelope — `status`, `priceCount` and the `data` HTML fragment —
and nothing else: no cookies, no CSRF token, no account details.

They are the parser's ground truth. Capturing them found that the parser read
**zero** prices from real pages: a result row is a `tr.tr-border` whose name, form,
manufacturer, booking and price cells *each* contain a `div.tooltip-info-header`, so
anchoring on that class alone never finds a result.

`price_page_drifted.html` is built from four of those captured rows with damage
introduced deliberately — a promo banner carrying a price outside any row, a row
with no price, a row with no form title and a row with no name — to pin the
parser's tolerance of markup drift.

## Synthetic markup elsewhere

Tests that need a page rather than a fixture build one with `tests/pages.py`, whose
template mirrors the captured markup. `test_the_test_markup_matches_the_captured_markup`
keeps the two in step, so the template cannot quietly drift away from what the site
actually sends.

## Refreshing a capture

```bash
curl -s --data 'sort=name&sort_type=asc&str=&_csrf=<token>&id=3563&page=0' \
  -H 'Cookie: <your session cookie, with lim-result=10>' \
  -H 'X-Requested-With: XMLHttpRequest' \
  https://tabletka.by/ajax-request/reload-pharmacy-price/ | python -m json.tool > live_price_page.json
```

The trailing slash is required — without it the endpoint returns HTTP 500. Keep the
response body verbatim, and keep your cookie and CSRF token out of the repository.

## Other fixtures

- `real_world_config.json` — a real user's `config.json` with the session cookie and
  CSRF token redacted.
- `golden_grids.json` — the three exported sheets as data, pinned by
  `tests/unit/test_grids_golden.py`.
