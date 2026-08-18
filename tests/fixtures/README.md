# Parser fixtures

`price_page.html` and friends are **synthetic**, reconstructed from the CSS selectors
the original `ParserEngine` used (`div.tooltip-info-header > a`, `span.form-title`,
`span.price-value`) rather than captured from tabletka.by.

They pin the parser's contract — row scoping, price cleanup, tolerance of drift —
but they cannot prove the selectors still match the live site. Replacing them with
real captured responses is worthwhile; drop the HTML in here and the tests should
pass unchanged if the structure is as assumed.
