"""Tests for the HTTP client, against a real endpoint on localhost."""

from __future__ import annotations

import aiohttp
import pytest
from pydantic import HttpUrl

from pharmparser.config import PharmacyEntry, RequestConfig
from pharmparser.scraping import PricePage, ScrapeError, TabletkaClient

from ..endpoint import FakeEndpoint
from ..pages import simple_page

URL = "https://tabletka.by/ajax-request/reload-pharmacy-price"
ENDPOINT = URL + "/"
"""What the client actually posts to: tabletka.by 500s without the trailing slash (B17)."""

PAGE_HTML = simple_page("от 5,00 р.")
EXPECTED = {
    "Аспирин, таблетки 100мг, Производитель": 5.00,
    "Цитрамон, таблетки N10, Производитель": 5.00,
}


@pytest.fixture
def request_config(endpoint: FakeEndpoint) -> RequestConfig:
    return RequestConfig(
        url=HttpUrl(endpoint.url),
        headers={"Cookie": "PHPSESSID=abc; lim-result=5000"},
        data={"sort": "name", "_csrf": "token"},
    )


@pytest.fixture
def entry() -> PharmacyEntry:
    return PharmacyEntry(name="Аптека 1", url="https://tabletka.by/pharmacies/111")


async def make_client(
    request_config: RequestConfig, **kwargs
) -> tuple[TabletkaClient, aiohttp.ClientSession]:
    session = aiohttp.ClientSession()
    return TabletkaClient(request_config, session, backoff=0, **kwargs), session


# -- response envelope ---------------------------------------------------------


def test_price_page_reads_the_camel_case_count() -> None:
    page = PricePage.model_validate({"priceCount": 12, "data": "<div/>"})
    assert page.price_count == 12


def test_price_page_tolerates_extra_and_missing_keys() -> None:
    page = PricePage.model_validate({"unexpected": 1})
    assert page.price_count == 0
    assert page.data == ""


# -- fetching ------------------------------------------------------------------


async def test_fetches_and_parses_a_single_page(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.serve("111", PAGE_HTML, price_count=2)
    client, session = await make_client(request_config)
    try:
        assert await client.prices_for(entry) == EXPECTED
    finally:
        await session.close()


async def test_paginates_over_the_reported_count(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.serve("111", PAGE_HTML, price_count=12000)  # 12000 prices at 5000 per page
    client, session = await make_client(request_config)
    try:
        await client.prices_for(entry)
    finally:
        await session.close()
    assert [request.page for request in endpoint.requests] == ["1", "2", "3"]


async def test_uses_the_configured_url(entry: PharmacyEntry, endpoint: FakeEndpoint) -> None:
    """B13: the configured URL used to be ignored for a hardcoded host and path.

    Nothing else is listening on this port, so a request that went anywhere else
    would simply not arrive.
    """
    config = RequestConfig(url=HttpUrl(endpoint.url), headers={"Cookie": "x=1"}, data={})
    endpoint.serve_all("", price_count=0)
    client, session = await make_client(config)
    try:
        await client.prices_for(entry)
    finally:
        await session.close()
    assert endpoint.requests
    assert all(request.path == "/ajax-request/reload-pharmacy-price/" for request in endpoint.requests)


async def test_retries_then_succeeds(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.serve("111", PAGE_HTML, price_count=2)
    endpoint.fail_next(503)
    client, session = await make_client(request_config)
    try:
        assert await client.prices_for(entry) == EXPECTED
    finally:
        await session.close()
    assert len(endpoint.requests) == 2  # the failure, then the page


async def test_gives_up_after_the_retry_budget(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.fail(500)
    client, session = await make_client(request_config, retries=2)
    try:
        with pytest.raises(ScrapeError, match="after 2 attempts"):
            await client.prices_for(entry)
    finally:
        await session.close()
    assert len(endpoint.requests) == 2


async def test_reports_the_pharmacy_in_the_error(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    """B6: failures used to vanish — the errors list was never populated."""
    endpoint.fail(500)
    client, session = await make_client(request_config, retries=1)
    try:
        with pytest.raises(ScrapeError, match="pharmacy 111"):
            await client.prices_for(entry)
    finally:
        await session.close()


async def test_a_single_page_list_costs_exactly_one_request(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    """The page response carries priceCount, so no separate count probe is needed."""
    endpoint.serve_all("", price_count=0)
    client, session = await make_client(request_config)
    try:
        await client.prices_for(entry)
    finally:
        await session.close()
    assert len(endpoint.requests) == 1
    assert "lim-result=5000" in endpoint.requests[0].cookie
    assert endpoint.requests[0].page == "1"



async def test_the_configured_form_fields_are_sent(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.serve_all("", price_count=0)
    client, session = await make_client(request_config)
    try:
        await client.prices_for(entry)
    finally:
        await session.close()
    assert endpoint.requests[0].form["_csrf"] == "token"
    assert endpoint.requests[0].form["sort"] == "name"
    assert endpoint.requests[0].pharmacy_id == "111"
    assert endpoint.requests[0].headers["X-Requested-With"] == "XMLHttpRequest"


# -- endpoint normalisation ----------------------------------------------------


def test_the_endpoint_gains_the_trailing_slash_the_site_requires() -> None:
    """B17.

    The live route carries a mandatory ``/`` suffix: the same POST returns 200 with
    it and 500 without. Real config files carry the URL without, so the client
    normalises while ``config.json`` keeps round-tripping unchanged.
    """
    config = RequestConfig(url=HttpUrl(URL), headers={"Cookie": "x=1"})
    assert str(config.url) == URL
    assert config.endpoint == ENDPOINT


def test_an_already_slashed_endpoint_is_left_alone() -> None:
    config = RequestConfig(url=HttpUrl(ENDPOINT), headers={"Cookie": "x=1"})
    assert config.endpoint == ENDPOINT


def test_a_query_string_stays_after_the_slash() -> None:
    config = RequestConfig(url=HttpUrl("https://example.test/prices?a=1"), headers={"Cookie": "x=1"})
    assert config.endpoint == "https://example.test/prices/?a=1"


@pytest.mark.parametrize("status", [400, 401, 403])
async def test_an_expired_session_fails_fast_and_says_so(
    status: int, request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    """These are what the endpoint answers with once the cookies or _csrf expire.

    Retrying cannot help, and "400 Bad Request" on its own tells the user nothing.
    """
    endpoint.fail(status)
    client, session = await make_client(request_config, retries=3)
    try:
        with pytest.raises(ScrapeError, match="session looks expired"):
            await client.prices_for(entry)
    finally:
        await session.close()
    assert len(endpoint.requests) == 1, "a 4xx must not be retried"


async def test_an_expired_session_is_refreshed_from_the_public_pharmacy_page(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.serve("111", PAGE_HTML, price_count=2)
    endpoint.allow_session_refresh()
    endpoint.fail_next(400)
    client, session = await make_client(request_config)
    try:
        assert await client.prices_for(entry) == EXPECTED
    finally:
        await session.close()

    assert endpoint.session_refreshes == 1
    assert [request.form["_csrf"] for request in endpoint.requests] == ["token", "fresh-token"]


async def test_placeholder_credentials_are_refreshed_before_the_first_post(
    entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    config = RequestConfig(
        url=HttpUrl(endpoint.url),
        headers={"Cookie": "REDACTED"},
        data={"sort": "name", "_csrf": "REDACTED"},
    )
    endpoint.serve("111", PAGE_HTML, price_count=2)
    endpoint.allow_session_refresh()
    client, session = await make_client(config)
    try:
        assert await client.prices_for(entry) == EXPECTED
    finally:
        await session.close()

    assert endpoint.session_refreshes == 1
    assert len(endpoint.requests) == 1
    assert endpoint.requests[0].form["_csrf"] == "fresh-token"
    assert endpoint.requests[0].cookie != "REDACTED"
    assert "lim-result=5000" in endpoint.requests[0].cookie


async def test_a_wrong_pharmacy_url_points_at_the_config(
    request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.fail(404)
    client, session = await make_client(request_config, retries=3)
    try:
        with pytest.raises(ScrapeError, match="check the pharmacy URL"):
            await client.prices_for(entry)
    finally:
        await session.close()
    assert len(endpoint.requests) == 1


@pytest.mark.parametrize("status", [429, 408])
async def test_throttling_and_timeouts_are_still_retried(
    status: int, request_config: RequestConfig, entry: PharmacyEntry, endpoint: FakeEndpoint
) -> None:
    endpoint.serve("111", PAGE_HTML, price_count=2)
    endpoint.fail_next(status)
    client, session = await make_client(request_config)
    try:
        assert await client.prices_for(entry) == EXPECTED
    finally:
        await session.close()
    assert len(endpoint.requests) == 2  # the rejection, then the page
