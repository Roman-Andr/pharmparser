"""Tests for the HTTP client, with the network faked out."""

import aiohttp
import pytest
from aioresponses import aioresponses
from pydantic import HttpUrl

from pharmparser.config import PharmacyEntry, RequestConfig
from pharmparser.scraping import PricePage, ScrapeError, TabletkaClient

from ..pages import simple_page

URL = "https://tabletka.by/ajax-request/reload-pharmacy-price"
ENDPOINT = URL + "/"
"""What the client actually posts to: tabletka.by 500s without the trailing slash (B17)."""

PAGE_HTML = simple_page("от 5,00 р.")


@pytest.fixture
def request_config() -> RequestConfig:
    return RequestConfig(
        url=HttpUrl(URL),
        headers={"Cookie": "PHPSESSID=abc; lim-result=5000"},
        data={"sort": "name", "_csrf": "token"},
    )


@pytest.fixture
def entry() -> PharmacyEntry:
    return PharmacyEntry(name="Аптека 1", url="https://tabletka.by/pharmacies/111")


async def make_client(request_config: RequestConfig, **kwargs) -> tuple[TabletkaClient, aiohttp.ClientSession]:
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


async def test_fetches_and_parses_a_single_page(request_config: RequestConfig, entry: PharmacyEntry) -> None:
    client, session = await make_client(request_config)
    try:
        with aioresponses() as mocked:
            mocked.post(ENDPOINT, payload={"priceCount": 1, "data": ""})
            mocked.post(ENDPOINT, payload={"priceCount": 1, "data": PAGE_HTML})
            assert await client.prices_for(entry) == {
                "Аспирин, таблетки 100мг": 5.00,
                "Цитрамон, таблетки N10": 5.00,
            }
    finally:
        await session.close()


async def test_paginates_over_the_reported_count(request_config: RequestConfig, entry: PharmacyEntry) -> None:
    client, session = await make_client(request_config)
    try:
        with aioresponses() as mocked:
            mocked.post(ENDPOINT, payload={"priceCount": 12000, "data": ""})  # count probe
            for _ in range(3):  # 12000 prices at 5000 per page
                mocked.post(ENDPOINT, payload={"priceCount": 12000, "data": PAGE_HTML})
            await client.prices_for(entry)
            assert len(mocked.requests[("POST", aiohttp.helpers.URL(ENDPOINT))]) == 4
    finally:
        await session.close()


async def test_uses_the_configured_url(entry: PharmacyEntry) -> None:
    """B13: the configured URL used to be ignored in favour of a hardcoded host and path."""
    other = "https://example.test/prices"
    config = RequestConfig(url=HttpUrl(other), headers={"Cookie": "x=1"}, data={})
    client, session = await make_client(config)
    try:
        with aioresponses() as mocked:
            mocked.post(other + "/", payload={"priceCount": 0, "data": ""})
            mocked.post(other + "/", payload={"priceCount": 0, "data": ""})
            await client.prices_for(entry)
            assert ("POST", aiohttp.helpers.URL(other + "/")) in mocked.requests
    finally:
        await session.close()


async def test_retries_then_succeeds(request_config: RequestConfig, entry: PharmacyEntry) -> None:
    client, session = await make_client(request_config)
    try:
        with aioresponses() as mocked:
            mocked.post(ENDPOINT, status=503)
            mocked.post(ENDPOINT, payload={"priceCount": 1, "data": ""})
            mocked.post(ENDPOINT, payload={"priceCount": 1, "data": PAGE_HTML})
            assert await client.prices_for(entry) == {
                "Аспирин, таблетки 100мг": 5.00,
                "Цитрамон, таблетки N10": 5.00,
            }
    finally:
        await session.close()


async def test_gives_up_after_the_retry_budget(request_config: RequestConfig, entry: PharmacyEntry) -> None:
    client, session = await make_client(request_config, retries=2)
    try:
        with aioresponses() as mocked:
            mocked.post(ENDPOINT, status=500)
            mocked.post(ENDPOINT, status=500)
            with pytest.raises(ScrapeError, match="after 2 attempts"):
                await client.prices_for(entry)
    finally:
        await session.close()


async def test_reports_the_pharmacy_in_the_error(request_config: RequestConfig, entry: PharmacyEntry) -> None:
    """B6: failures used to vanish — the errors list was never populated."""
    client, session = await make_client(request_config, retries=1)
    try:
        with aioresponses() as mocked:
            mocked.post(ENDPOINT, status=500)
            with pytest.raises(ScrapeError, match="pharmacy 111"):
                await client.prices_for(entry)
    finally:
        await session.close()


async def test_page_size_is_narrowed_for_the_count_probe(
    request_config: RequestConfig, entry: PharmacyEntry
) -> None:
    """The first call only needs the total, so it asks for 10 rows rather than 5000."""
    client, session = await make_client(request_config)
    try:
        with aioresponses() as mocked:
            mocked.post(ENDPOINT, payload={"priceCount": 0, "data": ""})
            mocked.post(ENDPOINT, payload={"priceCount": 0, "data": ""})
            await client.prices_for(entry)
            probe, full = mocked.requests[("POST", aiohttp.helpers.URL(ENDPOINT))]
            assert "lim-result=10" in probe.kwargs["headers"]["Cookie"]
            assert "lim-result=5000" in full.kwargs["headers"]["Cookie"]
    finally:
        await session.close()


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
