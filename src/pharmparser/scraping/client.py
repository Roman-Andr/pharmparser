"""HTTP access to the tabletka.by price endpoint."""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Mapping
from types import TracebackType

import aiohttp
from lxml import html as lxml_html
from pydantic import BaseModel, Field
from yarl import URL

from ..config import PharmacyEntry, RequestConfig
from ..domain import ProductPrice
from .parallel import ParsePool

logger = logging.getLogger(__name__)

PAGE_SIZE = 5000
DEFAULT_TIMEOUT = aiohttp.ClientTimeout(total=120, connect=15, sock_read=60)
DEFAULT_RETRIES = 3
DEFAULT_BACKOFF = 1.0

RETRYABLE_STATUSES = frozenset({408, 429})
"""4xx codes worth a second try. Every other 4xx is the request's own fault."""

STALE_SESSION_STATUSES = frozenset({400, 401, 403})
"""What the endpoint answers once the session cookies or the CSRF token expire."""

STALE_SESSION_HINT = (
    "the session looks expired and could not be refreshed automatically — open the pharmacy "
    "page in a browser and check that tabletka.by is available"
)


class ScrapeError(Exception):
    """A pharmacy could not be scraped. Carries a message meant for the user."""


class PricePage(BaseModel):
    """The endpoint's JSON envelope.

    Validated rather than trusted: the old code called ``json.loads`` and indexed
    straight into the result, so any error page produced a KeyError far from the
    cause (B13).
    """

    model_config = {"extra": "ignore"}

    data: str = ""
    """An HTML fragment holding the results table."""
    price_count: int = Field(default=0, alias="priceCount")


class TabletkaClient:
    """Fetches a pharmacy's full price list, one page at a time.

    Unlike the previous ``http.client`` implementation this honours the configured
    URL (B13), applies timeouts and retries, checks status codes, and does not
    raise the process to realtime priority (B11).
    """

    def __init__(
        self,
        config: RequestConfig,
        session: aiohttp.ClientSession,
        *,
        parse_pool: ParsePool | None = None,
        retries: int = DEFAULT_RETRIES,
        backoff: float = DEFAULT_BACKOFF,
    ) -> None:
        self._config = config
        self._session = session
        self._parse_pool = parse_pool or ParsePool(0)
        self._retries = retries
        self._backoff = backoff
        self._session_generation = 0
        self._csrf = config.data.get("_csrf", "")
        self._refresh_lock = asyncio.Lock()
        placeholders = {"redacted", "<redacted>", "..."}
        configured_cookie = config.cookie.strip().casefold()
        self._configured_session_usable = (
            "=" in config.cookie
            and configured_cookie not in placeholders
            and "redacted" not in configured_cookie
        )

    def _headers(self, limit: int, *, configured_cookie: bool) -> dict[str, str]:
        excluded = {"cookie", "content-length"}
        if not configured_cookie:
            excluded.add("host")
        headers = {
            key: value
            for key, value in self._config.headers.items()
            if key.casefold() not in excluded
        }
        headers.setdefault("Accept", "application/json, text/javascript, */*; q=0.01")
        headers.setdefault("X-Requested-With", "XMLHttpRequest")
        if configured_cookie:
            cookie = self._config.cookie.replace(f"lim-result={PAGE_SIZE}", f"lim-result={limit}")
            if cookie:
                headers["Cookie"] = cookie
        elif self._session_generation > 0:
            cookies = self._session.cookie_jar.filter_cookies(URL(self._config.endpoint))
            values = {
                key: morsel.value
                for key, morsel in cookies.items()
                if key.casefold() != "lim-result"
            }
            values["lim-result"] = str(limit)
            headers["Cookie"] = "; ".join(f"{key}={value}" for key, value in values.items())
        return headers

    async def _refresh_public_session(self, pharmacy_id: str, observed_generation: int) -> None:
        """Obtain a fresh public CSRF token and regional cookie from a pharmacy page.

        tabletka.by does not require an authenticated account for price lists. Its
        pharmacy page establishes the short-lived session used by the AJAX endpoint,
        so stale legacy cookies can be repaired without asking the user for secrets.
        """
        async with self._refresh_lock:
            if self._session_generation != observed_generation:
                return
            endpoint = URL(self._config.endpoint)
            pharmacy_url = endpoint.with_path(f"/pharmacies/{pharmacy_id}/").with_query(None)
            try:
                async with self._session.get(
                    pharmacy_url,
                    headers=self._headers(PAGE_SIZE, configured_cookie=False),
                ) as response:
                    response.raise_for_status()
                    page = await response.text()
                root = lxml_html.fromstring(page)
                tokens = root.xpath('//meta[@name="csrf-token"]/@content')
                token = str(tokens[0]).strip() if tokens else ""
                if not token:
                    raise ValueError("на странице аптеки отсутствует CSRF-токен")
            except (aiohttp.ClientError, ValueError) as error:
                raise ScrapeError(f"{STALE_SESSION_HINT}: {error}") from error
            self._csrf = token
            self._session_generation += 1

    async def _post(self, pharmacy_id: str, page: int, limit: int) -> PricePage:
        last_error: Exception | None = None
        attempt = 1
        refreshed = False
        if not self._configured_session_usable and self._session_generation == 0:
            await self._refresh_public_session(pharmacy_id, observed_generation=0)
            refreshed = True
        while attempt <= self._retries:
            generation = self._session_generation
            headers = self._headers(limit, configured_cookie=generation == 0)
            payload = {
                **self._config.data,
                "_csrf": self._csrf,
                "id": pharmacy_id,
                "page": str(page),
            }
            try:
                async with self._session.post(
                    self._config.endpoint, data=payload, headers=headers
                ) as response:
                    response.raise_for_status()
                    return PricePage.model_validate(await response.json(content_type=None))
            except aiohttp.ClientResponseError as error:
                if error.status in STALE_SESSION_STATUSES and not refreshed:
                    await self._refresh_public_session(pharmacy_id, generation)
                    refreshed = True
                    continue
                # A 4xx will answer the same way however often it is asked; retrying
                # only delays the report and hammers the endpoint.
                if error.status < 500 and error.status not in RETRYABLE_STATUSES:
                    raise ScrapeError(self._explain(pharmacy_id, error)) from error
                last_error = error
            except (TimeoutError, aiohttp.ClientError, ValueError) as error:
                last_error = error

            if attempt == self._retries:
                break
            delay = self._backoff * 2 ** (attempt - 1)
            logger.warning(
                "Request for pharmacy %s page %d failed (%s); retrying in %.1fs",
                pharmacy_id, page, last_error, delay,
            )
            await asyncio.sleep(delay)
            attempt += 1

        raise ScrapeError(
            f"could not fetch prices for pharmacy {pharmacy_id} after {self._retries} attempts: {last_error}"
        ) from last_error

    @staticmethod
    def _explain(pharmacy_id: str, error: aiohttp.ClientResponseError) -> str:
        """Turn a bare status code into something the user can act on."""
        if error.status in STALE_SESSION_STATUSES:
            return f"pharmacy {pharmacy_id}: HTTP {error.status} — {STALE_SESSION_HINT}"
        if error.status == 404:
            return f"pharmacy {pharmacy_id}: HTTP 404 — check the pharmacy URL in your config"
        return f"pharmacy {pharmacy_id}: HTTP {error.status} {error.message}"

    async def prices_for(self, entry: PharmacyEntry) -> Mapping[str, float]:
        """Fetch and parse every page of one pharmacy's price list.

        The first page carries ``priceCount`` itself, so there is no separate probe
        request: asking for the count first cost an extra round trip per pharmacy
        and told us nothing the page we needed anyway did not already say.
        """
        first = await self._post(entry.pharmacy_id, page=1, limit=PAGE_SIZE)
        page_count = max(1, -(-first.price_count // PAGE_SIZE))
        logger.info("%s: %d prices across %d page(s)", entry.name, first.price_count, page_count)

        rest = [
            await self._post(entry.pharmacy_id, page=page, limit=PAGE_SIZE)
            for page in range(2, page_count + 1)
        ]
        return await self._parse_pool.parse([page.data for page in (first, *rest)])

    async def product_prices_for(self, entry: PharmacyEntry) -> list[ProductPrice]:
        """Fetch a pharmacy using the structured product representation."""
        first = await self._post(entry.pharmacy_id, page=1, limit=PAGE_SIZE)
        page_count = max(1, -(-first.price_count // PAGE_SIZE))
        rest = [
            await self._post(entry.pharmacy_id, page=page, limit=PAGE_SIZE)
            for page in range(2, page_count + 1)
        ]
        prices = await self._parse_pool.parse_products([page.data for page in (first, *rest)])
        logger.info(
            "%s: parsed %d structured prices out of %d reported",
            entry.name,
            len(prices),
            first.price_count,
        )
        if first.price_count and not prices:
            raise ScrapeError(
                f"pharmacy {entry.pharmacy_id}: tabletka.by reported {first.price_count} prices, "
                "but the parser read none; the site markup may have changed"
            )
        return prices


class ClientSessionFactory:
    """Owns the aiohttp session so callers do not have to."""

    def __init__(
        self,
        config: RequestConfig,
        *,
        parse_pool: ParsePool | None = None,
        timeout: aiohttp.ClientTimeout = DEFAULT_TIMEOUT,
    ) -> None:
        self._config = config
        self._parse_pool = parse_pool or ParsePool(0)
        self._timeout = timeout
        self._session: aiohttp.ClientSession | None = None

    async def __aenter__(self) -> TabletkaClient:
        self._session = aiohttp.ClientSession(timeout=self._timeout)
        return TabletkaClient(self._config, self._session, parse_pool=self._parse_pool)

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if self._session is not None:
            await self._session.close()
            self._session = None
