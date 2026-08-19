"""HTTP access to the tabletka.by price endpoint."""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Mapping
from types import TracebackType

import aiohttp
from pydantic import BaseModel, Field

from ..config import PharmacyEntry, RequestConfig
from .parser import merge, parse_page

logger = logging.getLogger(__name__)

PAGE_SIZE = 5000
COUNT_PROBE_LIMIT = 10
"""Rows requested by the first call, which only needs the total count."""

DEFAULT_TIMEOUT = aiohttp.ClientTimeout(total=120, connect=15, sock_read=60)
DEFAULT_RETRIES = 3
DEFAULT_BACKOFF = 1.0

RETRYABLE_STATUSES = frozenset({408, 429})
"""4xx codes worth a second try. Every other 4xx is the request's own fault."""

STALE_SESSION_STATUSES = frozenset({400, 401, 403})
"""What the endpoint answers once the session cookies or the CSRF token expire."""

STALE_SESSION_HINT = (
    "the session looks expired — refresh the Cookie header and the _csrf value in your "
    "config from the browser's DevTools (Network tab, any request to the prices page)"
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
        retries: int = DEFAULT_RETRIES,
        backoff: float = DEFAULT_BACKOFF,
    ) -> None:
        self._config = config
        self._session = session
        self._retries = retries
        self._backoff = backoff

    async def _post(self, pharmacy_id: str, page: int, limit: int) -> PricePage:
        cookie = self._config.cookie.replace(f"lim-result={PAGE_SIZE}", f"lim-result={limit}")
        headers = {**self._config.headers, "Cookie": cookie}
        payload = {**self._config.data, "id": pharmacy_id, "page": str(page)}

        last_error: Exception | None = None
        for attempt in range(1, self._retries + 1):
            try:
                async with self._session.post(
                    self._config.endpoint, data=payload, headers=headers
                ) as response:
                    response.raise_for_status()
                    return PricePage.model_validate(await response.json(content_type=None))
            except aiohttp.ClientResponseError as error:
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
        """Fetch and parse every page of one pharmacy's price list."""
        first = await self._post(entry.pharmacy_id, page=0, limit=COUNT_PROBE_LIMIT)
        page_count = max(1, -(-first.price_count // PAGE_SIZE))
        logger.info("%s: %d prices across %d page(s)", entry.name, first.price_count, page_count)

        pages = [await self._post(entry.pharmacy_id, page=page, limit=PAGE_SIZE) for page in range(1, page_count + 1)]
        return merge([parse_page(page.data) for page in pages])


class ClientSessionFactory:
    """Owns the aiohttp session so callers do not have to."""

    def __init__(self, config: RequestConfig, *, timeout: aiohttp.ClientTimeout = DEFAULT_TIMEOUT) -> None:
        self._config = config
        self._timeout = timeout
        self._session: aiohttp.ClientSession | None = None

    async def __aenter__(self) -> TabletkaClient:
        self._session = aiohttp.ClientSession(timeout=self._timeout)
        return TabletkaClient(self._config, self._session)

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if self._session is not None:
            await self._session.close()
            self._session = None
