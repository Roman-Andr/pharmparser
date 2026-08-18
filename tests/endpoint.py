"""A real HTTP server standing in for the tabletka.by price endpoint.

Tests used to fake the network with ``aioresponses``, which works by reaching into
aiohttp's internals — so it broke outright on aiohttp 3.14 (``ClientResponse``
gained a required ``stream_writer`` argument) with no released fix. Serving real
HTTP on localhost instead costs a few milliseconds per test and exercises the
client the way production does: real sockets, real status codes, real headers.

Responses are chosen from the request's own fields rather than from a queue, so
the fan-out's concurrency cannot make a test flaky.
"""

from __future__ import annotations

import json
import threading
from collections.abc import Iterator, Mapping
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs

DEFAULT_PHARMACY_HTML = ""


@dataclass(frozen=True)
class ReceivedRequest:
    """One request the endpoint was sent."""

    method: str
    path: str
    headers: Mapping[str, str]
    form: Mapping[str, str]

    @property
    def cookie(self) -> str:
        return self.headers.get("Cookie", "")

    @property
    def pharmacy_id(self) -> str:
        return self.form.get("id", "")

    @property
    def page(self) -> str:
        return self.form.get("page", "")


@dataclass
class _Page:
    html: str
    price_count: int


class FakeEndpoint:
    """Serves the price endpoint's contract over real HTTP.

    The client asks for ``page=0`` to learn the total and ``page>=1`` for data, so
    that is what the responses key off — never call order.
    """

    def __init__(self) -> None:
        self.requests: list[ReceivedRequest] = []
        self._pages: dict[str, _Page] = {}
        self._default: _Page | None = None
        self._status: int | None = None
        self._transient: list[int] = []
        self._lock = threading.Lock()
        self._server: ThreadingHTTPServer | None = None

    # -- configuring -----------------------------------------------------------

    def serve(self, pharmacy_id: str, html: str, price_count: int = 1) -> None:
        """Answer requests for one pharmacy with ``html``."""
        self._pages[pharmacy_id] = _Page(html, price_count)

    def serve_all(self, html: str, price_count: int = 1) -> None:
        """Answer every pharmacy the same way."""
        self._default = _Page(html, price_count)

    def fail(self, status: int = 500) -> None:
        """Answer everything with ``status`` from now on."""
        self._status = status

    def fail_next(self, status: int, times: int = 1) -> None:
        """Fail the next ``times`` requests, then behave normally again."""
        self._transient.extend([status] * times)

    # -- running ---------------------------------------------------------------

    @property
    def url(self) -> str:
        assert self._server is not None, "the endpoint is not running"
        host, port = self._server.server_address[:2]
        return f"http://{host!s}:{port}/ajax-request/reload-pharmacy-price"

    def _handle(self, request: ReceivedRequest) -> tuple[int, bytes]:
        with self._lock:
            self.requests.append(request)
            if self._transient:
                return self._transient.pop(0), b'{"error": "transient"}'
            if self._status is not None:
                return self._status, b'{"error": "configured failure"}'
            page = self._pages.get(request.pharmacy_id, self._default)

        if page is None:
            return 404, b'{"error": "no pharmacy configured for this id"}'
        body = {
            "status": 1,
            "priceCount": page.price_count,
            # The count probe's body is discarded, and the live site sends it empty.
            "data": DEFAULT_PHARMACY_HTML if request.page == "0" else page.html,
        }
        return 200, json.dumps(body, ensure_ascii=False).encode("utf-8")


class _Handler(BaseHTTPRequestHandler):
    endpoint: FakeEndpoint

    protocol_version = "HTTP/1.1"

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length).decode("utf-8")
        form = {key: values[0] for key, values in parse_qs(raw, keep_blank_values=True).items()}
        status, body = self.endpoint._handle(
            ReceivedRequest(
                method="POST", path=self.path, headers=dict(self.headers.items()), form=form
            )
        )
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args: object) -> None:
        """Silence the default stderr access log."""


def running_endpoint() -> Iterator[FakeEndpoint]:
    """Start a :class:`FakeEndpoint` on a free port for the duration of a test."""
    endpoint = FakeEndpoint()
    handler = type("BoundHandler", (_Handler,), {"endpoint": endpoint})
    with ThreadingHTTPServer(("127.0.0.1", 0), handler) as server:
        endpoint._server = server
        # serve_forever polls at 0.5s by default, which each test would pay on teardown.
        thread = threading.Thread(target=server.serve_forever, kwargs={"poll_interval": 0.01}, daemon=True)
        thread.start()
        try:
            yield endpoint
        finally:
            server.shutdown()
            thread.join(timeout=5)
