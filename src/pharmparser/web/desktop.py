"""Windows-first desktop host with pywebview and browser fallback."""

from __future__ import annotations

import contextlib
import logging
import socket
import threading
import time
import webbrowser
from collections.abc import Callable
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

import uvicorn

from ..config.paths import modern_log_path
from ..logging_ import configure
from .api import AppServices, create_app, create_services

logger = logging.getLogger(__name__)
CONTROL_PORT = 45839
HEARTBEAT_TIMEOUT = timedelta(minutes=5)


class SingleInstance:
    def __init__(self) -> None:
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.primary = False
        self.url = ""
        try:
            self.socket.bind(("127.0.0.1", CONTROL_PORT))
            self.socket.listen(2)
            self.primary = True
        except OSError:
            self.socket.close()

    def activate_existing(self) -> bool:
        if self.primary:
            return False
        with socket.create_connection(("127.0.0.1", CONTROL_PORT), timeout=2) as client:
            client.sendall(b"ACTIVATE\n")
            url = client.recv(4096).decode("utf-8").strip()
        if url:
            webbrowser.open(url)
        return True

    def serve(self, url: str, activate: Callable[[], None]) -> None:
        self.url = url

        def listen() -> None:
            while self.primary:
                try:
                    client, _ = self.socket.accept()
                except OSError:
                    return
                with client:
                    if client.recv(128).startswith(b"ACTIVATE"):
                        client.sendall(f"{self.url}\n".encode())
                        activate()

        threading.Thread(target=listen, name="single-instance", daemon=True).start()

    def close(self) -> None:
        self.primary = False
        with contextlib.suppress(OSError):
            self.socket.close()


def _server(services: AppServices) -> tuple[uvicorn.Server, socket.socket, str]:
    listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    listener.bind(("127.0.0.1", 0))
    listener.listen(128)
    port = listener.getsockname()[1]
    static = Path(__file__).resolve().parent / "static"
    app = create_app(services, frontend_dir=static, production=True)
    server = uvicorn.Server(
        uvicorn.Config(app, host="127.0.0.1", port=port, log_config=None, access_log=False)
    )
    return server, listener, f"http://127.0.0.1:{port}/#{services.token}"


def _run_server(server: uvicorn.Server, listener: socket.socket) -> threading.Thread:
    thread = threading.Thread(target=server.run, kwargs={"sockets": [listener]}, name="local-api", daemon=True)
    thread.start()
    deadline = time.monotonic() + 10
    while not server.started and thread.is_alive() and time.monotonic() < deadline:
        time.sleep(0.02)
    if not server.started:
        raise RuntimeError("локальный сервер не запустился")
    return thread


def _monitor_browser(services: AppServices, server: uvicorn.Server) -> None:
    while not server.should_exit:
        time.sleep(10)
        stale = datetime.now(UTC) - services.last_heartbeat > HEARTBEAT_TIMEOUT
        if stale and services.runs.active_run_id is None:
            logger.info("Browser heartbeat expired; stopping local backend")
            server.should_exit = True


def run_desktop() -> int:
    configure(path=modern_log_path())
    instance = SingleInstance()
    if instance.activate_existing():
        return 0

    services = create_services()
    server, listener, url = _server(services)
    thread = _run_server(server, listener)
    window: Any = None

    def activate() -> None:
        if window is not None:
            with contextlib.suppress(Exception):
                window.restore()
                window.on_top = True
                window.on_top = False

    instance.serve(url, activate)
    try:
        try:
            import webview

            config = services.settings.load()
            window = webview.create_window(
                "PharmParser",
                url,
                width=config.window_width,
                height=config.window_height,
                min_size=(960, 640),
                x=config.window_x,
                y=config.window_y,
                text_select=True,
            )
            webview.start(debug=False)
        except Exception as error:
            logger.warning("WebView недоступен, открываем системный браузер: %s", error)
            webbrowser.open(url)
            monitor = threading.Thread(
                target=_monitor_browser, args=(services, server), name="browser-heartbeat", daemon=True
            )
            monitor.start()
            while thread.is_alive() and not services.shutdown.is_set():
                time.sleep(0.2)
    finally:
        server.should_exit = True
        thread.join(timeout=10)
        listener.close()
        instance.close()
    return 0
