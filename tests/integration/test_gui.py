"""A smoke test for the actual window.

Everything the app *does* is covered headless through the controller, so this only
has to answer the question nothing else can: does the toolkit still build the
window and run the callbacks? That question is worth asking automatically —
customtkinter went 5.x to 6.0 during this project's life, and nothing else in the
suite would have noticed if the widgets had changed under it.

Skipped only where there is genuinely no display. Windows and macOS give Tk one
without an X server; Linux needs `DISPLAY`, so on a headless machine run it under
one:

    xvfb-run -a uv run pytest tests/integration/test_gui.py
"""

from __future__ import annotations

import contextlib
import json
import os
import sys
import threading
import time
from pathlib import Path

import pytest

from pharmparser.scraping import ScrapeError

NEEDS_X11 = sys.platform not in ("win32", "darwin")
HAS_DISPLAY = not NEEDS_X11 or bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))

pytestmark = pytest.mark.skipif(
    not HAS_DISPLAY, reason="no display; run under xvfb-run to exercise the GUI"
)

CONFIG = {
    "profiles": {"Основной": {"Аптека 1": "https://tabletka.by/pharmacies/111"}},
    "settings": {"fileName": "out.xlsx", "title": "Тест"},
    "request": {
        "url": "https://tabletka.by/ajax-request/reload-pharmacy-price",
        "headers": {"Cookie": "PHPSESSID=abc; lim-result=5000"},
        "data": {"_csrf": "token"},
    },
}


@pytest.fixture
def config_file(tmp_path: Path) -> Path:
    path = tmp_path / "config.json"
    path.write_text(json.dumps(CONFIG, ensure_ascii=False), encoding="utf-8")
    return path


@pytest.fixture
def app(config_file: Path):
    import customtkinter

    from pharmparser.ui import App

    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = App(config_file=config_file)
    window.update_idletasks()
    try:
        yield window
    finally:
        # A test may have closed it already, which destroys the window itself.
        with contextlib.suppress(Exception):
            window.destroy()


def test_the_window_builds_from_a_config(app) -> None:
    assert app.title() == "PharmParser"
    assert [profile.name for profile in app.profiles] == ["Основной"]
    assert app.current_profile is not None
    assert len(app.current_profile.entries) == 1


def test_add_and_delete_entry(app) -> None:
    app.add_entry()
    app.update_idletasks()
    assert len(app.current_profile.entries) == 2

    app.delete_entry()
    app.update_idletasks()
    assert len(app.current_profile.entries) == 1


def test_add_and_delete_profile(app) -> None:
    app.selector.add()
    app.update_idletasks()
    assert [profile.name for profile in app.profiles] == ["Основной", "Profile 2"]
    assert app.current_profile.name == "Profile 2"

    app.selector.remove()
    app.update_idletasks()
    assert [profile.name for profile in app.profiles] == ["Основной"]


def test_the_error_dialog_opens_and_clears_the_progress_bar(app) -> None:
    app.processing = True
    app.progress.grid(row=9, column=0)
    app._failed("a test failure")
    app.update_idletasks()
    assert app.processing is False


def run_loop(app, action, timeout: float = 10.0) -> None:
    """Run the real event loop, doing ``action`` inside it, until the app goes idle.

    The click has to happen while the loop is running, exactly as it does for a
    user: the worker hands its result back with ``self.after``, and Tk only accepts
    a call from another thread while the main thread sits in ``mainloop``. Polling
    with ``update()`` instead would raise "main thread is not in main loop" — which
    is the whole reason B5's rule is "marshal back", not "call from the worker".
    """
    deadline = time.monotonic() + timeout
    timed_out: list[bool] = []

    def check() -> None:
        if not app.processing:
            app.quit()
        elif time.monotonic() > deadline:
            timed_out.append(True)
            app.quit()
        else:
            app.after(10, check)

    def start() -> None:
        action()
        app.after(10, check)

    app.after(0, start)
    app.mainloop()
    if timed_out:
        raise AssertionError("the worker never reported back")


def test_parse_runs_off_the_main_thread_and_reports_back(
    app, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """B5: Tkinter is not thread-safe, so the worker must touch no widget.

    Checked by watching where the work happens and where the result is delivered:
    the scrape on a worker thread, the file opened on Tk's own.
    """
    workbook = tmp_path / "report.xlsx"
    workbook.write_bytes(b"")
    worked_on: list[str] = []
    opened: list[tuple[Path, str]] = []

    def record_thread(profile, use_cache=False, path=None) -> Path:
        worked_on.append(threading.current_thread().name)
        return workbook

    monkeypatch.setattr(app.controller, "run", record_thread)
    monkeypatch.setattr(
        "pharmparser.ui.app.open_file",
        lambda path: opened.append((path, threading.current_thread().name)),
    )

    run_loop(app, app.click)

    main = threading.main_thread().name
    assert worked_on and worked_on[0] != main, "the scrape ran off the main thread"
    assert opened == [(workbook, main)], "the result was delivered on the main thread"


def test_parse_passes_the_edited_entries_and_the_cache_choice(
    app, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    calls: list[tuple] = []
    monkeypatch.setattr("pharmparser.ui.app.open_file", lambda path: None)
    def record_call(profile, use_cache=False, path=None) -> Path:
        calls.append((profile, use_cache))
        return tmp_path

    monkeypatch.setattr(app.controller, "run", record_call)

    def click_with_cache() -> None:
        app.cache_checkbox.select()
        app.click()

    run_loop(app, click_with_cache)

    profile, use_cache = calls[0]
    assert use_cache is True
    assert [entry.name for entry in profile.pharmacies] == ["Аптека 1"]


def test_a_second_parse_is_ignored_while_one_is_running(app, monkeypatch: pytest.MonkeyPatch) -> None:
    release = threading.Event()
    calls: list[int] = []

    def blocking_run(profile, use_cache=False, path=None):
        calls.append(1)
        release.wait(timeout=10)
        raise ScrapeError("done blocking")

    monkeypatch.setattr(app.controller, "run", blocking_run)
    monkeypatch.setattr("pharmparser.ui.app.CTkMessagebox", lambda **kwargs: None)

    def click_twice() -> None:
        app.click()
        app.click()  # ignored: the guard is set synchronously by the first
        app.after(50, release.set)

    run_loop(app, click_twice)

    assert calls == [1]


def test_a_scrape_failure_reaches_the_dialog(app, monkeypatch: pytest.MonkeyPatch) -> None:
    """B6 through the UI: the error text is the user's only signal."""
    shown: list[str] = []
    def fail(profile, use_cache=False, path=None) -> Path:
        raise ScrapeError("tabletka.by said no")

    monkeypatch.setattr(app.controller, "run", fail)
    monkeypatch.setattr("pharmparser.ui.app.CTkMessagebox", lambda **kwargs: shown.append(kwargs["message"]))

    run_loop(app, app.click)

    assert shown == ["tabletka.by said no"]


def test_an_unexpected_failure_is_named_rather_than_hidden(app, monkeypatch: pytest.MonkeyPatch) -> None:
    shown: list[str] = []
    def explode(profile, use_cache=False, path=None) -> Path:
        raise ZeroDivisionError("division by zero")

    monkeypatch.setattr(app.controller, "run", explode)
    monkeypatch.setattr("pharmparser.ui.app.CTkMessagebox", lambda **kwargs: shown.append(kwargs["message"]))

    run_loop(app, app.click)

    assert shown == ["ZeroDivisionError: division by zero"]


def test_the_progress_bar_is_cleared_afterwards(app, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("pharmparser.ui.app.open_file", lambda path: None)
    monkeypatch.setattr(app.controller, "run", lambda profile, use_cache=False, path=None: tmp_path)

    run_loop(app, app.click)

    assert app.processing is False
    assert not app.progress.winfo_ismapped()


def test_closing_saves_the_edited_profiles(app, config_file: Path) -> None:
    """A6: names the user chose must survive the round trip."""
    app.profiles[0].name = "Переименованный"
    app.on_closing()

    saved = json.loads(config_file.read_text(encoding="utf-8"))
    assert list(saved["profiles"]) == ["Переименованный"]
    assert saved["settings"]["title"] == "Тест"
    assert saved["request"]["headers"]["Cookie"].startswith("PHPSESSID=")
