"""A smoke test for the actual window.

Everything the app *does* is covered headless through the controller, so this only
has to answer the question nothing else can: does the toolkit still build the
window and run the callbacks? That question is worth asking automatically —
customtkinter went 5.x to 6.0 during this project's life, and nothing else in the
suite would have noticed if the widgets had changed under it.

Skipped when there is no display. On a headless machine, run it under one:

    xvfb-run -a uv run pytest tests/integration/test_gui.py
"""

from __future__ import annotations

import contextlib
import json
import os
from pathlib import Path

import pytest

pytestmark = pytest.mark.skipif(
    not os.environ.get("DISPLAY"), reason="no display; run under xvfb-run to exercise the GUI"
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


def test_closing_saves_the_edited_profiles(app, config_file: Path) -> None:
    """A6: names the user chose must survive the round trip."""
    app.profiles[0].name = "Переименованный"
    app.on_closing()

    saved = json.loads(config_file.read_text(encoding="utf-8"))
    assert list(saved["profiles"]) == ["Переименованный"]
    assert saved["settings"]["title"] == "Тест"
    assert saved["request"]["headers"]["Cookie"].startswith("PHPSESSID=")
