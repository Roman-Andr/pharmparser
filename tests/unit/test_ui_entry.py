from __future__ import annotations

from collections.abc import Callable
from typing import Any

from pharmparser.ui.entry import DELETE_CONFIRMATION_MS, Entry


class FakeParent:
    def __init__(self) -> None:
        self.delay: int | None = None
        self.callback: Callable[[], None] | None = None
        self.cancelled: list[str] = []

    def after(self, delay: int, callback: Callable[[], None]) -> str:
        self.delay = delay
        self.callback = callback
        return "delete-confirmation"

    def after_cancel(self, confirmation_id: str) -> None:
        self.cancelled.append(confirmation_id)


class FakeButton:
    def __init__(self) -> None:
        self.text = "✕"

    def configure(self, *, text: str) -> None:
        self.text = text


def make_entry() -> tuple[Any, FakeParent, FakeButton, list[Any]]:
    parent = FakeParent()
    button = FakeButton()
    deleted: list[Any] = []
    entry: Any = Entry.__new__(Entry)
    entry.parent = parent
    entry.on_delete = deleted.append
    entry.delete_confirmation_id = None
    entry.delete_button = button
    return entry, parent, button, deleted


def test_first_delete_click_arms_confirmation_for_five_seconds() -> None:
    entry, parent, button, deleted = make_entry()

    entry.request_delete()

    assert button.text == "✓"
    assert parent.delay == DELETE_CONFIRMATION_MS == 5_000
    assert deleted == []


def test_delete_confirmation_expires_without_removing_the_entry() -> None:
    entry, parent, button, deleted = make_entry()
    entry.request_delete()
    assert parent.callback is not None

    parent.callback()

    assert button.text == "✕"
    assert entry.delete_confirmation_id is None
    assert deleted == []


def test_second_delete_click_removes_the_entry() -> None:
    entry, parent, _button, deleted = make_entry()
    entry.request_delete()

    entry.request_delete()

    assert deleted == [entry]
    assert parent.cancelled == ["delete-confirmation"]
