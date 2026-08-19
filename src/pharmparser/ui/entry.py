"""One pharmacy row: a name field and a URL field."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from customtkinter import CTkButton

from .widgets import create_custom_entry

DELETE_CONFIRMATION_MS = 5_000


class Entry:
    def __init__(
        self,
        parent: Any,
        on_delete: Callable[[Entry], None],
        text_placeholder: str = "Pharmacy Name",
        url_placeholder: str = "https://tabletka.by/pharmacies/****",
        initial_text: str = "",
        initial_url: str = "",
    ) -> None:
        self.parent = parent
        self.on_delete = on_delete
        self.delete_confirmation_id: str | None = None
        self.text = create_custom_entry(parent, text_placeholder, initial_text)
        self.url = create_custom_entry(parent, url_placeholder, initial_url)
        self.delete_button = CTkButton(parent, text="✕", width=32, command=self.request_delete)

    def grid(
        self, text_row: int, url_row: int, column: int, padx: Any, pady: Any, sticky: str
    ) -> None:
        self.text.grid(row=text_row, column=column, padx=padx, pady=pady, sticky=sticky)
        self.url.grid(row=url_row, column=column + 1, padx=padx, pady=pady, sticky=sticky)
        self.delete_button.grid(row=text_row, column=column + 2, padx=(5, 0), pady=pady)

    def hide(self) -> None:
        self.text.grid_forget()
        self.url.grid_forget()
        self.delete_button.grid_forget()

    def request_delete(self) -> None:
        """Require a second click within five seconds before deleting the row."""
        if self.delete_confirmation_id is None:
            self.delete_button.configure(text="✓")
            self.delete_confirmation_id = self.parent.after(
                DELETE_CONFIRMATION_MS, self.reset_delete_confirmation
            )
            return

        self.parent.after_cancel(self.delete_confirmation_id)
        self.delete_confirmation_id = None
        self.on_delete(self)

    def reset_delete_confirmation(self) -> None:
        self.delete_confirmation_id = None
        self.delete_button.configure(text="✕")

    def destroy(self) -> None:
        if self.delete_confirmation_id is not None:
            self.parent.after_cancel(self.delete_confirmation_id)
            self.delete_confirmation_id = None
        self.text.destroy()
        self.url.destroy()
        self.delete_button.destroy()

    def get_text(self) -> str:
        return self.text.get()

    def get_url(self) -> str:
        return self.url.get()
