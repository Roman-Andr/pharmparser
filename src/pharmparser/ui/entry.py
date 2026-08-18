"""One pharmacy row: a name field and a URL field."""

from __future__ import annotations

from typing import Any

from .widgets import create_custom_entry


class Entry:
    def __init__(
        self,
        parent: Any,
        text_placeholder: str = "Pharmacy Name",
        url_placeholder: str = "https://tabletka.by/pharmacies/****",
        initial_text: str = "",
        initial_url: str = "",
    ) -> None:
        self.text = create_custom_entry(parent, text_placeholder, initial_text)
        self.url = create_custom_entry(parent, url_placeholder, initial_url)

    def grid(
        self, text_row: int, url_row: int, column: int, padx: Any, pady: Any, sticky: str
    ) -> None:
        self.text.grid(row=text_row, column=column, padx=padx, pady=pady, sticky=sticky)
        self.url.grid(row=url_row, column=column + 1, padx=padx, pady=pady, sticky=sticky)

    def destroy(self) -> None:
        self.text.grid_forget()
        self.url.grid_forget()

    def get_text(self) -> str:
        return self.text.get()

    def get_url(self) -> str:
        return self.url.get()
