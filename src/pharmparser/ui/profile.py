from __future__ import annotations

from typing import Any

from ..config import PharmacyEntry
from ..config import Profile as ProfileConfig
from .entry import Entry


class Profile:
    """The widgets for one profile's pharmacy list."""

    def __init__(self, parent: Any, config: ProfileConfig) -> None:
        self.parent = parent
        self.name = config.name
        self.entries = [
            Entry(parent, self.delete_entry, initial_text=entry.name, initial_url=entry.url)
            for entry in config.pharmacies
        ]

    def to_config(self) -> ProfileConfig:
        """Snapshot the current widget contents back into a config model."""
        return ProfileConfig(
            name=self.name,
            pharmacies=[
                PharmacyEntry(name=entry.get_text(), url=entry.get_url()) for entry in self.entries
            ],
        )

    def hide(self) -> None:
        for entry in self.entries:
            entry.hide()

    def display(self) -> None:
        for i, entry in enumerate(self.entries):
            entry.grid(text_row=i + 2, url_row=i + 2, column=0, padx=(5, 0), pady=(5, 5), sticky="nsew")

    def add_entry(self) -> None:
        self.entries.append(Entry(self.parent, self.delete_entry))
        self.display()

    def delete_entry(self, entry: Entry | None = None) -> None:
        if not self.entries:
            return

        target = entry or self.entries[-1]
        if target not in self.entries:
            return
        self.entries.remove(target)
        target.destroy()
        self.display()
