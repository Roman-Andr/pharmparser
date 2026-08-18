from ..config import PharmacyEntry
from ..config import Profile as ProfileConfig
from .entry import Entry


class Profile:
    """The widgets for one profile's pharmacy list."""

    __slots__ = ["entries", "name", "parent"]

    def __init__(self, parent, config: ProfileConfig):
        self.parent = parent
        self.name = config.name
        self.entries = [
            Entry(parent, initial_text=entry.name, initial_url=entry.url) for entry in config.pharmacies
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
            entry.destroy()

    def display(self) -> None:
        for i, entry in enumerate(self.entries):
            entry.grid(text_row=i + 2, url_row=i + 2, column=0, padx=(5, 0), pady=(5, 5), sticky="nsew")

    def add_entry(self) -> None:
        self.entries.append(Entry(self.parent))
        self.display()

    def delete_entry(self) -> None:
        if self.entries:
            self.entries.pop().destroy()
            self.display()
