"""Application state and use cases, with no front end attached.

Everything the window used to do besides drawing itself lives here: loading and
saving the configuration, choosing between the cache and a fresh scrape, running
the scrape, and exporting the workbook. It imports no toolkit, so it runs — and is
tested — without a display, and both front ends drive it: ``ui.App`` from its
callbacks and ``cli`` straight through.

That split is what A1 was about: ``ui/app.py`` used to own config I/O, thread
management, cache read/write, export orchestration and shelling out to the OS in
130 lines, so none of it could be exercised at all.
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Callable, Iterable, Sequence
from pathlib import Path

from .cache import read_table, write_table
from .config import (
    AppConfig,
    ConfigError,
    ExportSettings,
    Profile,
    RequestConfig,
    cache_path,
    config_path,
    load_config,
    save_config,
)
from .domain import PriceTable
from .export import Exporter, select_exporter
from .scraping import scrape_profile

logger = logging.getLogger(__name__)

DEFAULT_PROFILE_NAME = "Profile 1"

ScrapeFn = Callable[[RequestConfig, Sequence], PriceTable]
"""Injection seam for the scrape, so tests need no network."""


def _scrape(request: RequestConfig, pharmacies: Sequence) -> PriceTable:
    return asyncio.run(scrape_profile(request, pharmacies))


class Controller:
    """Owns the configuration and the parse/export use cases."""

    def __init__(
        self,
        config: AppConfig,
        config_file: Path,
        *,
        exporter: Exporter | None = None,
        scrape: ScrapeFn = _scrape,
        cache_file: Callable[[str], Path] = cache_path,
    ) -> None:
        self.config = config
        self.config_file = config_file
        self._exporter = exporter
        self._scrape = scrape
        self._cache_file = cache_file

    @classmethod
    def load(cls, config_file: Path | None = None, **kwargs: object) -> Controller:
        """Read the configuration from disk. Raises ``ConfigError`` with the reason."""
        path = config_file or config_path()
        return cls(load_config(path), path, **kwargs)  # type: ignore[arg-type]

    # -- configuration ---------------------------------------------------------

    @property
    def settings(self) -> ExportSettings:
        return self.config.settings

    @property
    def request(self) -> RequestConfig:
        return self.config.request

    @property
    def profiles(self) -> list[Profile]:
        """The configured profiles, or a single empty one to start from."""
        return self.config.profiles or [Profile(name=DEFAULT_PROFILE_NAME)]

    @staticmethod
    def next_profile_name(existing: Iterable[str]) -> str:
        """A "Profile N" that is not taken yet."""
        taken = set(existing)
        index = len(taken) + 1
        while f"Profile {index}" in taken:
            index += 1
        return f"Profile {index}"

    def save(self, profiles: Sequence[Profile]) -> None:
        """Write the configuration back, keeping the user's profile names (A6)."""
        self.config = self.config.model_copy(update={"profiles": list(profiles)})
        save_config(self.config, self.config_file)

    # -- use cases -------------------------------------------------------------

    def collect(self, profile: Profile, use_cache: bool = False) -> PriceTable:
        """Prices for one profile, scraped or read from that profile's own cache.

        The cache is entirely opt-in, and reading and writing it are the same
        choice (A7): the old code wrote one shared ``data.json`` on every run
        whatever the checkbox said, so switching profile silently reused another
        profile's prices.
        """
        if not use_cache:
            return self._scrape(self.request, profile.pharmacies)

        cache = self._cache_file(profile.name)
        if cache.exists():
            try:
                logger.info("Reading cached prices for %r from %s", profile.name, cache)
                return read_table(cache)
            except Exception:
                logger.warning("Ignoring unreadable cache at %s", cache, exc_info=True)

        table = self._scrape(self.request, profile.pharmacies)
        try:
            cache.parent.mkdir(parents=True, exist_ok=True)
            write_table(table, cache)
        except OSError:
            logger.warning("Could not write the cache at %s", cache, exc_info=True)
        return table

    def export(self, table: PriceTable, path: Path | None = None) -> Path:
        """Write the workbook, using the best backend this machine supports."""
        exporter = self._exporter or select_exporter()
        return exporter.export(self.settings, table, path).absolute()

    def run(self, profile: Profile, use_cache: bool = False, path: Path | None = None) -> Path:
        """The whole job: collect prices and export them. Returns the workbook path."""
        return self.export(self.collect(profile, use_cache), path)

    def select_profile(self, name: str | None) -> Profile:
        """The named profile, or the first one. Raises ``ConfigError`` if there is none."""
        profiles = self.config.profiles
        if not profiles:
            raise ConfigError("no profiles are configured")
        if name is None:
            return profiles[0]
        for profile in profiles:
            if profile.name == name:
                return profile
        available = ", ".join(profile.name for profile in profiles)
        raise ConfigError(f"no profile named {name!r}; available profiles: {available}")
