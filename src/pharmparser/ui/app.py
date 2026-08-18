import asyncio
import logging
from pathlib import Path
from threading import Thread

from CTkMessagebox import CTkMessagebox
from customtkinter import CTk, CTkButton, CTkCheckBox, CTkProgressBar

from ..cache import read_table, write_table
from ..config import AppConfig, ConfigError, cache_path, config_path, load_config, save_config
from ..config import Profile as ProfileConfig
from ..domain import PriceTable
from ..export import select_exporter
from ..platform_ import open_file
from ..scraping import NoPharmaciesError, ScrapeError, scrape_profile
from .profile import Profile
from .profile_selector import ProfileSelector

logger = logging.getLogger(__name__)


class App(CTk):
    processing: bool = False

    def __init__(self, config: AppConfig | None = None, config_file: Path | None = None):
        super().__init__()

        self.config_file = config_file or config_path()
        self.app_config = config if config is not None else load_config(self.config_file)

        self.geometry("1100x600")
        self.title("PharmParser")

        self.progress: CTkProgressBar = CTkProgressBar(self)
        self.profiles: list[Profile] = [
            Profile(self, profile) for profile in self.app_config.profiles
        ] or [Profile(self, ProfileConfig(name="Profile 1"))]
        self.current_profile: Profile | None = None

        CTkButton(self, text="Add", command=self.add_entry).grid(row=1, column=0, padx=30, pady=5)
        CTkButton(self, text="Delete", command=self.delete_entry).grid(row=1, column=1, padx=45, pady=5)
        CTkButton(self, text="Parse", command=self.click).grid(row=1, column=2, padx=30, pady=5)

        self.selector = ProfileSelector(self, self.profiles)
        self.selector.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        CTkButton(self, text="Add Profile", command=self.selector.add).grid(row=0, column=3, padx=10, pady=10)
        CTkButton(self, text="Delete Profile", command=self.selector.remove).grid(row=0, column=4, padx=10, pady=10)

        self.cache_checkbox = CTkCheckBox(self, text="Cache")
        self.cache_checkbox.grid(row=1, column=3, padx=10, pady=5)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    # -- entries ---------------------------------------------------------------

    def add_entry(self) -> None:
        if self.current_profile:
            self.current_profile.add_entry()

    def delete_entry(self) -> None:
        if self.current_profile:
            self.current_profile.delete_entry()

    # -- parsing ---------------------------------------------------------------

    def click(self) -> None:
        if self.processing or self.current_profile is None:
            return
        self.processing = True

        profile = self.current_profile.to_config()
        use_cache = self.cache_checkbox.get() == 1

        self.progress.grid(
            row=len(self.current_profile.entries) + 2,
            column=0,
            columnspan=3,
            padx=(20, 10),
            pady=(10, 10),
            sticky="ew",
        )
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        Thread(target=self._run, args=(profile, use_cache), daemon=True).start()

    def _run(self, profile: ProfileConfig, use_cache: bool) -> None:
        """Worker thread.

        Never touches a widget: results and errors are handed back to the main
        thread with ``self.after`` (B5 — Tkinter is not thread-safe, and the previous
        version called CTkMessagebox and mutated the progress bar from here).
        """
        try:
            table = self._collect(profile, use_cache)
            path = self._write(table)
        except (ConfigError, NoPharmaciesError, ScrapeError) as error:
            self.after(0, self._failed, str(error))
        except Exception:
            logger.exception("Parsing failed")
            self.after(0, self._failed, "Unexpected error — see the log for details.")
        else:
            self.after(0, self._succeeded, path)

    def _collect(self, profile: ProfileConfig, use_cache: bool) -> PriceTable:
        cache_file = cache_path(profile.name)
        if use_cache and cache_file.exists():
            try:
                return read_table(cache_file)
            except Exception:
                logger.warning("Ignoring unreadable cache at %s", cache_file, exc_info=True)

        table = asyncio.run(scrape_profile(self.app_config.request, profile.pharmacies))
        write_table(table, cache_file)
        return table

    def _write(self, table: PriceTable) -> Path:
        settings = self.app_config.settings
        return select_exporter().export(settings, table).absolute()

    def _succeeded(self, path: Path) -> None:
        self._stop_progress()
        open_file(path)

    def _failed(self, message: str) -> None:
        self._stop_progress()
        CTkMessagebox(title="Error", message=message, icon="cancel")

    def _stop_progress(self) -> None:
        self.processing = False
        self.progress.stop()
        self.progress.grid_forget()

    # -- persistence -----------------------------------------------------------

    def current_config(self) -> AppConfig:
        return self.app_config.model_copy(
            update={"profiles": [profile.to_config() for profile in self.profiles]}
        )

    def on_closing(self) -> None:
        try:
            save_config(self.current_config(), self.config_file)
        except Exception:
            logger.exception("Could not save %s", self.config_file)
        self.destroy()
