"""The window: widgets, layout and callbacks. No business logic.

Everything the app *does* lives in :class:`~pharmparser.ui.controller.Controller`;
this module only draws it and marshals results back onto the main thread.
"""

from __future__ import annotations

import logging
from pathlib import Path
from threading import Thread

from CTkMessagebox import CTkMessagebox
from customtkinter import CTk, CTkButton, CTkCheckBox, CTkProgressBar

from ..config import AppConfig, ConfigError, config_path
from ..config import Profile as ProfileConfig
from ..controller import Controller
from ..platform_ import open_file
from ..scraping import NoPharmaciesError, ScrapeError
from .profile import Profile
from .profile_selector import ProfileSelector

logger = logging.getLogger(__name__)

EXPECTED_FAILURES = (ConfigError, NoPharmaciesError, ScrapeError)
"""Errors with a message worth showing the user verbatim."""


class App(CTk):
    """The main window.

    ``self.config`` is deliberately *not* used for application state: it is
    ``tkinter.Misc.config``, the alias of ``configure``, and shadowing it made any
    call to it raise ``TypeError: 'str' object is not callable`` (B3).
    """

    processing: bool = False

    def __init__(
        self,
        config: AppConfig | None = None,
        config_file: Path | None = None,
        controller: Controller | None = None,
    ):
        super().__init__()

        if controller is not None:
            self.controller = controller
        elif config is not None:
            self.controller = Controller(config, config_file or config_path())
        else:
            self.controller = Controller.load(config_file)

        self.geometry("1100x600")
        self.title("PharmParser")

        self.progress: CTkProgressBar = CTkProgressBar(self)
        self.profiles: list[Profile] = [Profile(self, profile) for profile in self.controller.profiles]
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
            path = self.controller.run(profile, use_cache)
        except EXPECTED_FAILURES as error:
            self.after(0, self._failed, str(error))
        except Exception as error:
            logger.exception("Parsing failed")
            self.after(0, self._failed, f"{type(error).__name__}: {error}")
        else:
            self.after(0, self._succeeded, path)

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

    def on_closing(self) -> None:
        try:
            self.controller.save([profile.to_config() for profile in self.profiles])
        except Exception as error:
            logger.exception("Could not save %s", self.controller.config_file)
            CTkMessagebox(title="Could not save", message=str(error), icon="warning")
        self.destroy()
