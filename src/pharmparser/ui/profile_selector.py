from __future__ import annotations

from typing import Any

from customtkinter import CTkSegmentedButton

from ..config import Profile as ProfileConfig
from ..controller import Controller
from .profile import Profile


class ProfileSelector(CTkSegmentedButton):
    """Switches between profiles, keyed by the profile's own name.

    Names used to be regenerated as "Profile 1..N" on every save, so a name the
    user chose could not survive a restart (A6).
    """

    def __init__(self, app: Any, profiles: list[Profile], **kwargs: Any) -> None:
        super().__init__(app, **kwargs)
        self.app = app
        self.profiles = profiles
        self._refresh()
        first = self.profiles[0].name
        self.set(first)
        self.change_profile(first)

    def _names(self) -> list[str]:
        return [profile.name for profile in self.profiles]

    def _refresh(self) -> None:
        self.configure(values=self._names(), command=self.change_profile)

    def change_profile(self, name: str) -> None:
        for profile in self.profiles:
            profile.hide()
        selected = next((p for p in self.profiles if p.name == name), None)
        if selected is None:
            return
        self.app.current_profile = selected
        selected.display()

    def add(self) -> None:
        name = Controller.next_profile_name(self._names())
        self.profiles.append(Profile(self.app, ProfileConfig(name=name)))
        self._refresh()
        self.set(name)
        self.change_profile(name)

    def remove(self) -> None:
        if self.app.current_profile and len(self.profiles) > 1:
            self.app.current_profile.hide()
            self.profiles.remove(self.app.current_profile)
            self._refresh()
            name = self.profiles[-1].name
            self.set(name)
            self.change_profile(name)
