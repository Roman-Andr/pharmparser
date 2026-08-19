"""Profile use cases with stable UUIDs and explicit reference pharmacies."""

from __future__ import annotations

import builtins
from uuid import UUID

from .models import ProfileRecord
from .settings import SettingsService


class ProfileService:
    def __init__(self, settings: SettingsService) -> None:
        self.settings = settings

    def list(self, *, include_archived: bool = False) -> builtins.list[ProfileRecord]:
        profiles = self.settings.load().profiles
        return profiles if include_archived else [profile for profile in profiles if not profile.archived]

    def get(self, profile_id: UUID) -> ProfileRecord:
        for profile in self.settings.load().profiles:
            if profile.id == profile_id:
                return profile
        raise KeyError(str(profile_id))

    def save(self, profile: ProfileRecord) -> ProfileRecord:
        config = self.settings.load()
        profiles = list(config.profiles)
        for index, existing in enumerate(profiles):
            if existing.id == profile.id:
                profiles[index] = profile
                break
        else:
            profiles.append(profile)
        self.settings.save(config.model_copy(update={"profiles": profiles}))
        return profile

    def archive(self, profile_id: UUID, archived: bool = True) -> ProfileRecord:
        profile = self.get(profile_id).model_copy(update={"archived": archived})
        return self.save(profile)

    def reorder_pharmacies(self, profile_id: UUID, pharmacy_ids: builtins.list[str]) -> ProfileRecord:
        profile = self.get(profile_id)
        current = {entry.id: entry for entry in profile.pharmacies}
        if len(pharmacy_ids) != len(set(pharmacy_ids)) or set(pharmacy_ids) != set(current):
            raise ValueError("порядок должен содержать каждую аптеку ровно один раз")
        return self.save(profile.model_copy(update={"pharmacies": [current[item] for item in pharmacy_ids]}))

    def delete_permanently(self, profile_id: UUID, *, confirmed: bool = False) -> None:
        if not confirmed:
            raise ValueError("для постоянного удаления требуется подтверждение")
        config = self.settings.load()
        profiles = [profile for profile in config.profiles if profile.id != profile_id]
        self.settings.save(config.model_copy(update={"profiles": profiles}))
