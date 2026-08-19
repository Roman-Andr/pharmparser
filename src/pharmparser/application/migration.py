"""One-time import of legacy config.json without copying secrets into new config."""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from uuid import uuid4

from ..config import load_config
from .credentials import Credentials, CredentialService
from .models import PharmacyProfileEntry, ProfileRecord
from .settings import DesktopSettings, SettingsService


class LegacyConfigMigrator:
    def __init__(self, settings: SettingsService, credentials: CredentialService) -> None:
        self.settings = settings
        self.credentials = credentials

    def migrate(self, legacy_path: Path) -> DesktopSettings:
        legacy = load_config(legacy_path)
        profiles: list[ProfileRecord] = []
        for old in legacy.profiles:
            entries = [
                PharmacyProfileEntry.model_validate(
                    {"id": item.pharmacy_id, "name": item.name, "url": item.url}
                )
                for item in old.scrapable
            ]
            profiles.append(
                ProfileRecord(
                    id=uuid4(),
                    name=old.name,
                    pharmacies=entries,
                    reference_pharmacy_id=entries[0].id if entries else None,
                )
            )

        credentials = Credentials(
            cookie=legacy.request.cookie,
            csrf=legacy.request.data.get("_csrf", ""),
        )
        if credentials.is_usable:
            self.credentials.update(credentials)

        current = self.settings.load()
        migrated = current.model_copy(
            update={
                "profiles": profiles,
                "selected_profile_id": str(profiles[0].id) if profiles else None,
                "green": legacy.settings.green,
                "red": legacy.settings.red,
                "file_name_template": Path(legacy.settings.file_name).stem,
                "legacy_migrated": True,
                "onboarding_complete": True,
            }
        )
        self.settings.save(migrated)
        return migrated

    @staticmethod
    def make_redacted_backup(legacy_path: Path) -> Path:
        payload = json.loads(legacy_path.read_text(encoding="utf-8"))
        request = payload.get("request", {})
        headers = request.get("headers", {})
        for key in list(headers):
            if key.casefold() in {"cookie", "authorization"}:
                headers[key] = "<redacted>"
        data = request.get("data", {})
        for key in list(data):
            if key.casefold() in {"_csrf", "csrf", "token"}:
                data[key] = "<redacted>"
        backup = legacy_path.with_name(f"{legacy_path.stem}.redacted.backup.json")
        backup.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return backup

    @staticmethod
    def remove_legacy_secrets(legacy_path: Path) -> Path:
        backup = LegacyConfigMigrator.make_redacted_backup(legacy_path)
        shutil.copyfile(backup, legacy_path)
        return backup
