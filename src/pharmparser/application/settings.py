"""Non-secret desktop settings stored in roaming application data."""

from __future__ import annotations

import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from ..config.paths import reports_dir, settings_path
from .models import ProfileRecord


class DesktopSettings(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: int = 1
    onboarding_complete: bool = False
    legacy_migrated: bool = False
    theme: Literal["light", "dark"] = "light"
    output_directory: str = Field(default_factory=lambda: str(reports_dir()))
    file_name_template: str = "{profile}_{date}_{time}"
    report_format: Literal["xlsm", "xlsx"] = "xlsm"
    green: str = "19CF1F"
    red: str = "E81737"
    retention: int | None = Field(default=50, ge=10, le=500)
    check_updates: bool = True
    window_width: int = Field(default=1180, ge=960)
    window_height: int = Field(default=760, ge=640)
    window_x: int | None = None
    window_y: int | None = None
    profiles: list[ProfileRecord] = Field(default_factory=list)
    selected_profile_id: str | None = None


class SettingsService:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or settings_path()

    def load(self) -> DesktopSettings:
        if not self.path.exists():
            return DesktopSettings()
        return DesktopSettings.model_validate_json(self.path.read_text(encoding="utf-8"))

    def save(self, settings: DesktopSettings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = settings.model_dump_json(indent=2)
        temporary: str | None = None
        try:
            with NamedTemporaryFile(
                "w", encoding="utf-8", dir=self.path.parent, prefix=f".{self.path.name}.", delete=False
            ) as handle:
                temporary = handle.name
                handle.write(payload)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, self.path)
            temporary = None
        finally:
            if temporary:
                Path(temporary).unlink(missing_ok=True)

    def update(self, changes: dict[str, object]) -> DesktopSettings:
        updated = self.load().model_copy(update=changes)
        # model_copy does not validate updates in Pydantic; round-trip once here.
        validated = DesktopSettings.model_validate(updated.model_dump())
        self.save(validated)
        return validated
