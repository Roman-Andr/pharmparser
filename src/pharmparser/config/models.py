"""Validated configuration schema.

The on-disk format is preserved exactly as earlier versions wrote it — profiles as
a name -> {pharmacy name: url} mapping, and camelCase keys under ``settings`` — so
existing ``config.json`` files keep working. Internally everything is snake_case
and validated.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    HttpUrl,
    StringConstraints,
    ValidationInfo,
    field_validator,
    model_serializer,
    model_validator,
)
from pydantic.alias_generators import to_camel

HexColour = Annotated[str, StringConstraints(pattern=r"^[0-9A-Fa-f]{6}$")]
"""An RRGGBB colour, as openpyxl's PatternFill expects it."""


class ExportSettings(BaseModel):
    """Appearance of the generated workbook."""

    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True, extra="forbid")

    green: HexColour = "19CF1F"
    """Fill for a competitor price above the reference."""
    red: HexColour = "E81737"
    """Fill for a competitor price below the reference."""
    title: str = "Отчёт"
    file_name: str = "data.xlsx"
    col_width: int = Field(default=50, gt=0)
    cell_width: int = Field(default=15, gt=0)
    diff_width: int = Field(default=10, gt=0)

    @field_validator("green", "red", mode="before")
    @classmethod
    def _strip_hash(cls, value: Any) -> Any:
        """Accept ``#RRGGBB`` as well as ``RRGGBB``."""
        return value[1:] if isinstance(value, str) and value.startswith("#") else value

    @property
    def macro_file_name(self) -> str:
        """The macro-enabled workbook actually delivered to the user."""
        return self.file_name.removesuffix(".xlsx") + ".xlsm"


class PharmacyEntry(BaseModel):
    """One row of a profile: a display name and the pharmacy's page URL."""

    model_config = ConfigDict(extra="forbid")

    name: str = ""
    url: str = ""

    @property
    def pharmacy_id(self) -> str:
        """The numeric id tabletka.by uses, taken from the end of the URL."""
        return self.url.rstrip("/").rsplit("/", 1)[-1]

    @property
    def is_complete(self) -> bool:
        """Whether this row is filled in enough to scrape.

        Blank rows are legal — the UI creates one whenever "Add" is pressed — and are
        skipped rather than rejected (B4: an empty profile used to crash the parse).
        """
        return bool(self.name and self.url)

    @model_validator(mode="after")
    def _url_must_end_with_a_numeric_id(self) -> PharmacyEntry:
        if self.url and not self.pharmacy_id.isdigit():
            raise ValueError(
                f"pharmacy URL must end with the numeric pharmacy id, got {self.url!r}"
            )
        return self


class Profile(BaseModel):
    """A named set of pharmacies to compare."""

    model_config = ConfigDict(extra="forbid")

    name: str
    pharmacies: list[PharmacyEntry] = Field(default_factory=list)

    @property
    def scrapable(self) -> list[PharmacyEntry]:
        return [entry for entry in self.pharmacies if entry.is_complete]

    @model_validator(mode="after")
    def _pharmacy_names_must_be_unique(self) -> Profile:
        names = [entry.name for entry in self.scrapable]
        duplicates = sorted({name for name in names if names.count(name) > 1})
        if duplicates:
            raise ValueError(
                f"profile {self.name!r} has duplicate pharmacy names: {duplicates}; "
                "names label columns in the report and must be distinguishable"
            )
        return self


class RequestConfig(BaseModel):
    """Everything needed to talk to the tabletka.by price endpoint."""

    model_config = ConfigDict(extra="forbid")

    url: HttpUrl = HttpUrl("https://tabletka.by/ajax-request/reload-pharmacy-price")
    headers: dict[str, str] = Field(default_factory=dict)
    data: dict[str, str] = Field(default_factory=dict)

    @field_validator("headers")
    @classmethod
    def _require_a_cookie(cls, value: dict[str, str], info: ValidationInfo) -> dict[str, str]:
        if not any(key.lower() == "cookie" for key in value):
            raise ValueError(
                "request.headers needs a Cookie header — see the README for how to copy "
                "your session cookie out of the browser"
            )
        return value

    @property
    def cookie(self) -> str:
        return next(value for key, value in self.headers.items() if key.lower() == "cookie")

    def with_cookie(self, cookie: str) -> RequestConfig:
        headers = {key: value for key, value in self.headers.items() if key.lower() != "cookie"}
        return self.model_copy(update={"headers": {**headers, "Cookie": cookie}})


class AppConfig(BaseModel):
    """The whole of ``config.json``."""

    model_config = ConfigDict(extra="forbid")

    profiles: list[Profile] = Field(default_factory=list)
    settings: ExportSettings = Field(default_factory=ExportSettings)
    request: RequestConfig

    @field_validator("profiles", mode="before")
    @classmethod
    def _from_disk_mapping(cls, value: Any) -> Any:
        """Accept the on-disk ``{profile name: {pharmacy name: url}}`` shape.

        Profile names are carried through as given, rather than being regenerated as
        "Profile 1..N" on every save (A6).
        """
        if not isinstance(value, dict):
            return value
        return [
            {
                "name": profile_name,
                "pharmacies": [{"name": name, "url": url} for name, url in (entries or {}).items()],
            }
            for profile_name, entries in value.items()
        ]

    @model_validator(mode="after")
    def _profile_names_must_be_unique(self) -> AppConfig:
        names = [profile.name for profile in self.profiles]
        duplicates = sorted({name for name in names if names.count(name) > 1})
        if duplicates:
            raise ValueError(f"duplicate profile names: {duplicates}")
        return self

    @model_serializer
    def _to_disk(self) -> dict[str, Any]:
        """Serialise back to the original on-disk layout."""
        return {
            "profiles": {
                profile.name: {entry.name: entry.url for entry in profile.pharmacies}
                for profile in self.profiles
            },
            "settings": self.settings.model_dump(by_alias=True),
            "request": {
                "url": str(self.request.url),
                "headers": self.request.headers,
                "data": self.request.data,
            },
        }
