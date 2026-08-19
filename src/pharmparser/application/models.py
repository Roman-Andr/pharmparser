"""Serializable application models; no framework or persistence imports."""

from __future__ import annotations

from datetime import datetime
from uuid import UUID, uuid4

from pydantic import BaseModel, ConfigDict, Field, HttpUrl, model_validator

from ..domain import RunStatus


class PharmacyProfileEntry(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: str
    name: str
    url: HttpUrl

    @model_validator(mode="after")
    def validate_id(self) -> PharmacyProfileEntry:
        found = str(self.url).rstrip("/").rsplit("/", 1)[-1]
        if not found.isdigit():
            raise ValueError("URL аптеки должен оканчиваться числовым идентификатором")
        if found != self.id:
            raise ValueError("идентификатор аптеки не совпадает с URL")
        return self


class ProfileRecord(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: UUID = Field(default_factory=uuid4)
    name: str
    pharmacies: list[PharmacyProfileEntry] = Field(default_factory=list)
    reference_pharmacy_id: str | None = None
    archived: bool = False

    @model_validator(mode="after")
    def validate_reference(self) -> ProfileRecord:
        ids = [entry.id for entry in self.pharmacies]
        if len(ids) != len(set(ids)):
            raise ValueError("идентификаторы аптек в профиле должны быть уникальны")
        if self.reference_pharmacy_id is not None and self.reference_pharmacy_id not in ids:
            raise ValueError("основная аптека должна входить в профиль")
        return self


class RunSummary(BaseModel):
    id: UUID
    profile_id: UUID
    parent_run_id: UUID | None = None
    status: RunStatus
    started_at: datetime
    finished_at: datetime | None = None
    reference_pharmacy_id: str
    pharmacy_count: int
    successful_pharmacies: int = 0
    product_count: int = 0
    pinned: bool = False
    report_path: str | None = None
    warning_count: int = 0


class ProgressEvent(BaseModel):
    sequence: int
    run_id: UUID
    kind: str
    pharmacy_id: str | None = None
    stage: str
    message: str
    current: int | None = None
    total: int | None = None
    timestamp: datetime
