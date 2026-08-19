"""Run states and policies shared by desktop, API and CLI."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from enum import StrEnum


class RunStatus(StrEnum):
    QUEUED = "queued"
    RUNNING = "running"
    PARTIAL = "partial"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"

    @property
    def terminal(self) -> bool:
        return self in {
            RunStatus.PARTIAL,
            RunStatus.COMPLETED,
            RunStatus.FAILED,
            RunStatus.CANCELLED,
        }


RETRY_WINDOW = timedelta(minutes=30)


def final_status(*, reference_succeeded: bool, successful_competitors: int, cancelled: bool = False) -> RunStatus:
    if cancelled:
        return RunStatus.CANCELLED
    if not reference_succeeded or successful_competitors < 1:
        return RunStatus.FAILED
    return RunStatus.PARTIAL


def may_export(status: RunStatus) -> bool:
    return status in {RunStatus.COMPLETED, RunStatus.PARTIAL}


def may_retry(finished_at: datetime, now: datetime | None = None) -> bool:
    current = now or datetime.now(UTC)
    if finished_at.tzinfo is None:
        finished_at = finished_at.replace(tzinfo=UTC)
    return current - finished_at <= RETRY_WINDOW
