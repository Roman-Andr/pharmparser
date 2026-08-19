"""Loopback-only FastAPI adapter over the application services."""

from __future__ import annotations

import asyncio
import secrets
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Annotated, Any
from uuid import UUID

from fastapi import APIRouter, Depends, FastAPI, Header, HTTPException, Request, Response, status
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .. import __version__
from ..application import (
    Credentials,
    CredentialService,
    HistoryRepository,
    ProfileService,
    ReportService,
    RunService,
    SettingsService,
)
from ..application.migration import LegacyConfigMigrator
from ..application.models import ProfileRecord
from ..config import RequestConfig, config_path
from ..domain import may_export
from ..platform_ import open_file, open_folder

ALLOWED_HOSTS = {"127.0.0.1", "localhost", "[::1]"}
CSP = (
    "default-src 'self'; script-src 'self'; style-src 'self'; "
    "img-src 'self' data:; connect-src 'self'; font-src 'self'; object-src 'none'; "
    "base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
)


@dataclass(slots=True)
class AppServices:
    token: str
    settings: SettingsService
    profiles: ProfileService
    credentials: CredentialService
    history: HistoryRepository
    runs: RunService
    reports: ReportService
    shutdown: asyncio.Event
    last_heartbeat: datetime
    background_tasks: set[asyncio.Task[None]]


def create_services(
    *,
    settings: SettingsService | None = None,
    credentials: CredentialService | None = None,
    history: HistoryRepository | None = None,
    token: str | None = None,
) -> AppServices:
    settings_service = settings or SettingsService()
    credential_service = credentials or CredentialService()
    history_repository = history or HistoryRepository()

    def request_config() -> RequestConfig:
        values = credential_service.get()
        return RequestConfig(
            headers={"Cookie": values.cookie},
            data={"sort": "name", "sort_type": "asc", "str": "", "_csrf": values.csrf},
        )

    return AppServices(
        token=token or secrets.token_urlsafe(32),
        settings=settings_service,
        profiles=ProfileService(settings_service),
        credentials=credential_service,
        history=history_repository,
        runs=RunService(history_repository, request_config),
        reports=ReportService(history_repository),
        shutdown=asyncio.Event(),
        last_heartbeat=datetime.now(UTC),
        background_tasks=set(),
    )


class CredentialsInput(BaseModel):
    cookie: str
    csrf: str


class ExportInput(BaseModel):
    format: str | None = None
    path: str | None = None


class PathInput(BaseModel):
    path: str


def create_app(
    services: AppServices | None = None,
    *,
    frontend_dir: Path | None = None,
    production: bool = True,
) -> FastAPI:
    state = services or create_services()
    app = FastAPI(
        title="PharmParser local API",
        version=__version__,
        docs_url=None if production else "/docs",
        redoc_url=None,
        openapi_url=None if production else "/openapi.json",
    )
    app.state.services = state

    @app.middleware("http")
    async def secure_local_request(request: Request, call_next):
        host = request.headers.get("host", "").split(":", 1)[0]
        if host not in ALLOWED_HOSTS:
            return Response("Недопустимый Host", status_code=400)
        origin = request.headers.get("origin")
        if origin is not None:
            allowed = {f"http://{request.headers.get('host')}", f"https://{request.headers.get('host')}"}
            if origin not in allowed:
                return Response("Недопустимый Origin", status_code=403)
        response = await call_next(request)
        response.headers["Content-Security-Policy"] = CSP
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Cache-Control"] = "no-store"
        return response

    async def authorize(authorization: Annotated[str | None, Header()] = None) -> None:
        expected = f"Bearer {state.token}"
        if authorization is None or not secrets.compare_digest(authorization, expected):
            raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Требуется токен локального приложения")

    api = APIRouter(prefix="/api", dependencies=[Depends(authorize)])

    @api.get("/bootstrap")
    async def bootstrap() -> dict[str, Any]:
        settings_value = state.settings.load()
        return {
            "version": __version__,
            "settings": settings_value,
            "credentials": state.credentials.status(),
            "active_run_id": state.runs.active_run_id,
            "history_size_bytes": state.history.size_bytes(),
            "legacy_config_present": config_path().is_file() and not settings_value.legacy_migrated,
        }

    @api.get("/settings")
    async def get_settings():
        return state.settings.load()

    @api.patch("/settings")
    async def patch_settings(changes: dict[str, object]):
        # Profiles have their own endpoints; accepting them here makes accidental
        # replacement of the complete roster too easy.
        changes.pop("profiles", None)
        return state.settings.update(changes)

    @api.get("/profiles")
    async def profiles(include_archived: bool = False):
        return state.profiles.list(include_archived=include_archived)

    @api.post("/profiles", status_code=201)
    async def create_profile(profile: ProfileRecord):
        return state.profiles.save(profile)

    @api.put("/profiles/{profile_id}")
    async def update_profile(profile_id: UUID, profile: ProfileRecord):
        if profile.id != profile_id:
            raise HTTPException(400, "UUID профиля в пути и теле не совпадает")
        return state.profiles.save(profile)

    @api.post("/profiles/{profile_id}/reorder")
    async def reorder_profile(profile_id: UUID, pharmacy_ids: list[str]):
        return state.profiles.reorder_pharmacies(profile_id, pharmacy_ids)

    @api.post("/profiles/{profile_id}/archive")
    async def archive_profile(profile_id: UUID, archived: bool = True):
        return state.profiles.archive(profile_id, archived)

    @api.delete("/profiles/{profile_id}", status_code=204)
    async def delete_profile(profile_id: UUID, confirm: bool = False) -> Response:
        state.profiles.delete_permanently(profile_id, confirmed=confirm)
        return Response(status_code=204)

    @api.get("/credentials/status")
    async def credential_status():
        return state.credentials.status()

    @api.post("/migration/legacy")
    async def migrate_legacy(remove_secrets: bool = False):
        legacy = config_path()
        if not legacy.is_file():
            raise HTTPException(404, "Старая конфигурация не найдена")
        migrator = LegacyConfigMigrator(state.settings, state.credentials)
        migrated = migrator.migrate(legacy)
        backup = migrator.remove_legacy_secrets(legacy) if remove_secrets else None
        return {"settings": migrated, "redacted_backup": str(backup) if backup else None}

    @api.put("/credentials")
    async def update_credentials(values: CredentialsInput):
        return state.credentials.update(Credentials(cookie=values.cookie, csrf=values.csrf))

    @api.post("/credentials/validate")
    async def validate_credentials():
        # A cheap real validation uses the selected profile's reference pharmacy.
        config = state.settings.load()
        profile_id = config.selected_profile_id
        if profile_id is None:
            raise HTTPException(400, "Сначала создайте профиль")
        profile = state.profiles.get(UUID(profile_id))
        reference = next(
            (entry for entry in profile.pharmacies if entry.id == profile.reference_pharmacy_id), None
        )
        if reference is None:
            raise HTTPException(400, "Выберите основную аптеку")

        async def validate(values: Credentials) -> None:
            from ..config import PharmacyEntry
            from ..scraping import ClientSessionFactory
            from ..scraping.parallel import ParsePool

            request = RequestConfig(headers={"Cookie": values.cookie}, data={"_csrf": values.csrf})
            async with ClientSessionFactory(request, parse_pool=ParsePool(0)) as client:
                await client.product_prices_for(PharmacyEntry(name=reference.name, url=str(reference.url)))

        return await state.credentials.validate(validate)

    @api.post("/runs", status_code=202)
    async def start_run(profile_id: UUID):
        try:
            run = await state.runs.start(state.profiles.get(profile_id))
            task = asyncio.create_task(_export_finished_run(state, run.id))
            state.background_tasks.add(task)
            task.add_done_callback(state.background_tasks.discard)
            return run
        except Exception as error:
            if error.__class__.__name__ == "ActiveRunError":
                raise HTTPException(409, str(error)) from error
            raise

    @api.post("/runs/{run_id}/cancel")
    async def cancel_run(run_id: UUID):
        return await state.runs.cancel(run_id)

    @api.post("/runs/{run_id}/retry", status_code=202)
    async def retry_run(run_id: UUID):
        try:
            return await state.runs.retry(run_id)
        except Exception as error:
            if error.__class__.__name__ == "ActiveRunError":
                raise HTTPException(409, str(error)) from error
            raise

    @api.get("/runs/{run_id}/events")
    async def run_events(run_id: UUID, request: Request):
        value = request.headers.get("last-event-id", "0")
        last_id = int(value) if value.isdigit() else 0

        async def stream():
            async for event in state.runs.events(run_id, last_id):
                if await request.is_disconnected():
                    break
                yield f"id: {event.sequence}\nevent: progress\ndata: {event.model_dump_json()}\n\n"

        return StreamingResponse(stream(), media_type="text/event-stream", headers={"X-Accel-Buffering": "no"})

    @api.get("/history")
    async def history(profile_id: UUID | None = None):
        return state.history.list_runs(profile_id)

    @api.post("/history/{run_id}/pin")
    async def pin_history(run_id: UUID, pinned: bool = True) -> dict[str, bool]:
        state.history.pin(run_id, pinned)
        return {"pinned": pinned}

    @api.delete("/history/{run_id}", status_code=204)
    async def delete_history(run_id: UUID) -> Response:
        state.history.delete(run_id)
        return Response(status_code=204)

    @api.post("/history/{run_id}/export")
    async def export_history(run_id: UUID, values: ExportInput):
        path = Path(values.path) if values.path else None
        result = state.reports.export(run_id, state.settings.load(), path=path, format_=values.format)
        return {"path": str(result)}

    @api.post("/system/open-report")
    async def open_report(values: PathInput) -> dict[str, bool]:
        open_file(Path(values.path))
        return {"opened": True}

    @api.post("/system/open-folder")
    async def open_report_folder(values: PathInput) -> dict[str, bool]:
        open_folder(Path(values.path))
        return {"opened": True}

    @api.post("/heartbeat", status_code=204)
    async def heartbeat() -> Response:
        state.last_heartbeat = datetime.now(UTC)
        return Response(status_code=204)

    @api.post("/exit", status_code=204)
    async def exit_app() -> Response:
        state.shutdown.set()
        return Response(status_code=204)

    app.include_router(api)

    root = frontend_dir or Path(__file__).resolve().parent / "static"
    if root.is_dir() and (root / "index.html").is_file():
        assets = root / "assets"
        if assets.is_dir():
            app.mount("/assets", StaticFiles(directory=assets), name="assets")

        @app.get("/{path:path}")
        async def spa(path: str):
            candidate = root / path
            if path and candidate.is_file() and candidate.resolve().is_relative_to(root.resolve()):
                return FileResponse(candidate)
            return FileResponse(root / "index.html")
    else:
        @app.get("/", response_class=HTMLResponse)
        async def missing_frontend() -> str:
            return "<h1>PharmParser</h1><p>Frontend не собран. Выполните bun run build.</p>"

    return app


async def _export_finished_run(services: AppServices, run_id: UUID) -> None:
    """Complete the one-click flow: successful scraping always produces a file."""
    run = await services.runs.wait(run_id)
    if not may_export(run.status):
        return
    try:
        await asyncio.to_thread(services.reports.export, run_id, services.settings.load())
    except Exception as error:
        services.history.add_warning(run_id, "report_failed", str(error))
