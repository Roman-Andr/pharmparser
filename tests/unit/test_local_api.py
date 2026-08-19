from pathlib import Path

import httpx
from fastapi import FastAPI

from pharmparser.application import CredentialService, HistoryRepository, SettingsService
from pharmparser.web import create_app, create_services


def app(tmp_path: Path, token: str = "test-token") -> FastAPI:
    services = create_services(
        settings=SettingsService(tmp_path / "settings.json"),
        credentials=CredentialService(fallback_path=tmp_path / "credentials.json"),
        history=HistoryRepository(tmp_path / "history.sqlite3"),
        token=token,
    )
    return create_app(services)


async def test_api_requires_bearer_token_and_never_returns_secret(tmp_path: Path) -> None:
    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app(tmp_path)), base_url="http://127.0.0.1"
    ) as api:
        assert (await api.get("/api/bootstrap")).status_code == 401
        await api.put(
            "/api/credentials",
            headers={"Authorization": "Bearer test-token"},
            json={"cookie": "session=secret-cookie", "csrf": "secret-csrf"},
        )
        response = await api.get("/api/bootstrap", headers={"Authorization": "Bearer test-token"})
        assert response.status_code == 200
        assert "secret-cookie" not in response.text
        assert "secret-csrf" not in response.text


async def test_api_rejects_host_and_cross_origin_requests(tmp_path: Path) -> None:
    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app(tmp_path)), base_url="http://127.0.0.1"
    ) as api:
        headers = {"Authorization": "Bearer test-token"}
        response = await api.get("/api/bootstrap", headers={**headers, "Host": "attacker.example"})
        assert response.status_code == 400
        response = await api.get(
            "/api/bootstrap", headers={**headers, "Origin": "https://attacker.example"}
        )
        assert response.status_code == 403


async def test_security_headers_are_present(tmp_path: Path) -> None:
    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app(tmp_path)), base_url="http://127.0.0.1"
    ) as api:
        response = await api.get(
            "/api/bootstrap", headers={"Authorization": "Bearer test-token"}
        )
        assert response.headers["content-security-policy"].startswith("default-src 'self'")
        assert "Access-Control-Allow-Origin" not in response.headers
