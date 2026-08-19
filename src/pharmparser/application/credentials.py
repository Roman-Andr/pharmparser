"""Secret storage isolated from settings, history and API response models."""

from __future__ import annotations

import os
from collections.abc import Awaitable, Callable
from contextlib import suppress
from pathlib import Path

from pydantic import BaseModel

from ..config.paths import local_data_dir

SERVICE = "PharmParser"
COOKIE_KEY = "tabletka-cookie"
CSRF_KEY = "tabletka-csrf"


class Credentials(BaseModel):
    cookie: str
    csrf: str

    def masked(self) -> str:
        return "••••••••" if self.cookie else ""

    @property
    def is_usable(self) -> bool:
        placeholders = {"redacted", "<redacted>", "..."}
        cookie = self.cookie.strip().casefold()
        csrf = self.csrf.strip().casefold()
        return (
            "=" in self.cookie
            and cookie not in placeholders
            and "redacted" not in cookie
            and bool(csrf)
            and csrf not in placeholders
            and "redacted" not in csrf
        )


class CredentialStatus(BaseModel):
    configured: bool
    backend: str
    masked_cookie: str | None = None
    warning: str | None = None


class _Backend:
    name = "unknown"
    warning: str | None = None

    def get(self) -> Credentials | None:
        raise NotImplementedError

    def set(self, credentials: Credentials) -> None:
        raise NotImplementedError

    def delete(self) -> None:
        raise NotImplementedError


class _KeyringBackend(_Backend):
    name = "system-keyring"

    def __init__(self) -> None:
        import keyring

        self.keyring = keyring

    def get(self) -> Credentials | None:
        cookie = self.keyring.get_password(SERVICE, COOKIE_KEY)
        csrf = self.keyring.get_password(SERVICE, CSRF_KEY)
        return Credentials(cookie=cookie, csrf=csrf) if cookie and csrf else None

    def set(self, credentials: Credentials) -> None:
        self.keyring.set_password(SERVICE, COOKIE_KEY, credentials.cookie)
        self.keyring.set_password(SERVICE, CSRF_KEY, credentials.csrf)

    def delete(self) -> None:
        for key in (COOKIE_KEY, CSRF_KEY):
            with suppress(self.keyring.errors.PasswordDeleteError):
                self.keyring.delete_password(SERVICE, key)


class _EnvironmentBackend(_Backend):
    name = "environment"
    warning = "Учетные данные читаются из переменных окружения и не могут быть изменены приложением."

    def get(self) -> Credentials | None:
        cookie = os.environ.get("PHARMPARSER_COOKIE")
        csrf = os.environ.get("PHARMPARSER_CSRF")
        return Credentials(cookie=cookie, csrf=csrf) if cookie and csrf else None

    def set(self, credentials: Credentials) -> None:
        raise RuntimeError(self.warning)

    def delete(self) -> None:
        raise RuntimeError(self.warning)


class _ProtectedFileBackend(_Backend):
    name = "protected-file"
    warning = (
        "Системное хранилище секретов недоступно. Учетные данные сохранены в локальном файле "
        "с правами только для текущего пользователя."
    )

    def __init__(self, path: Path) -> None:
        self.path = path

    def get(self) -> Credentials | None:
        if not self.path.exists():
            return None
        return Credentials.model_validate_json(self.path.read_text(encoding="utf-8"))

    def set(self, credentials: Credentials) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(credentials.model_dump_json(), encoding="utf-8")
        self.path.chmod(0o600)

    def delete(self) -> None:
        self.path.unlink(missing_ok=True)


def _select_backend(fallback_path: Path) -> _Backend:
    if os.environ.get("PHARMPARSER_COOKIE") and os.environ.get("PHARMPARSER_CSRF"):
        return _EnvironmentBackend()
    try:
        backend = _KeyringBackend()
        # A fail backend imports successfully but cannot persist anything.
        if backend.keyring.get_keyring().priority > 0:
            return backend
    except Exception:
        pass
    return _ProtectedFileBackend(fallback_path)


class CredentialService:
    def __init__(self, backend: _Backend | None = None, fallback_path: Path | None = None) -> None:
        path = fallback_path or local_data_dir() / "credentials.json"
        self._fallback_path = path
        self._backend = backend or (_ProtectedFileBackend(path) if fallback_path is not None else _select_backend(path))

    def status(self) -> CredentialStatus:
        try:
            credentials = self._backend.get()
        except Exception:
            credentials = None
        configured = credentials is not None and credentials.is_usable
        return CredentialStatus(
            configured=configured,
            backend=self._backend.name,
            masked_cookie=credentials.masked() if configured and credentials else None,
            warning=self._backend.warning,
        )

    def get(self) -> Credentials:
        credentials = self._backend.get()
        if credentials is None or not credentials.is_usable:
            raise RuntimeError("учетные данные не настроены")
        return credentials

    def update(self, credentials: Credentials) -> CredentialStatus:
        if not credentials.is_usable:
            raise ValueError("Cookie или CSRF пусты либо содержат тестовые заглушки")
        try:
            self._backend.set(credentials)
            read_back = self._backend.get()
        except Exception:
            if not isinstance(self._backend, _KeyringBackend):
                raise
            self._backend = _ProtectedFileBackend(self._fallback_path)
            self._backend.set(credentials)
            read_back = self._backend.get()
        if read_back != credentials:
            raise RuntimeError("не удалось проверить запись учетных данных")
        return self.status()

    async def validate(self, validator: Callable[[Credentials], Awaitable[None]]) -> CredentialStatus:
        credentials = self.get()
        await validator(credentials)
        return self.status()

    def delete(self) -> None:
        self._backend.delete()
