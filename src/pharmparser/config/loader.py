"""Reading and writing ``config.json``."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from tempfile import NamedTemporaryFile

from pydantic import ValidationError

from .env import EnvOverrides
from .models import AppConfig
from .paths import config_path

logger = logging.getLogger(__name__)


class ConfigError(Exception):
    """Raised when the config file is missing or invalid, with a message for the user."""


def _describe(error: ValidationError, path: Path) -> str:
    lines = [f"{path} is not valid:"]
    for detail in error.errors():
        location = ".".join(str(part) for part in detail["loc"]) or "(root)"
        lines.append(f"  - {location}: {detail['msg']}")
    return "\n".join(lines)


def load_config(path: Path | None = None) -> AppConfig:
    """Load, validate and apply environment overrides to the configuration.

    Raises :class:`ConfigError` with an actionable message rather than letting a
    bare ``FileNotFoundError`` or ``TypeError`` escape (A6).
    """
    path = path or config_path()
    if not path.exists():
        raise ConfigError(
            f"{path} not found. Copy config.json.example to {path.name} and fill in "
            "your session cookie — see the README."
        )
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise ConfigError(f"{path} is not valid JSON: {error}") from error

    try:
        config = AppConfig.model_validate(raw)
    except ValidationError as error:
        raise ConfigError(_describe(error, path)) from error

    return EnvOverrides().apply(config)


def save_config(config: AppConfig, path: Path | None = None) -> None:
    """Write the configuration atomically.

    The previous implementation truncated the file before writing, so an interrupted
    save destroyed the user's session cookie along with everything else (A6).
    """
    path = path or config_path()
    payload = json.dumps(config.model_dump(), ensure_ascii=False, indent=2)

    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path: str | None = None
    try:
        with NamedTemporaryFile(
            "w", encoding="utf-8", dir=path.parent, prefix=f".{path.name}.", suffix=".tmp", delete=False
        ) as tmp:
            tmp_path = tmp.name
            tmp.write(payload)
            tmp.flush()
            os.fsync(tmp.fileno())
        os.replace(tmp_path, path)
        tmp_path = None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
