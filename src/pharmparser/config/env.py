"""Environment-variable overrides.

The session cookie and CSRF token are credentials. Keeping them in ``config.json``
is still supported, but this lets them be supplied out of band instead::

    PHARMPARSER_COOKIE="PHPSESSID=...; _csrf=..." uv run pharmparser

Environment values win over the file, which is the usual precedence for secrets.
"""

from pydantic_settings import BaseSettings, SettingsConfigDict

from .models import AppConfig


class EnvOverrides(BaseSettings):
    """Credentials that may be supplied by the environment instead of the config file."""

    model_config = SettingsConfigDict(
        env_prefix="PHARMPARSER_",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    cookie: str | None = None
    """Full Cookie header value, overriding ``request.headers.Cookie``."""
    csrf: str | None = None
    """CSRF token, overriding ``request.data._csrf``."""
    file_name: str | None = None
    """Output workbook name, overriding ``settings.fileName``."""

    def apply(self, config: AppConfig) -> AppConfig:
        """Return ``config`` with any environment overrides layered on top."""
        request = config.request
        if self.cookie is not None:
            request = request.with_cookie(self.cookie)
        if self.csrf is not None:
            request = request.model_copy(update={"data": {**request.data, "_csrf": self.csrf}})

        settings = config.settings
        if self.file_name is not None:
            settings = settings.model_copy(update={"file_name": self.file_name})

        return config.model_copy(update={"request": request, "settings": settings})
