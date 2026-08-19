"""Where configuration and cache files live."""

from pathlib import Path

CONFIG_FILE_NAME = "config.json"
EXAMPLE_FILE_NAME = "config.json.example"
CACHE_FILE_NAME = "data.json"


def project_root() -> Path:
    """The installed package's project directory.

    Config has always lived beside the executable rather than in a user config
    directory; that is preserved so existing installations keep working.
    """
    return Path(__file__).resolve().parents[3]


def config_path() -> Path:
    return Path.cwd() / CONFIG_FILE_NAME


def example_path() -> Path:
    return project_root() / EXAMPLE_FILE_NAME


def cache_path(profile_name: str) -> Path:
    """Cache file for one profile.

    Previously a single ``data.json`` was shared by every profile, so switching
    profile silently reused the wrong prices (A7).
    """
    safe = "".join(char if char.isalnum() or char in "-_" else "_" for char in profile_name)
    return Path.cwd() / f".cache-{safe}.json"
