"""Platform-appropriate application paths and legacy compatibility paths."""

import os
import sys
from pathlib import Path

CONFIG_FILE_NAME = "config.json"
EXAMPLE_FILE_NAME = "config.json.example"
CACHE_FILE_NAME = "data.json"
APP_NAME = "PharmParser"


def roaming_config_dir() -> Path:
    if sys.platform == "win32":
        root = Path(os.environ.get("APPDATA", Path.home() / "AppData/Roaming"))
    elif sys.platform == "darwin":
        root = Path.home() / "Library/Application Support"
    else:
        root = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return root / APP_NAME


def local_data_dir() -> Path:
    if sys.platform == "win32":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData/Local"))
    elif sys.platform == "darwin":
        root = Path.home() / "Library/Application Support"
    else:
        root = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local/share"))
    return root / APP_NAME


def settings_path() -> Path:
    return roaming_config_dir() / "settings.json"


def history_path() -> Path:
    return local_data_dir() / "history.sqlite3"


def reports_dir() -> Path:
    return Path.home() / "Documents" / APP_NAME


def modern_log_path() -> Path:
    return local_data_dir() / "logs" / "pharmparser.log"


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
