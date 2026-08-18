"""One place that configures logging for both front ends.

The GUI has nowhere to show a traceback and the packaged binary has no console at
all, so a failed run used to leave nothing behind to diagnose it. Every run now
also writes to a rotating file next to the configuration, which is the thing to
ask a user for when something goes wrong.
"""

from __future__ import annotations

import logging
import logging.handlers
from pathlib import Path

LOG_FILE_NAME = "pharmparser.log"
MAX_BYTES = 1_000_000
BACKUP_COUNT = 3

CONSOLE_FORMAT = "%(levelname)s %(name)s: %(message)s"
FILE_FORMAT = "%(asctime)s %(levelname)-8s %(name)s %(filename)s:%(lineno)d — %(message)s"

PACKAGE = __name__.split(".")[0]

logger = logging.getLogger(__name__)


class _OwnDebugOnly(logging.Filter):
    """Keep the file readable: this package's chatter, plus anyone's warnings.

    The root logger runs at DEBUG so the file can record the full story of a run,
    which would otherwise mean pages of aiohttp and asyncio internals.
    """

    def filter(self, record: logging.LogRecord) -> bool:
        return record.levelno >= logging.WARNING or record.name.split(".")[0] == PACKAGE


def log_path() -> Path:
    """Where the log file lives: beside ``config.json``, as the cache does."""
    return Path.cwd() / LOG_FILE_NAME


def configure(
    *, verbose: bool = False, console: bool = True, path: Path | None = None
) -> Path | None:
    """Set up console and rotating-file logging. Returns the log file, if any.

    Safe to call more than once: handlers this function installed are replaced
    rather than stacked, so repeated calls do not duplicate every line.
    """
    root = logging.getLogger()
    # DEBUG at the root; each handler decides how much of it to keep.
    root.setLevel(logging.DEBUG)
    for handler in list(root.handlers):
        if getattr(handler, "_pharmparser", False):
            root.removeHandler(handler)
            handler.close()

    if console:
        stream = logging.StreamHandler()
        stream.setFormatter(logging.Formatter(CONSOLE_FORMAT))
        stream.setLevel(logging.DEBUG if verbose else logging.INFO)
        stream._pharmparser = True  # type: ignore[attr-defined]
        root.addHandler(stream)

    target = path or log_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        rotating = logging.handlers.RotatingFileHandler(
            target, maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT, encoding="utf-8"
        )
    except OSError:
        # A read-only install directory must not stop the app from running.
        logger.warning("Could not open the log file at %s; logging to the console only", target)
        return None

    rotating.setFormatter(logging.Formatter(FILE_FORMAT))
    rotating.setLevel(logging.DEBUG)
    rotating.addFilter(_OwnDebugOnly())
    rotating._pharmparser = True  # type: ignore[attr-defined]
    root.addHandler(rotating)
    return target
