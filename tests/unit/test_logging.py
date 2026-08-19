"""Tests for logging setup.

The GUI has nowhere to show a traceback and the packaged binary has no console, so
the log file is the only record of a failed run — it has to actually get written.
"""

from __future__ import annotations

import logging
from pathlib import Path

import pytest

from pharmparser.logging_ import BACKUP_COUNT, LOG_FILE_NAME, configure, log_path, redact


@pytest.fixture(autouse=True)
def _restore_root_logger():
    root = logging.getLogger()
    handlers, level = list(root.handlers), root.level
    yield
    for handler in list(root.handlers):
        root.removeHandler(handler)
    for handler in handlers:
        root.addHandler(handler)
    root.setLevel(level)


def test_writes_to_the_log_file(tmp_path: Path) -> None:
    target = configure(path=tmp_path / LOG_FILE_NAME, console=False)
    assert target == tmp_path / LOG_FILE_NAME
    logging.getLogger("pharmparser.test").info("hello")
    assert "hello" in target.read_text(encoding="utf-8")


def test_the_file_records_more_detail_than_the_console(tmp_path: Path) -> None:
    """Debug lines reach the file even when the console is at INFO."""
    target = configure(path=tmp_path / LOG_FILE_NAME, console=True, verbose=False)
    assert target is not None
    logging.getLogger("pharmparser.test").debug("quiet detail")
    assert "quiet detail" in target.read_text(encoding="utf-8")


def test_third_party_chatter_stays_out_of_the_file(tmp_path: Path) -> None:
    target = configure(path=tmp_path / LOG_FILE_NAME, console=False)
    assert target is not None
    logging.getLogger("aiohttp.client").debug("connection pool detail")
    logging.getLogger("aiohttp.client").warning("something worth knowing")
    written = target.read_text(encoding="utf-8")
    assert "connection pool detail" not in written
    assert "something worth knowing" in written


def test_repeated_configuration_does_not_duplicate_lines(tmp_path: Path) -> None:
    target = tmp_path / LOG_FILE_NAME
    for _ in range(3):
        configure(path=target, console=False)
    logging.getLogger("pharmparser.test").info("once")
    assert target.read_text(encoding="utf-8").count("once") == 1


def test_configuration_leaves_foreign_handlers_alone(tmp_path: Path) -> None:
    outsider = logging.NullHandler()
    logging.getLogger().addHandler(outsider)
    configure(path=tmp_path / LOG_FILE_NAME, console=False)
    assert outsider in logging.getLogger().handlers


def test_an_unwritable_location_is_survivable(tmp_path: Path, caplog: pytest.LogCaptureFixture) -> None:
    """A read-only install directory must not stop the app from running."""
    blocked = tmp_path / "file"
    blocked.write_text("not a directory", encoding="utf-8")
    with caplog.at_level("WARNING"):
        assert configure(path=blocked / "sub" / LOG_FILE_NAME, console=False) is None
    assert "logging to the console only" in caplog.text


def test_the_log_lives_beside_the_configuration() -> None:
    assert log_path() == Path.cwd() / LOG_FILE_NAME


def test_backups_are_kept() -> None:
    assert BACKUP_COUNT >= 1


def test_secret_values_are_redacted_from_mapping_and_header_text() -> None:
    value = "{'Cookie': 'session=abc; region=1', '_csrf': 'token-value', 'safe': 'visible'}"
    cleaned = redact(value)
    assert "session=abc" not in cleaned
    assert "token-value" not in cleaned
    assert "visible" in cleaned
