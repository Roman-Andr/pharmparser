"""Checking GitHub Releases for a newer build, and installing it.

Deliberately *not* a silent auto-updater. The app asks before replacing itself: a
desktop program that swaps its own binary without being told to is hard to trust and
harder to debug when a release goes wrong.

What is verified before anything is executed:

* the release is a real, published, non-prerelease tag on this repository;
* every URL is https and points at a GitHub host;
* the downloaded file's SHA-256 matches the ``SHA256SUMS`` asset published beside it.

What is *not* verified is the publisher: the binaries are unsigned, so Windows
SmartScreen will warn on first run and the trust root is GitHub's account security,
not a code-signing certificate. Signing needs a paid certificate; until there is one,
this is the honest limit of what the check buys.
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse

from . import __version__

logger = logging.getLogger(__name__)

REPOSITORY = "Roman-Andr/pharmparser"
RELEASES_API = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
RELEASES_PAGE = f"https://github.com/{REPOSITORY}/releases/latest"

CHECKSUMS_ASSET = "SHA256SUMS"
TIMEOUT_SECONDS = 15
DOWNLOAD_TIMEOUT_SECONDS = 300

ALLOWED_HOSTS = frozenset(
    {"api.github.com", "github.com", "objects.githubusercontent.com", "release-assets.githubusercontent.com"}
)

SUPERSEDED_SUFFIX = ".old"
"""A running executable cannot be deleted on Windows, but it can be renamed."""

_VERSION = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)$")


class UpdateError(RuntimeError):
    """An update could not be checked for, downloaded, or installed."""


@dataclass(frozen=True, slots=True)
class Release:
    """A published release and the binary it offers for this platform."""

    version: str
    tag: str
    page_url: str
    asset_name: str
    asset_url: str
    checksums_url: str | None
    notes: str = ""


def parse_version(text: str) -> tuple[int, int, int] | None:
    """``"v1.2.3"`` -> ``(1, 2, 3)``; anything else -> None.

    Strict on purpose: a tag this cannot read is a tag this will not update to.
    """
    match = _VERSION.match(text.strip())
    return (int(match[1]), int(match[2]), int(match[3])) if match else None


def is_newer(candidate: str, current: str) -> bool:
    """Whether ``candidate`` is a release strictly newer than ``current``."""
    new, old = parse_version(candidate), parse_version(current)
    return new is not None and old is not None and new > old


def current_version() -> str:
    return __version__


def running_frozen() -> bool:
    """Whether this is the PyInstaller binary rather than a source checkout."""
    return bool(getattr(sys, "frozen", False))


def executable_path() -> Path:
    return Path(sys.executable).resolve()


def asset_name_for(version: str) -> str:
    """The release asset this platform installs."""
    if sys.platform == "win32":
        return f"pharmparser-{version}-windows-x64.exe"
    if sys.platform == "darwin":
        return f"pharmparser-{version}-macos-x64"
    return f"pharmparser-{version}-linux-x64"


def _check_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in ALLOWED_HOSTS:
        raise UpdateError(f"refusing a release URL outside GitHub: {url!r}")
    return url


def _fetch(url: str, timeout: int) -> bytes:
    request = urllib.request.Request(
        _check_url(url),
        headers={"Accept": "application/vnd.github+json", "User-Agent": f"pharmparser/{current_version()}"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def latest_release(timeout: int = TIMEOUT_SECONDS) -> Release | None:
    """The newest published release, or None when there is nothing usable."""
    try:
        payload = json.loads(_fetch(RELEASES_API, timeout))
    except UpdateError:
        raise
    except Exception as error:  # network down, rate limited, malformed JSON
        raise UpdateError(f"could not reach GitHub Releases: {error}") from error

    if payload.get("draft") or payload.get("prerelease"):
        return None

    tag = str(payload.get("tag_name") or "")
    version = tag.removeprefix("v")
    if parse_version(tag) is None:
        logger.debug("Ignoring release with an unreadable tag: %r", tag)
        return None

    assets = {asset.get("name"): asset for asset in payload.get("assets") or []}
    wanted = asset_name_for(version)
    if wanted not in assets:
        logger.debug("Release %s has no asset named %s", tag, wanted)
        return None

    checksums = assets.get(CHECKSUMS_ASSET)
    return Release(
        version=version,
        tag=tag,
        page_url=payload.get("html_url") or RELEASES_PAGE,
        asset_name=wanted,
        asset_url=assets[wanted]["browser_download_url"],
        checksums_url=checksums["browser_download_url"] if checksums else None,
        notes=payload.get("body") or "",
    )


def available_update() -> Release | None:
    """The newest release, if it is newer than what is running."""
    release = latest_release()
    if release is None or not is_newer(release.version, current_version()):
        return None
    return release


def _expected_digest(release: Release) -> str:
    """The published SHA-256 for this release's asset."""
    if release.checksums_url is None:
        raise UpdateError(
            f"release {release.tag} publishes no {CHECKSUMS_ASSET}; refusing to install it unverified"
        )
    body = _fetch(release.checksums_url, TIMEOUT_SECONDS).decode("utf-8", "replace")
    for line in body.splitlines():
        parts = line.split()
        # "<sha256>  <name>", the sha256sum format, with an optional leading "*".
        if len(parts) == 2 and parts[1].lstrip("*") == release.asset_name:
            return parts[0].lower()
    raise UpdateError(f"{CHECKSUMS_ASSET} lists no entry for {release.asset_name}")


def download(release: Release, into: Path | None = None) -> Path:
    """Fetch the release binary and verify it against the published checksum."""
    expected = _expected_digest(release)
    directory = into or Path(tempfile.mkdtemp(prefix="pharmparser-update-"))
    directory.mkdir(parents=True, exist_ok=True)
    target = directory / release.asset_name

    payload = _fetch(release.asset_url, DOWNLOAD_TIMEOUT_SECONDS)
    actual = hashlib.sha256(payload).hexdigest()
    if actual != expected:
        raise UpdateError(
            f"checksum mismatch for {release.asset_name}: expected {expected}, got {actual}"
        )

    target.write_bytes(payload)
    target.chmod(0o755)
    logger.info("Downloaded %s (%d bytes, sha256 verified)", release.asset_name, len(payload))
    return target


def install(downloaded: Path, executable: Path | None = None) -> Path:
    """Put ``downloaded`` in place of the running binary.

    The running file is renamed rather than deleted, because Windows will not let a
    running image be removed. :func:`clean_superseded` clears it on the next start.
    """
    if not running_frozen():
        raise UpdateError("only the packaged application can update itself in place")

    current = executable or executable_path()
    superseded = current.with_name(current.name + SUPERSEDED_SUFFIX)
    superseded.unlink(missing_ok=True)

    os.replace(current, superseded)
    try:
        os.replace(downloaded, current)
    except OSError:
        os.replace(superseded, current)  # put the working binary back
        raise
    logger.info("Installed %s over %s", downloaded.name, current)
    return current


def clean_superseded(executable: Path | None = None) -> None:
    """Remove the previous binary left behind by an update."""
    current = executable or executable_path()
    superseded = current.with_name(current.name + SUPERSEDED_SUFFIX)
    try:
        superseded.unlink(missing_ok=True)
    except OSError as error:  # still locked; the next start will get it
        logger.debug("Could not remove %s: %s", superseded, error)


def restart(executable: Path | None = None) -> None:
    """Launch the freshly installed binary and leave."""
    target = executable or executable_path()
    subprocess.Popen([str(target), *sys.argv[1:]], close_fds=True)
    sys.exit(0)
