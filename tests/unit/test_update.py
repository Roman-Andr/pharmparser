"""The GitHub Releases update check, against a local stand-in for the API.

Nothing here talks to github.com. The module's own host allow-list is relaxed for
the fixture's loopback address, which is also the point of the two tests that leave
it in place: a release must not be able to send the app somewhere else.
"""

from __future__ import annotations

import hashlib
import json
import threading
from collections.abc import Iterator
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import ClassVar

import pytest

from pharmparser import update
from pharmparser.update import UpdateError

BINARY = b"MZ\x90\x00 pretend this is a windows executable"
DIGEST = hashlib.sha256(BINARY).hexdigest()
ASSET = "pharmparser-9.9.9-linux-x64"


class Api(BaseHTTPRequestHandler):
    payloads: ClassVar[dict[str, tuple[int, bytes]]] = {}

    def do_GET(self) -> None:
        status, body = self.payloads.get(self.path, (404, b"not found"))
        self.send_response(status)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args: object) -> None:
        pass


@pytest.fixture
def api(monkeypatch: pytest.MonkeyPatch) -> Iterator[str]:
    server = HTTPServer(("127.0.0.1", 0), Api)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base = f"http://127.0.0.1:{server.server_port}"

    # The fixture is loopback plain HTTP, so relax the two checks that would reject it.
    monkeypatch.setattr(update, "_check_url", lambda url: url)
    monkeypatch.setattr(update, "asset_name_for", lambda version: f"pharmparser-{version}-linux-x64")
    Api.payloads = {}
    try:
        yield base
    finally:
        server.shutdown()
        server.server_close()


def publish(base: str, monkeypatch: pytest.MonkeyPatch, **overrides: object) -> None:
    release = {
        "tag_name": "v9.9.9",
        "html_url": "https://github.com/Roman-Andr/pharmparser/releases/tag/v9.9.9",
        "draft": False,
        "prerelease": False,
        "body": "notes",
        "assets": [
            {"name": ASSET, "browser_download_url": f"{base}/asset"},
            {"name": "SHA256SUMS", "browser_download_url": f"{base}/sums"},
        ],
    }
    release.update(overrides)
    Api.payloads = {
        "/latest": (200, json.dumps(release).encode()),
        "/asset": (200, BINARY),
        "/sums": (200, f"{DIGEST}  {ASSET}\n".encode()),
    }
    monkeypatch.setattr(update, "RELEASES_API", f"{base}/latest")


# ---- version comparison ------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [("v1.2.3", (1, 2, 3)), ("1.2.3", (1, 2, 3)), (" v0.1.0 ", (0, 1, 0)),
     ("v1.2", None), ("v1.2.3-rc1", None), ("latest", None), ("", None)],
)
def test_parse_version(text: str, expected: tuple[int, int, int] | None) -> None:
    assert update.parse_version(text) == expected


@pytest.mark.parametrize(
    ("candidate", "current", "expected"),
    [("0.2.0", "0.1.0", True), ("1.0.0", "0.9.9", True), ("0.1.1", "0.1.0", True),
     ("0.1.0", "0.1.0", False), ("0.1.0", "0.2.0", False), ("nightly", "0.1.0", False),
     ("0.2.0", "0.0.0+unknown", False)],
)
def test_is_newer(candidate: str, current: str, expected: bool) -> None:
    assert update.is_newer(candidate, current) is expected


# ---- discovering a release ---------------------------------------------


def test_finds_the_asset_for_this_platform(api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    publish(api, monkeypatch)
    release = update.latest_release()
    assert release is not None
    assert (release.version, release.tag, release.asset_name) == ("9.9.9", "v9.9.9", ASSET)
    assert release.checksums_url is not None


@pytest.mark.parametrize("flag", ["draft", "prerelease"])
def test_drafts_and_prereleases_are_ignored(flag: str, api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    """The updater follows finished releases only."""
    publish(api, monkeypatch, **{flag: True})
    assert update.latest_release() is None


def test_a_tag_that_cannot_be_read_is_ignored(api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    publish(api, monkeypatch, tag_name="nightly")
    assert update.latest_release() is None


def test_a_release_without_an_asset_for_us_is_ignored(api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    publish(api, monkeypatch, assets=[{"name": "source.zip", "browser_download_url": f"{api}/asset"}])
    assert update.latest_release() is None


def test_an_unreachable_api_is_an_error_not_a_crash(api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(update, "RELEASES_API", f"{api}/missing")
    with pytest.raises(UpdateError, match="could not reach"):
        update.latest_release()


def test_no_update_is_offered_for_the_running_version(api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    publish(api, monkeypatch, tag_name="v9.9.9")
    monkeypatch.setattr(update, "current_version", lambda: "9.9.9")
    assert update.available_update() is None


def test_a_newer_release_is_offered(api: str, monkeypatch: pytest.MonkeyPatch) -> None:
    publish(api, monkeypatch)
    monkeypatch.setattr(update, "current_version", lambda: "0.1.0")
    assert update.available_update() is not None


# ---- refusing what should be refused -----------------------------------


@pytest.mark.parametrize(
    "url",
    ["http://github.com/x", "https://evil.example.com/x", "file:///etc/passwd",
     "https://github.com.evil.example/x"],
)
def test_urls_outside_github_are_refused(url: str) -> None:
    """A compromised release must not be able to point the app elsewhere."""
    with pytest.raises(UpdateError, match="outside GitHub"):
        update._check_url(url)


@pytest.mark.parametrize(
    "url",
    ["https://api.github.com/x", "https://github.com/x", "https://objects.githubusercontent.com/x"],
)
def test_github_urls_are_allowed(url: str) -> None:
    assert update._check_url(url) == url


def test_a_release_without_checksums_is_refused(api: str, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    publish(api, monkeypatch)
    release = update.latest_release()
    assert release is not None
    with pytest.raises(UpdateError, match="refusing to install it unverified"):
        update.download(release.model_copy(update={"checksums_url": None}), tmp_path)


def test_a_checksum_mismatch_is_refused(api: str, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """The one thing standing between a tampered download and being executed."""
    publish(api, monkeypatch)
    Api.payloads["/sums"] = (200, f"{'0' * 64}  {ASSET}\n".encode())
    release = update.latest_release()
    assert release is not None
    with pytest.raises(UpdateError, match="checksum mismatch"):
        update.download(release, tmp_path)
    assert list(tmp_path.iterdir()) == [], "nothing is written when the hash is wrong"


def test_checksums_without_our_asset_are_refused(api: str, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    publish(api, monkeypatch)
    Api.payloads["/sums"] = (200, f"{DIGEST}  something-else\n".encode())
    release = update.latest_release()
    assert release is not None
    with pytest.raises(UpdateError, match="lists no entry"):
        update.download(release, tmp_path)


def test_a_verified_download_lands_on_disk(api: str, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    publish(api, monkeypatch)
    release = update.latest_release()
    assert release is not None
    path = update.download(release, tmp_path)
    assert path.read_bytes() == BINARY
    assert path.name == ASSET


# ---- installing --------------------------------------------------------


def test_installing_from_a_source_checkout_is_refused(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(update, "running_frozen", lambda: False)
    with pytest.raises(UpdateError, match="packaged application"):
        update.install(tmp_path / "new")


def test_install_swaps_the_binary_and_keeps_the_old_one(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(update, "running_frozen", lambda: True)
    current = tmp_path / "pharmparser.exe"
    current.write_bytes(b"old")
    new = tmp_path / "new.exe"
    new.write_bytes(b"new")

    update.install(new, current)

    assert current.read_bytes() == b"new"
    assert (tmp_path / "pharmparser.exe.old").read_bytes() == b"old"


def test_a_failed_swap_puts_the_working_binary_back(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Losing the installed app to a half-finished update is the worst outcome here."""
    monkeypatch.setattr(update, "running_frozen", lambda: True)
    current = tmp_path / "pharmparser.exe"
    current.write_bytes(b"old")

    real_replace = update.os.replace
    calls: list[int] = []

    def fail_second(src: str | Path, dst: str | Path) -> None:
        calls.append(1)
        if len(calls) == 2:
            raise OSError("disk full")
        real_replace(src, dst)

    monkeypatch.setattr(update.os, "replace", fail_second)
    with pytest.raises(OSError, match="disk full"):
        update.install(tmp_path / "new.exe", current)

    assert current.read_bytes() == b"old"


def test_clean_superseded_removes_the_previous_binary(tmp_path: Path) -> None:
    current = tmp_path / "pharmparser.exe"
    current.write_bytes(b"new")
    (tmp_path / "pharmparser.exe.old").write_bytes(b"old")

    update.clean_superseded(current)

    assert not (tmp_path / "pharmparser.exe.old").exists()
    assert current.exists()


def test_clean_superseded_tolerates_a_locked_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    current = tmp_path / "pharmparser.exe"
    current.write_bytes(b"new")
    monkeypatch.setattr(Path, "unlink", lambda *a, **k: (_ for _ in ()).throw(OSError("locked")))
    update.clean_superseded(current)  # must not raise


# -- the asset name is a contract with the release workflow -------------------


@pytest.mark.parametrize(
    ("platform", "expected"),
    [
        ("win32", "pharmparser-1.2.3-windows-x64.exe"),
        ("linux", "pharmparser-1.2.3-linux-x64"),
        ("darwin", "pharmparser-1.2.3-macos-x64"),
    ],
)
def test_asset_name_per_platform(platform: str, expected: str, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(update.sys, "platform", platform)
    assert update.asset_name_for("1.2.3") == expected


def test_the_updater_looks_for_the_name_the_release_workflow_builds(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A silent-death test.

    The workflow names assets and the updater looks them up by exact name. If either
    side is renamed the updater simply never finds a release, on every machine, with
    no error anywhere — so the two are compared here rather than left to agree by
    memory.
    """
    workflow = Path(__file__).resolve().parents[2] / ".github/workflows/release.yml"
    text = workflow.read_text(encoding="utf-8")

    built = [
        line.split('"')[-2]
        for line in text.splitlines()
        if 'cp "dist/pharmparser$SUFFIX"' in line
    ]
    assert built, "release.yml no longer names a GUI asset the way this test expects"
    template = built[0]  # release/pharmparser-$VERSION-$PLATFORM$SUFFIX

    for platform, suffix, expected_platform in (
        ("win32", ".exe", "windows-x64"),
        ("linux", "", "linux-x64"),
    ):
        produced = (
            template.removeprefix("release/")
            .replace("$VERSION", "1.2.3")
            .replace("$PLATFORM", expected_platform)
            .replace("$SUFFIX", suffix)
        )
        monkeypatch.setattr(update.sys, "platform", platform)
        assert update.asset_name_for("1.2.3") == produced, (
            f"release.yml builds {produced!r} but the updater looks for "
            f"{update.asset_name_for('1.2.3')!r}"
        )


def test_the_workflow_publishes_the_checksums_the_updater_requires() -> None:
    workflow = Path(__file__).resolve().parents[2] / ".github/workflows/release.yml"
    assert update.CHECKSUMS_ASSET in workflow.read_text(encoding="utf-8")


# -- restarting ---------------------------------------------------------------


def test_restart_launches_the_new_binary_and_leaves(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    launched: list[tuple[list[str], dict[str, object]]] = []
    monkeypatch.setattr(update.subprocess, "Popen", lambda argv, **kwargs: launched.append((argv, kwargs)))
    monkeypatch.setattr(update.sys, "argv", ["pharmparser", "--verbose"])
    monkeypatch.setenv("PHARMPARSER_RESTART_TEST", "preserved")
    monkeypatch.setenv("_PYI_ARCHIVE_FILE", "old-pharmparser.exe")
    monkeypatch.setenv("_PYI_PARENT_PROCESS_LEVEL", "1")
    monkeypatch.setenv("_PYI_APPLICATION_HOME_DIR", "old-onefile-directory")

    target = tmp_path / "pharmparser.exe"
    with pytest.raises(SystemExit) as exit_info:
        update.restart(target)

    assert exit_info.value.code == 0
    assert len(launched) == 1
    argv, kwargs = launched[0]
    assert argv == [str(target), "--verbose"]
    assert kwargs["close_fds"] is True
    environment = kwargs["env"]
    assert isinstance(environment, dict)
    assert environment["PHARMPARSER_RESTART_TEST"] == "preserved"
    assert environment["PYINSTALLER_RESET_ENVIRONMENT"] == "1"
    assert not any(name.startswith("_PYI_") for name in environment)


# -- release-please drives the version the updater compares -------------------


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def test_the_manifest_agrees_with_the_installed_version() -> None:
    """release-please writes both; if they drift it is bumping from the wrong base."""
    manifest = json.loads((_repo_root() / ".release-please-manifest.json").read_text())
    assert manifest["."] == update.current_version()


def test_release_please_produces_tags_the_updater_can_read() -> None:
    """`include-component-in-tag: false` is what keeps tags as plain vX.Y.Z.

    With it on, tags look like `pharmparser-v0.2.0`, which parse_version rejects —
    the updater would then never offer an update, silently.
    """
    config = json.loads((_repo_root() / "release-please-config.json").read_text())
    package = config["packages"]["."]
    assert package["include-component-in-tag"] is False
    assert package["release-type"] == "python"
    assert update.parse_version(f"v{update.current_version()}") is not None


def test_releases_are_not_published_as_prereleases() -> None:
    """available_update() skips prereleases, so publishing them would ship nothing."""
    config = json.loads((_repo_root() / "release-please-config.json").read_text())
    assert config["packages"]["."]["prerelease"] is False
    assert config["packages"]["."]["draft"] is False
