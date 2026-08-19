"""Generate the checked-in local API contract without touching user data."""

from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

from pharmparser.application import CredentialService, HistoryRepository, SettingsService
from pharmparser.web import create_app, create_services

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "frontend/src/generated/openapi.json"


def generate() -> str:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        services = create_services(
            settings=SettingsService(root / "settings.json"),
            credentials=CredentialService(fallback_path=root / "credentials.json"),
            history=HistoryRepository(root / "history.sqlite3"),
            token="contract-generation-token",
        )
        schema = create_app(services, production=False).openapi()
    return json.dumps(schema, ensure_ascii=False, indent=2, sort_keys=True) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    payload = generate()
    if args.check:
        if not TARGET.exists() or TARGET.read_text(encoding="utf-8") != payload:
            print("OpenAPI contract is stale; run: cd frontend && bun run generate:api")
            return 1
        return 0
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    TARGET.write_text(payload, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
