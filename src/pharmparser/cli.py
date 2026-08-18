"""Headless entry point: ``pharmparser-cli``.

Runs the whole pipeline — config, scrape, export — without a display, which is
what lets CI exercise it end to end.
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import sys
from pathlib import Path

from .config import AppConfig, ConfigError, Profile, load_config
from .export import select_exporter
from .scraping import NoPharmaciesError, ScrapeError, scrape_profile

logger = logging.getLogger(__name__)


def select_profile(config: AppConfig, name: str | None) -> Profile:
    if not config.profiles:
        raise ConfigError("no profiles are configured")
    if name is None:
        return config.profiles[0]
    for profile in config.profiles:
        if profile.name == name:
            return profile
    available = ", ".join(profile.name for profile in config.profiles)
    raise ConfigError(f"no profile named {name!r}; available profiles: {available}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pharmparser", description=__doc__)
    parser.add_argument("--config", type=Path, default=None, help="path to config.json")
    parser.add_argument("--profile", default=None, help="profile name (default: the first one)")
    parser.add_argument("--output", type=Path, default=None, help="output .xlsx path")
    parser.add_argument(
        "--macros",
        action="store_true",
        help="also inject the VBA sort/filter buttons and produce an .xlsm (Windows only)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="log every request")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )

    try:
        config = load_config(args.config)
        profile = select_profile(config, args.profile)
        logger.info("Parsing profile %r (%d pharmacies)", profile.name, len(profile.scrapable))

        table = asyncio.run(scrape_profile(config.request, profile.pharmacies))

        exporter = select_exporter(macros=args.macros)
        logger.info("Wrote %s", exporter.export(config.settings, table, args.output))
    except (ConfigError, NoPharmaciesError, ScrapeError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        return 130

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
