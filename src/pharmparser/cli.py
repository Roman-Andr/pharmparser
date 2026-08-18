"""Headless entry point: ``pharmparser-cli``.

Runs the whole pipeline — config, scrape, export — without a display, which is
what lets CI exercise it end to end.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .config import ConfigError
from .controller import Controller
from .export import select_exporter
from .scraping import NoPharmaciesError, ScrapeError

logger = logging.getLogger(__name__)


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
    parser.add_argument(
        "--cache",
        action="store_true",
        help="reuse this profile's cached prices when present, and cache the result",
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
        controller = Controller.load(args.config, exporter=select_exporter(macros=args.macros))
        profile = controller.select_profile(args.profile)
        logger.info("Parsing profile %r (%d pharmacies)", profile.name, len(profile.scrapable))

        logger.info("Wrote %s", controller.run(profile, use_cache=args.cache, path=args.output))
    except (ConfigError, NoPharmaciesError, ScrapeError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        return 130

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
