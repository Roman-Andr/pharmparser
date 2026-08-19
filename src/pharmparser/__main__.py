"""Desktop entry point: authenticated local React UI."""

from multiprocessing import freeze_support


def main() -> int:
    from .web.desktop import run_desktop

    return run_desktop()


if __name__ == "__main__":
    freeze_support()
    raise SystemExit(main())
