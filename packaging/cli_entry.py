"""PyInstaller entry point for the headless CLI. See gui_entry.py for why."""

from pharmparser.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
