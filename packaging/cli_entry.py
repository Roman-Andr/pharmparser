"""PyInstaller entry point for the headless CLI. See gui_entry.py for why.

``freeze_support`` matters here as well as in the GUI: the scrape starts a process
pool for parsing, and a frozen build without this call re-runs the whole program in
every worker.
"""

from multiprocessing import freeze_support

from pharmparser.cli import main

if __name__ == "__main__":
    freeze_support()
    raise SystemExit(main())
