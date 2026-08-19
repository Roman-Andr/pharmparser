"""GUI entry point: ``python -m pharmparser`` (or the ``pharmparser`` console script)."""

import sys
from multiprocessing import freeze_support


def main() -> int:
    from customtkinter import set_appearance_mode, set_default_color_theme

    from .config import ConfigError
    from .logging_ import configure as configure_logging
    from .ui import App

    # The packaged Windows binary has no console, so the log file is the only
    # record of a failed run.
    log_file = configure_logging(console=sys.stderr is not None)

    set_appearance_mode("System")
    set_default_color_theme("blue")

    try:
        app = App()
    except ConfigError as error:
        print(f"error: {error}", file=sys.stderr)
        if log_file is not None:
            print(f"see {log_file} for details", file=sys.stderr)
        return 1

    app.mainloop()
    return 0


if __name__ == "__main__":
    freeze_support()
    raise SystemExit(main())
