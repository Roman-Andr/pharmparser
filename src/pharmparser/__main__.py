"""GUI entry point: ``python -m pharmparser`` (or the ``pharmparser`` console script)."""

import logging
import sys
from multiprocessing import freeze_support


def main() -> int:
    from customtkinter import set_appearance_mode, set_default_color_theme

    from .config import ConfigError
    from .ui import App

    set_appearance_mode("System")
    set_default_color_theme("blue")

    try:
        app = App()
    except ConfigError as error:
        print(f"error: {error}", file=sys.stderr)
        return 1

    app.mainloop()
    return 0


if __name__ == "__main__":
    freeze_support()
    logging.basicConfig(level=logging.INFO)
    raise SystemExit(main())
