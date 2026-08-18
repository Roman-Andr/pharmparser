"""GUI entry point: ``python -m pharmparser`` (or the ``pharmparser`` console script)."""

import logging
from multiprocessing import freeze_support


def main() -> None:
    from customtkinter import set_appearance_mode, set_default_color_theme

    from .ui import App

    set_appearance_mode("System")
    set_default_color_theme("blue")

    app = App()
    app.mainloop()


if __name__ == "__main__":
    freeze_support()
    logging.basicConfig(level=logging.INFO)
    main()
