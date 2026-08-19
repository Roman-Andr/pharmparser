"""The desktop front end.

Widget classes are resolved lazily, so importing this package does not pull in
customtkinter — only asking for a widget does. Until phase 5 the reverse held:
``utils/__init__.py`` re-exported a widget factory, so ``import utils`` dragged the
GUI toolkit into the scraper (A3).
"""

from typing import Any

__all__ = ["App", "Entry", "Profile", "ProfileSelector", "create_custom_entry"]

_WIDGETS = {
    "App": ".app",
    "Entry": ".entry",
    "Profile": ".profile",
    "ProfileSelector": ".profile_selector",
    "create_custom_entry": ".widgets",
}


def __getattr__(name: str) -> Any:
    if name in _WIDGETS:
        from importlib import import_module

        return getattr(import_module(_WIDGETS[name], __name__), name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
