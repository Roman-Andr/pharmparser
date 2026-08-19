"""VBA macro buttons for the exported workbook.

Deciding, compiling and packaging the macros is cross-platform; only
:mod:`.injector` needs Windows, and it imports its COM bindings lazily."""

from .button import Button, rgb
from .criteria import FilterCriteria, SortOrder
from .injector import excel_application, inject, replace_atomically
from .layout import buttons_for
from .macros import ApplyFiltersMacro, Macro, RemoveFiltersMacro, SortMacro, macro_identifier
from .source import MODULE_NAME, module_source
from .xlsm import ButtonSpec, package

__all__ = [
    "MODULE_NAME",
    "ApplyFiltersMacro",
    "Button",
    "ButtonSpec",
    "FilterCriteria",
    "Macro",
    "RemoveFiltersMacro",
    "SortMacro",
    "SortOrder",
    "buttons_for",
    "excel_application",
    "inject",
    "macro_identifier",
    "module_source",
    "package",
    "replace_atomically",
    "rgb",
]
