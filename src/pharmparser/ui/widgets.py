"""Small widget factories.

Lived in ``utils/`` until phase 5, where its re-export meant ``import utils``
pulled in the whole GUI toolkit — so ``core.parser_engine`` transitively depended
on customtkinter (A3).
"""

from __future__ import annotations

from typing import Any

from customtkinter import CTkEntry


def create_custom_entry(parent: Any, placeholder: str, initial_text: str = "") -> CTkEntry:
    """A text entry with select-all and clear-selection bound the usual way."""
    entry = CTkEntry(parent, placeholder_text=placeholder)
    if initial_text:
        entry.insert(0, initial_text)
    entry.bind("<Control-a>", lambda e: ["break", e.widget.select_range(0, "end"), e.widget.icursor("end")][0])
    entry.bind("<Escape>", lambda e: e.widget.select_clear())
    return entry
