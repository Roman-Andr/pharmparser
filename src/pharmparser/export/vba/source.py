"""Assembling the VBA module the workbook ships with.

The cross-platform packer writes legacy VML form buttons whose anchors are already
outside the sortable range. Unlike shapes created through Excel itself, those controls
can reject a runtime write to ``Shape.Placement`` with COM error 80020009, so their
placement is left entirely to the workbook markup.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from .button import Button

MODULE_NAME = "PharmParser"


def module_source(sheets: Mapping[str, Sequence[Button]]) -> str:
    """Render one module holding every macro the workbook's buttons reference."""
    parts = ["Option Explicit"]
    seen: set[str] = set()

    for buttons in sheets.values():
        for button in buttons:
            macro = button.macro
            if macro.name in seen:
                continue
            seen.add(macro.name)
            parts.append(_dedent(macro.code))

    return "\n\n".join(parts) + "\n"


def _dedent(code: str) -> str:
    """Strip the f-string indentation the macro templates carry."""
    lines = [line.rstrip() for line in code.strip("\n").splitlines()]
    indents = [len(line) - len(line.lstrip()) for line in lines if line.strip()]
    margin = min(indents, default=0)
    return "\n".join(line[margin:] if line.strip() else "" for line in lines).strip()
