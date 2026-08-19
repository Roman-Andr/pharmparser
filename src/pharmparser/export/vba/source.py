"""Assembling the VBA module the workbook ships with.

The COM path lets Excel name each shape and then brackets every macro with code that
saves and restores those shapes' geometry. The cross-platform packer never gets shape
names from Excel, so it pins the shapes instead — one helper, called by every macro,
which is both shorter and what the geometry dance was trying to achieve.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from .button import Button

MODULE_NAME = "PharmParser"

PIN_SHAPES = "PharmParserPinShapes"

PIN_SHAPES_HELPER = f"""Private Sub {PIN_SHAPES}()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Placement = xlFreeFloating
    Next shp
End Sub"""


def module_source(sheets: Mapping[str, Sequence[Button]]) -> str:
    """Render one module holding every macro the workbook's buttons reference."""
    parts = ["Option Explicit", PIN_SHAPES_HELPER]
    seen: set[str] = set()

    for buttons in sheets.values():
        for button in buttons:
            macro = button.macro
            if macro.name in seen:
                continue
            seen.add(macro.name)
            macro.prologue = PIN_SHAPES
            parts.append(_dedent(macro.code))

    return "\n\n".join(parts) + "\n"


def _dedent(code: str) -> str:
    """Strip the f-string indentation the macro templates carry."""
    lines = [line.rstrip() for line in code.strip("\n").splitlines()]
    indents = [len(line) - len(line.lstrip()) for line in lines if line.strip()]
    margin = min(indents, default=0)
    return "\n".join(line[margin:] if line.strip() else "" for line in lines).strip()
