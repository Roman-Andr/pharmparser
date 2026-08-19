"""Shapes on the sheet that run a macro when clicked."""

from __future__ import annotations

from typing import Any

from .macros import Macro

MSO_SHAPE_RECTANGLE = 1
XL_CENTER = -4108


def rgb(red: int, green: int, blue: int) -> int:
    """Excel's packed BGR colour value.

    Computed here rather than imported from ``win32api`` so that describing a
    button needs nothing from Windows; only drawing one does.
    """
    return red | (green << 8) | (blue << 16)


class Button:
    """A push button drawn over ``cell_address`` and wired to ``macro``."""

    def __init__(
        self,
        cell_address: str,
        caption: str,
        macro: Macro,
        back_color: int | None = None,
        fore_color: int | None = None,
    ) -> None:
        self.cell_address = cell_address
        self.caption = caption
        self.macro = macro
        self.back_color = back_color
        self.fore_color = fore_color
        self.button_name: str | None = None

    def create(self, worksheet: Any) -> None:
        """Draw the shape, then register its geometry with the macro it triggers.

        Registration has to happen here because the name and size are Excel's to
        assign. Macro source is rendered afterwards, from
        :attr:`~pharmparser.export.vba.macros.Macro.code` (B1).
        """
        cell = worksheet.Range(self.cell_address)
        button = worksheet.Shapes.AddShape(
            MSO_SHAPE_RECTANGLE, cell.Left, cell.Top, cell.Width, cell.Height
        )
        button.TextFrame.Characters().Text = self.caption
        button.OnAction = self.macro.name
        self.button_name = button.Name

        if self.back_color is not None:
            button.Fill.BackColor.RGB = self.back_color
            button.Fill.ForeColor.RGB = self.fore_color

        button.TextFrame.HorizontalAlignment = XL_CENTER
        button.TextFrame.VerticalAlignment = XL_CENTER

        self.macro.add_position_code(self.save_position_code(), self.restore_position_code())

    @property
    def _variable(self) -> str:
        if self.button_name is None:
            raise RuntimeError("the button has not been drawn yet, so Excel has not named it")
        return f"btn{self.button_name.replace(' ', '')}"

    def save_position_code(self) -> str:
        """VBA that records where Excel currently has this button."""
        name = self._variable
        return f"""
        Dim {name} As Shape
        Set {name} = ActiveSheet.Shapes("{self.button_name}")
        Dim {name}Left As Double
        Dim {name}Top As Double
        Dim {name}Width As Double
        Dim {name}Height As Double
        {name}Left = {name}.Left
        {name}Top = {name}.Top
        {name}Width = {name}.Width
        {name}Height = {name}.Height
        """

    def restore_position_code(self) -> str:
        """VBA that puts the button back where :meth:`save_position_code` found it."""
        name = self._variable
        return f"""
        {name}.Left = {name}Left
        {name}.Top = {name}Top
        {name}.Width = {name}Width
        {name}.Height = {name}Height
        """
