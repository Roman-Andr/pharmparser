"""The VBA the sort/filter buttons run.

Each macro brackets its statements with code that saves every registered button's
geometry and puts it back afterwards, because sorting and filtering move shapes
anchored to the rows involved.

That bracketing is where bug B1 lived: the source was interpolated in
``__init__``, but the position code is only registered later, when
:meth:`~pharmparser.export.vba.button.Button.create` has an actual shape to
measure. Every workbook therefore shipped with the save/restore lines empty.
:attr:`Macro.code` is now rendered on demand, so registration order no longer
matters.
"""

from __future__ import annotations

from abc import ABC, abstractmethod

from openpyxl.utils import get_column_letter

from .criteria import FilterCriteria, SortOrder

FIRST_DIFFERENCE_COLUMN = 4
"""First "Разница" column; the sortable/filterable ones run from here in steps of 2."""

FIRST_DATA_ROW = 3
LAST_DATA_ROW = 100000

_TRANSLITERATION = str.maketrans({
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "ts",
    "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu",
    "я": "ya",
})


def macro_identifier(name: str) -> str:
    """An ASCII, VBA-legal fragment of a sheet name, for use in a ``Sub`` name.

    Module streams are stored in the project code page, so a Cyrillic ``Sub`` name
    comes back as mojibake — and the sheet titles here are Russian. Transliterating
    keeps the generated names readable without leaving ASCII.
    """
    lowered = name.casefold().translate(_TRANSLITERATION)
    cleaned = "".join(char if char.isascii() and char.isalnum() else "_" for char in lowered)
    cleaned = "_".join(part for part in cleaned.split("_") if part)
    return cleaned or "Sheet"


class Macro(ABC):
    """One ``Sub`` in the generated module."""

    def __init__(self, name: str, end_column: int) -> None:
        self.name = name
        self.end_column = end_column
        self.data_range = f"A{FIRST_DATA_ROW}:{get_column_letter(end_column)}{LAST_DATA_ROW}"
        self.position_codes: list[tuple[str, str]] = []
        self.prologue = ""
        """A statement run first, before the saved geometry is read.

        The cross-platform packer uses it to pin every shape in place, which the
        COM path achieves with :attr:`position_codes` instead."""

    def add_position_code(self, save: str, restore: str) -> None:
        """Register a button's geometry to be saved and restored around this macro."""
        self.position_codes.append((save, restore))

    @property
    @abstractmethod
    def body(self) -> str:
        """The statements this macro exists to run."""

    @property
    def code(self) -> str:
        """The full ``Sub``, rendered from the position codes registered *so far*.

        Rendering here rather than in ``__init__`` is the B1 fix: buttons only
        register their geometry once they have been drawn on the sheet.
        """
        saves = "\n".join(save for save, _ in self.position_codes)
        restores = "\n".join(restore for _, restore in self.position_codes)
        return f"""
        Sub {self.name}()
            Application.ScreenUpdating = False
            {self.prologue}
            {saves}
            {self.body}
            {restores}
            Application.ScreenUpdating = True
        End Sub
        """


class SortMacro(Macro):
    """Sort the data block by one column."""

    def __init__(self, column: str, end_column: int, sort_order: SortOrder, sheet_name: str) -> None:
        self.column = column
        self.sort_order = sort_order
        super().__init__(f"Sort{sort_order.name}{column}_{macro_identifier(sheet_name)}", end_column)

    @property
    def body(self) -> str:
        return (
            f'ActiveSheet.Range("{self.data_range}").Sort '
            f'Key1:=ActiveSheet.Columns("{self.column}"), '
            f"Order1:={self.sort_order.value}, Header:=xlYes"
        )


class RemoveFiltersMacro(Macro):
    """Clear every difference-column filter and restore the original item order."""

    def __init__(self, end_column: int, sheet_name: str) -> None:
        super().__init__(f"RemoveFilters_{macro_identifier(sheet_name)}", end_column)

    @property
    def body(self) -> str:
        return f"""If ActiveSheet.AutoFilterMode Then
                Dim col As Integer
                For col = {FIRST_DIFFERENCE_COLUMN} To {self.end_column} Step 2
                    ActiveSheet.Range("{self.data_range}").AutoFilter Field:=col
                Next col
            End If
            {self._reset_sort}"""

    @property
    def _reset_sort(self) -> str:
        return (
            f'ActiveSheet.Range("{self.data_range}").Sort '
            f'Key1:=ActiveSheet.Columns("A"), Order1:=xlAscending, Header:=xlYes'
        )


class ApplyFiltersMacro(Macro):
    """Filter every difference column by the same criterion."""

    def __init__(self, end_column: int, criteria: FilterCriteria, sheet_name: str) -> None:
        self.criteria = criteria
        super().__init__(f"ApplyFilters_{macro_identifier(sheet_name)}", end_column)

    @property
    def body(self) -> str:
        return f"""ActiveSheet.AutoFilterMode = False
            Dim col As Integer
            For col = {FIRST_DIFFERENCE_COLUMN} To {self.end_column} Step 2
                ActiveSheet.Range("{self.data_range}").AutoFilter Field:=col, Criteria1:="{self.criteria.value}"
            Next col"""
