"""Characterisation tests for the generated VBA.

These lock in the exact macro source, so the phase 3 rewrite is verifiable as
behaviour-preserving, and they cover bug B1: macro source used to be interpolated
in ``__init__``, but ``position_codes`` is only filled later by ``Button.create()``,
so the button save/restore geometry never reached the workbook.
"""

import pytest

from pharmparser.export.vba import (
    ApplyFiltersMacro,
    FilterCriteria,
    RemoveFiltersMacro,
    SortMacro,
    SortOrder,
)
from pharmparser.export.vba.macros import macro_identifier


def test_sort_macro_emits_single_line_sort_statement() -> None:
    macro = SortMacro("D", 6, SortOrder.ASCENDING, "Data")
    assert macro.name == "SortASCENDINGD_data"
    assert 'ActiveSheet.Range("A3:F100000").Sort Key1:=ActiveSheet.Columns("D"), ' \
           "Order1:=xlAscending, Header:=xlYes" in macro.code
    assert "Sub SortASCENDINGD_data()" in macro.code


def test_remove_filters_macro_clears_every_diff_column() -> None:
    macro = RemoveFiltersMacro(6, "Data")
    assert "For col = 4 To 6 Step 2" in macro.code
    assert 'ActiveSheet.Range("A3:F100000").AutoFilter Field:=col' in macro.code


def test_apply_filters_macro_carries_the_criteria() -> None:
    macro = ApplyFiltersMacro(6, FilterCriteria.GREATER_THAN_ZERO, "Data")
    assert 'Criteria1:=">0"' in macro.code


def test_position_code_reaches_the_generated_macro() -> None:
    """Regression for B1.

    ``code`` is rendered on demand, so geometry registered after construction —
    which is the only time it can be known — is part of the emitted ``Sub``.
    """
    macro = SortMacro("D", 6, SortOrder.ASCENDING, "Data")
    macro.add_position_code("Dim btnFoo As Shape", "btnFoo.Left = btnFooLeft")
    assert "Dim btnFoo As Shape" in macro.code
    assert "btnFoo.Left = btnFooLeft" in macro.code


def test_position_code_brackets_the_macro_body() -> None:
    """The save runs before the sort and the restore after it, or the geometry is lost."""
    macro = SortMacro("D", 6, SortOrder.ASCENDING, "Data")
    macro.add_position_code("SAVE_MARKER", "RESTORE_MARKER")
    code = macro.code
    assert code.index("SAVE_MARKER") < code.index("Key1:=") < code.index("RESTORE_MARKER")


def test_every_macro_renders_registered_geometry() -> None:
    macros = [
        SortMacro("D", 6, SortOrder.DESCENDING, "Data"),
        RemoveFiltersMacro(6, "Data"),
        ApplyFiltersMacro(6, FilterCriteria.LESS_THAN_ZERO, "Data"),
    ]
    for macro in macros:
        macro.add_position_code("SAVE_MARKER", "RESTORE_MARKER")
        assert "SAVE_MARKER" in macro.code, macro.name
        assert "RESTORE_MARKER" in macro.code, macro.name


def test_a_macro_without_buttons_still_renders() -> None:
    assert "Sub RemoveFilters_data()" in RemoveFiltersMacro(6, "Data").code


def test_macro_names_are_ascii_even_for_russian_sheets() -> None:
    """Module streams live in the project code page, so a Cyrillic Sub name is mojibake."""
    macro = ApplyFiltersMacro(6, FilterCriteria.GREATER_THAN_ZERO, "Данные")
    assert macro.name == "ApplyFilters_dannye"
    assert macro.code.isascii()


@pytest.mark.parametrize(
    ("sheet", "expected"),
    [("Данные", "dannye"), ("Проценты", "protsenty"), ("Анализ", "analiz"),
     ("Sheet 1", "sheet_1"), ("Ц/Ч", "ts_ch"), ("", "Sheet")],
)
def test_macro_identifier_transliterates(sheet: str, expected: str) -> None:
    assert macro_identifier(sheet) == expected
