"""Characterisation tests for the generated VBA.

These lock in the exact macro source so the phase 3 rewrite can be verified as
behaviour-preserving. They also pin bug B1 from docs/REFACTOR_PLAN.md: macro
source is interpolated in ``__init__``, but ``position_codes`` is only filled
later by ``Button.create()``, so the button save/restore geometry never reaches
the workbook.
"""

import pytest

from pharmparser.excel.macros import ApplyFiltersMacro, RemoveFiltersMacro, SortMacro
from pharmparser.utils import FilterCriteria, SortOrder


def test_sort_macro_emits_single_line_sort_statement() -> None:
    macro = SortMacro("D", 6, SortOrder.ASCENDING, "Data")
    assert macro.name == "SortASCENDINGD_Data"
    assert 'ActiveSheet.Range("A3:F100000").Sort Key1:=ActiveSheet.Columns("D"), ' \
           "Order1:=xlAscending, Header:=xlYes" in macro.code
    assert "Sub SortASCENDINGD_Data()" in macro.code


def test_remove_filters_macro_clears_every_diff_column() -> None:
    macro = RemoveFiltersMacro(6, "Data")
    assert "For col = 4 To 6 Step 2" in macro.code
    assert 'ActiveSheet.Range("A3:F100000").AutoFilter Field:=col' in macro.code


def test_apply_filters_macro_carries_the_criteria() -> None:
    macro = ApplyFiltersMacro(6, FilterCriteria.GREATER_THAN_ZERO, "Data")
    assert 'Criteria1:=">0"' in macro.code


@pytest.mark.xfail(
    reason="B1: code is interpolated in __init__, before add_position_code runs; "
    "fixed in phase 3",
    strict=True,
)
def test_position_code_reaches_the_generated_macro() -> None:
    macro = SortMacro("D", 6, SortOrder.ASCENDING, "Data")
    macro.add_position_code("Dim btnFoo As Shape", "btnFoo.Left = btnFooLeft")
    assert "Dim btnFoo As Shape" in macro.code
