"""The Windows-only .xlsm export path, driven against a fake Excel.

COM is the only thing here that needs Windows, so replacing the Excel session with
a recorder lets the whole flow — build, inject, replace — run in CI. Covers B1
(button geometry reaching the emitted VBA) and B12 (one Excel process for the
whole workbook, no temp files left behind, atomic replace).
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from pharmparser.config import ExportSettings
from pharmparser.domain import PriceTable
from pharmparser.export import MacroExporter, export_with_macros

from ..fakes import FakeExcel


def test_export_starts_exactly_one_excel_process(
    tmp_path: Path, settings: ExportSettings, table: PriceTable, excel_sessions: list[FakeExcel]
) -> None:
    """Regression for B12: the old code started a fresh Excel per data sheet."""
    export_with_macros(settings, table, tmp_path / "data.xlsm")

    assert len(excel_sessions) == 1
    excel = excel_sessions[0]
    assert excel.quit is True
    assert len(excel.opened) == 1, "the whole workbook is opened once"


def test_export_lands_on_the_target_and_leaves_no_temp_files(
    tmp_path: Path, settings: ExportSettings, table: PriceTable, excel_sessions: list[FakeExcel]
) -> None:
    """Regression for B12: 0data.xlsm/1data.xlsm used to litter the working directory."""
    target = tmp_path / "data.xlsm"
    assert export_with_macros(settings, table, target) == target
    assert target.exists()
    assert [path.name for path in tmp_path.iterdir()] == ["data.xlsm"]


def test_export_replaces_a_previous_report_in_place(
    tmp_path: Path, settings: ExportSettings, table: PriceTable, excel_sessions: list[FakeExcel]
) -> None:
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"stale")

    export_with_macros(settings, table, target)

    assert zipfile.is_zipfile(target), "the stale file was replaced by a real workbook"
    assert [path.name for path in tmp_path.iterdir()] == ["data.xlsm"]


def test_buttons_are_drawn_only_on_the_data_sheets(
    tmp_path: Path, settings: ExportSettings, table: PriceTable, excel_sessions: list[FakeExcel]
) -> None:
    export_with_macros(settings, table, tmp_path / "data.xlsm")

    workbook = excel_sessions[0].opened[0]
    # Apply + Remove, then an up/down pair per difference column (D and F here).
    for sheet_name in ("Данные", "Проценты"):
        sheet = workbook.Sheets(sheet_name)
        assert [shape.OnAction for shape in sheet.Shapes.shapes] == [
            f"ApplyFilters_{sheet_name}",
            f"RemoveFilters_{sheet_name}",
            f"SortDESCENDINGD_{sheet_name}",
            f"SortASCENDINGD_{sheet_name}",
            f"SortDESCENDINGF_{sheet_name}",
            f"SortASCENDINGF_{sheet_name}",
        ]
    assert workbook.Sheets(settings.title).Shapes.shapes == []


def test_emitted_vba_carries_the_button_geometry(
    tmp_path: Path, settings: ExportSettings, table: PriceTable, excel_sessions: list[FakeExcel]
) -> None:
    """Regression for B1, end to end.

    Every button is drawn before any macro source is read, so the save/restore
    lines are no longer interpolated from an empty list.
    """
    export_with_macros(settings, table, tmp_path / "data.xlsm")

    workbook = excel_sessions[0].opened[0]
    modules = workbook.VBProject.VBComponents.components
    assert len(modules) == 2, "one module per data sheet"

    source = "\n".join(code for module in modules for code in module.CodeModule.sources)
    assert "Set btnShape1 = ActiveSheet.Shapes(\"Shape 1\")" in source
    assert "btnShape1.Left = btnShape1Left" in source
    # Each of the six buttons on each of the two data sheets saves its own geometry.
    assert source.count("As Shape") == 12


def test_macro_exporter_defaults_to_the_configured_macro_file(
    tmp_path: Path, settings: ExportSettings, table: PriceTable, excel_sessions: list[FakeExcel],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.chdir(tmp_path)
    assert MacroExporter().export(settings, table) == (tmp_path / "data.xlsm").absolute()
