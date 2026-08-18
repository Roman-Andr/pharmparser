"""End-to-end workbook check that runs without Excel.

Spreadsheet.export drives Excel over COM to inject the macro buttons, so it is
Windows-only. The openpyxl half — which produces all of the actual content — is
not, and this exercises it: build every sheet, save a real .xlsx, read it back.
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

from pharmparser.config import ExportSettings
from pharmparser.domain import PriceTable, absolute_difference, percentage_difference
from pharmparser.excel.formatters import AnalysisFormatter, DataFormatter

SHEETS = ["Данные", "Проценты", "Анализ"]


def build_workbook(settings: ExportSettings, table: PriceTable) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    formatters = [
        (DataFormatter(settings, table, absolute_difference), "Данные"),
        (DataFormatter(settings, table, percentage_difference), "Проценты"),
        (AnalysisFormatter(settings, table), "Анализ"),
    ]
    for formatter, title in formatters:
        formatter.format(wb.create_sheet(title))
    return wb


def test_workbook_saves_and_reloads(tmp_path: Path, settings: ExportSettings, table: PriceTable) -> None:
    target = tmp_path / "data.xlsx"
    build_workbook(settings, table).save(target)

    reloaded = load_workbook(target)
    assert reloaded.sheetnames == SHEETS

    data = reloaded["Данные"]
    assert [cell.value for cell in data[3]] == [
        "Название",
        "Аптека 1",
        "Аптека 2",
        "Разница",
        "Аптека 3",
        "Разница",
    ]
    assert data["A4"].value == "Аспирин, 100мг"
    assert data["D4"].value == 1.5

    # Percentages differ from absolute differences on the same layout.
    assert reloaded["Проценты"]["D4"].value == 30.0

    assert reloaded["Анализ"]["A1"].value == "Аптека 1"
    assert reloaded["Анализ"]["B4"].value == 2  # "Позиций ниже всех" — the B2 fix


def test_conditional_formatting_survives_a_save(tmp_path: Path, settings: ExportSettings, table: PriceTable) -> None:
    target = tmp_path / "data.xlsx"
    build_workbook(settings, table).save(target)

    ranges = {str(rng.sqref) for rng in load_workbook(target)["Данные"].conditional_formatting}
    assert ranges == {"D4:D7", "F4:F7"}


def test_single_pharmacy_workbook_is_still_valid(tmp_path: Path, settings: ExportSettings) -> None:
    """A profile with one pharmacy has no difference columns at all."""
    from pharmparser.domain import Pharmacy

    table = PriceTable.build([(Pharmacy("1", "Аптека 1"), {"Аспирин": 5.0})])
    target = tmp_path / "single.xlsx"
    build_workbook(settings, table).save(target)

    reloaded = load_workbook(target)
    assert [cell.value for cell in reloaded["Данные"][3]] == ["Название", "Аптека 1"]
    assert reloaded["Анализ"]["B3"].value == 0  # mean competitor assortment
