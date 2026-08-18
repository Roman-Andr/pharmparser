"""Baseline import and enum checks."""

from pharmparser.config import ExportSettings
from pharmparser.export.vba import FilterCriteria, SortOrder


def test_package_imports_without_windows() -> None:
    """The VBA package must import on Linux; only COM *use* requires Windows."""
    import pharmparser.export
    import pharmparser.export.vba.injector

    assert pharmparser.export.MacroExporter is not None
    assert pharmparser.export.vba.injector.inject is not None


def test_cli_imports_without_a_display() -> None:
    from pharmparser.cli import build_parser

    assert build_parser().prog == "pharmparser"


def test_settings_defaults_round_trip_to_camel_case() -> None:
    dumped = ExportSettings().model_dump(by_alias=True)
    assert dumped["fileName"] == "data.xlsx"
    assert dumped["colWidth"] == 50
    assert ExportSettings.model_validate(dumped) == ExportSettings()


def test_enums_have_stable_values() -> None:
    assert FilterCriteria.GREATER_THAN_ZERO.value == ">0"
    assert SortOrder.ASCENDING.value == "xlAscending"
