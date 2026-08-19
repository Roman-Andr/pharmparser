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


def test_the_controller_and_web_adapter_import_without_a_gui_toolkit() -> None:
    """Use cases and the HTTP adapter remain importable in a headless process."""
    import subprocess
    import sys

    script = (
        "import sys; from pharmparser.controller import Controller; "
        "from pharmparser.cli import main; from pharmparser.web import create_app; "
        "assert Controller is not None and main is not None and create_app is not None; "
        "print('webview' in sys.modules, 'tkinter' in sys.modules)"
    )
    result = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, check=True)
    assert result.stdout.strip() == "False False"
