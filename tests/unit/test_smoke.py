"""Baseline tests locking in current behaviour before the phase 1-5 rewrite.

These are intentionally thin: their job is to prove the package imports and the
test harness runs on a non-Windows machine, which was impossible before phase 0.
"""

from dataclasses import asdict

from pharmparser.utils import DataType, FilterCriteria, Settings, SortOrder


def test_package_imports_without_windows() -> None:
    """The Excel package must import on Linux; only COM *use* requires Windows."""
    import pharmparser.excel
    import pharmparser.excel.spreadsheet  # noqa: F401


def test_settings_roundtrip(settings: Settings) -> None:
    assert asdict(settings)["fileName"] == "data.xlsx"


def test_enums_have_stable_values() -> None:
    assert FilterCriteria.GREATER_THAN_ZERO.value == ">0"
    assert SortOrder.ASCENDING.value == "xlAscending"


def test_datatype_alias_is_usable(price_table: DataType) -> None:
    assert price_table["Аптека 1"]["Аспирин, 100мг"] == 5.00
