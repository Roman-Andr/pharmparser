"""Compatibility tests against a real user's config.json.

``tests/fixtures/real_world_config.json`` is a real configuration with the
credentials redacted — six profiles, Cyrillic and Latin pharmacy names, names
carrying trailing and doubled spaces, the same pharmacy id reused across
profiles, a full browser header set (including a lowercase ``host``, an explicit
``Content-Type`` alongside form-encoded data and the browser's whole sixteen-key
cookie) and non-default column widths.

These guard the promise that existing config files keep working unchanged.
"""

import asyncio
import json
from pathlib import Path

import aiohttp
import pytest
from aioresponses import aioresponses
from openpyxl import load_workbook

from pharmparser.cli import main
from pharmparser.config import load_config, save_config
from pharmparser.export import export_with_macros
from pharmparser.scraping import scrape_profile

from ..pages import page as build_page
from ..pages import row

FIXTURE = Path(__file__).parent.parent / "fixtures" / "real_world_config.json"
URL = "https://tabletka.by/ajax-request/reload-pharmacy-price"
ENDPOINT = URL + "/"
"""What the client actually posts to: tabletka.by 500s without the trailing slash (B17)."""


def page(price: str) -> str:
    return build_page(
        row("Аспирин", "100мг", f"от {price} р.", ls=1),
        row("Цитрамон", "10шт", f"от {price} р.", ls=2),
    )


@pytest.fixture
def config_file(tmp_path: Path) -> Path:
    path = tmp_path / "config.json"
    path.write_text(FIXTURE.read_text(encoding="utf-8"), encoding="utf-8")
    return path


def stub(mocked: aioresponses, pharmacies: int) -> None:
    for i in range(pharmacies):
        mocked.post(ENDPOINT, payload={"priceCount": 2, "data": ""})
        mocked.post(ENDPOINT, payload={"priceCount": 2, "data": page(f"{5 + i},00")})


# -- loading -------------------------------------------------------------------


def test_loads_without_error(config_file: Path) -> None:
    config = load_config(config_file)
    assert [p.name for p in config.profiles] == [f"Profile {i}" for i in range(1, 7)]
    assert [len(p.pharmacies) for p in config.profiles] == [9, 5, 3, 3, 3, 8]


def test_custom_widths_are_read(config_file: Path) -> None:
    settings = load_config(config_file).settings
    assert (settings.col_width, settings.cell_width, settings.diff_width) == (45, 13, 9)


def test_awkward_pharmacy_names_are_preserved(config_file: Path) -> None:
    """Trailing spaces, doubled spaces, "№1" and "2-2" are all legal names."""
    profiles = {p.name: p for p in load_config(config_file).profiles}
    assert "Искамед " in [e.name for e in profiles["Profile 2"].pharmacies]
    assert "ADEL  Сов.40А" in [e.name for e in profiles["Profile 2"].pharmacies]
    assert "№1" in [e.name for e in profiles["Profile 6"].pharmacies]
    assert "2-2" in [e.name for e in profiles["Profile 5"].pharmacies]


def test_a_pharmacy_id_may_repeat_across_profiles(config_file: Path) -> None:
    """Pharmacy 381 appears in Profile 2 and Profile 6 under different names."""
    profiles = {p.name: p for p in load_config(config_file).profiles}
    assert "381" in [e.pharmacy_id for e in profiles["Profile 2"].pharmacies]
    assert "381" in [e.pharmacy_id for e in profiles["Profile 6"].pharmacies]


def test_round_trip_is_byte_for_byte_identical(config_file: Path, tmp_path: Path) -> None:
    """Saving must not reorder, rename or drop anything (A6)."""
    out = tmp_path / "saved.json"
    save_config(load_config(config_file), out)
    assert json.loads(out.read_text(encoding="utf-8")) == json.loads(
        config_file.read_text(encoding="utf-8")
    )


# -- scraping ------------------------------------------------------------------


def test_browser_headers_survive_to_the_request(config_file: Path, tmp_path: Path) -> None:
    """An explicit Content-Type alongside form-encoded data must not be rejected."""
    with aioresponses() as mocked:
        stub(mocked, 9)
        assert main(["--config", str(config_file), "--output", str(tmp_path / "r.xlsx")]) == 0
        sent = mocked.requests[("POST", aiohttp.helpers.URL(ENDPOINT))][0]

    headers = sent.kwargs["headers"]
    assert headers["Content-Type"] == "application/x-www-form-urlencoded; charset=UTF-8"
    assert headers["host"] == "tabletka.by"
    assert "lim-result=10" in headers["Cookie"]
    assert sent.kwargs["data"]["id"] == "3563"


# -- exporting -----------------------------------------------------------------


def test_nine_pharmacy_profile_exports(config_file: Path, tmp_path: Path) -> None:
    output = tmp_path / "report.xlsx"
    with aioresponses() as mocked:
        stub(mocked, 9)
        assert main(["--config", str(config_file), "--output", str(output)]) == 0

    sheet = load_workbook(output)["Данные"]
    assert sheet.max_column == 2 + 2 * 8
    assert [cell.value for cell in sheet[3]][:4] == [
        "Название",
        "Аптека 195/3",
        "Добрые леки",
        "Разница",
    ]
    assert sheet.column_dimensions["A"].width == 45
    assert sheet.column_dimensions["C"].width == 13
    assert sheet.column_dimensions["D"].width == 9


def test_every_profile_can_be_selected(config_file: Path, tmp_path: Path) -> None:
    counts = {"Profile 1": 9, "Profile 2": 5, "Profile 3": 3, "Profile 4": 3, "Profile 5": 3, "Profile 6": 8}
    for name, pharmacies in counts.items():
        output = tmp_path / f"{name}.xlsx"
        with aioresponses() as mocked:
            stub(mocked, pharmacies)
            assert main(["--config", str(config_file), "--profile", name, "--output", str(output)]) == 0
        assert load_workbook(output)["Данные"].max_column == 2 + 2 * (pharmacies - 1)


def test_the_whole_browser_cookie_survives_to_the_request(config_file: Path, tmp_path: Path) -> None:
    """Only ``lim-result`` is rewritten; the other fifteen cookies travel untouched."""
    with aioresponses() as mocked:
        stub(mocked, 9)
        assert main(["--config", str(config_file), "--output", str(tmp_path / "r.xlsx")]) == 0
        sent = mocked.requests[("POST", aiohttp.helpers.URL(ENDPOINT))][0]

    configured = load_config(config_file).request.cookie
    keys = [cookie.split("=", 1)[0].strip() for cookie in sent.kwargs["headers"]["Cookie"].split(";")]
    assert keys == [cookie.split("=", 1)[0].strip() for cookie in configured.split(";")]
    assert "region=%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C" in sent.kwargs["headers"]["Cookie"]


def test_the_configured_title_names_the_analysis_sheet(config_file: Path, tmp_path: Path) -> None:
    """B15, against the real value: this config sets ``title`` to "Анализ"."""
    output = tmp_path / "report.xlsx"
    with aioresponses() as mocked:
        stub(mocked, 9)
        assert main(["--config", str(config_file), "--profile", "Profile 1", "--output", str(output)]) == 0

    workbook = load_workbook(output)
    assert load_config(config_file).settings.title == "Анализ"
    assert workbook.sheetnames == ["Данные", "Проценты", "Анализ"]
    assert workbook.properties.title == "Анализ"


def test_the_macro_export_runs_over_the_real_profile(
    config_file: Path, tmp_path: Path, excel_sessions: list
) -> None:
    """The nine-pharmacy profile through the .xlsm path: one Excel, one output file."""
    config = load_config(config_file)
    with aioresponses() as mocked:
        stub(mocked, 9)
        table = asyncio.run(scrape_profile(config.request, config.profiles[0].pharmacies))

    target = export_with_macros(config.settings, table, tmp_path / "data.xlsm")

    assert len(excel_sessions) == 1
    assert [path.name for path in tmp_path.iterdir()] == ["config.json", "data.xlsm"]
    workbook = excel_sessions[0].opened[0]
    # 8 competitors -> 8 difference columns -> 16 sort buttons plus the 2 filter ones.
    assert len(workbook.Sheets("Данные").Shapes.shapes) == 18
    assert target.exists()
