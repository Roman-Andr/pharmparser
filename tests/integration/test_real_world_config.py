"""Compatibility tests against a real user's config.json.

``tests/fixtures/real_world_config.json`` is a real configuration with the
credentials redacted — six profiles, Cyrillic and Latin pharmacy names, names
carrying trailing and doubled spaces, the same pharmacy id reused across
profiles, a full browser header set (including a lowercase ``host`` and an
explicit ``Content-Type`` alongside form-encoded data) and non-default column
widths.

These guard the promise that existing config files keep working unchanged.
"""

import json
from pathlib import Path

import aiohttp
import pytest
from aioresponses import aioresponses
from openpyxl import load_workbook

from pharmparser.cli import main
from pharmparser.config import load_config, save_config

FIXTURE = Path(__file__).parent.parent / "fixtures" / "real_world_config.json"
URL = "https://tabletka.by/ajax-request/reload-pharmacy-price"

ROW = """
<div class="result-row">
  <div class="tooltip-info-header"><a href="/d/{n}">{name}</a></div>
  <div><span class="form-title">{form}</span><span class="price-value">от {price} р.</span></div>
</div>
"""


def page(price: str) -> str:
    return ROW.format(n=1, name="Аспирин", form="100мг", price=price) + ROW.format(
        n=2, name="Цитрамон", form="10шт", price=price
    )


@pytest.fixture
def config_file(tmp_path: Path) -> Path:
    path = tmp_path / "config.json"
    path.write_text(FIXTURE.read_text(encoding="utf-8"), encoding="utf-8")
    return path


def stub(mocked: aioresponses, pharmacies: int) -> None:
    for i in range(pharmacies):
        mocked.post(URL, payload={"priceCount": 2, "data": ""})
        mocked.post(URL, payload={"priceCount": 2, "data": page(f"{5 + i},00")})


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
        sent = mocked.requests[("POST", aiohttp.helpers.URL(URL))][0]

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
