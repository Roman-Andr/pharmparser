"""End-to-end run of the CLI against a faked tabletka.by."""

import json
from pathlib import Path

import pytest
from aioresponses import aioresponses
from openpyxl import load_workbook

from pharmparser.cli import main
from pharmparser.platform_ import supports_excel_macros

from ..pages import page, row

URL = "https://tabletka.by/ajax-request/reload-pharmacy-price"
ENDPOINT = URL + "/"
"""What the client actually posts to: tabletka.by 500s without the trailing slash (B17)."""


def PAGE(price: str) -> str:
    return page(row("Аспирин", "100мг", f"от {price} р."))


@pytest.fixture
def config_file(tmp_path: Path) -> Path:
    path = tmp_path / "config.json"
    path.write_text(
        json.dumps(
            {
                "profiles": {
                    "Основной": {
                        "Аптека 1": "https://tabletka.by/pharmacies/111",
                        "Аптека 2": "https://tabletka.by/pharmacies/222",
                    },
                    "Пустой": {},
                },
                "settings": {"fileName": "out.xlsx", "title": "Тест"},
                "request": {
                    "url": URL,
                    "headers": {"Cookie": "PHPSESSID=abc; lim-result=5000"},
                    "data": {"_csrf": "token"},
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return path


def stub_two_pharmacies(mocked: aioresponses) -> None:
    for price in ("5,00", "6,50"):
        mocked.post(ENDPOINT, payload={"priceCount": 1, "data": ""})
        mocked.post(ENDPOINT, payload={"priceCount": 1, "data": PAGE(price)})


def test_cli_writes_a_workbook(config_file: Path, tmp_path: Path) -> None:
    output = tmp_path / "report.xlsx"
    with aioresponses() as mocked:
        stub_two_pharmacies(mocked)
        assert main(["--config", str(config_file), "--output", str(output)]) == 0

    sheet = load_workbook(output)["Данные"]
    assert [cell.value for cell in sheet[3]] == ["Название", "Аптека 1", "Аптека 2", "Разница"]
    assert [cell.value for cell in sheet[4]] == ["Аспирин, 100мг", 5.0, 6.5, 1.5]


def test_cli_defaults_to_the_first_profile(config_file: Path, tmp_path: Path) -> None:
    output = tmp_path / "report.xlsx"
    with aioresponses() as mocked:
        stub_two_pharmacies(mocked)
        assert main(["--config", str(config_file), "--output", str(output)]) == 0
    assert output.exists()


def test_cli_reports_an_unknown_profile(config_file: Path, capsys: pytest.CaptureFixture) -> None:
    assert main(["--config", str(config_file), "--profile", "Нет такого"]) == 1
    assert "available profiles: Основной, Пустой" in capsys.readouterr().err


def test_cli_reports_an_empty_profile(config_file: Path, capsys: pytest.CaptureFixture) -> None:
    """B4: an empty profile used to crash inside a worker thread with Pool(0)."""
    assert main(["--config", str(config_file), "--profile", "Пустой"]) == 1
    assert "no pharmacies" in capsys.readouterr().err


def test_cli_reports_a_missing_config(tmp_path: Path, capsys: pytest.CaptureFixture) -> None:
    assert main(["--config", str(tmp_path / "nope.json")]) == 1
    assert "not found" in capsys.readouterr().err


def test_cli_reports_a_scrape_failure(config_file: Path, tmp_path: Path, capsys: pytest.CaptureFixture) -> None:
    """B6: scrape failures reach the user instead of vanishing."""
    with aioresponses() as mocked:
        for _ in range(6):
            mocked.post(ENDPOINT, status=500)
        assert main(["--config", str(config_file), "--output", str(tmp_path / "x.xlsx")]) == 1
    assert "could not be parsed" in capsys.readouterr().err


def test_env_cookie_reaches_the_request(
    config_file: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("PHARMPARSER_COOKIE", "PHPSESSID=from-env; lim-result=5000")
    with aioresponses() as mocked:
        stub_two_pharmacies(mocked)
        assert main(["--config", str(config_file), "--output", str(tmp_path / "r.xlsx")]) == 0
        sent = mocked.requests[("POST", __import__("aiohttp").helpers.URL(ENDPOINT))][0]
    assert "from-env" in sent.kwargs["headers"]["Cookie"]


@pytest.mark.skipif(supports_excel_macros(), reason="Excel is available, so macros are not skipped")
def test_macros_flag_degrades_gracefully_off_windows(
    config_file: Path, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    output = tmp_path / "report.xlsx"
    with aioresponses() as mocked, caplog.at_level("WARNING"):
        stub_two_pharmacies(mocked)
        assert main(["--config", str(config_file), "--output", str(output), "--macros"]) == 0
    assert output.exists()
    assert "Windows only" in caplog.text
