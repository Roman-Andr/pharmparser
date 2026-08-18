"""Tests for the pydantic configuration schema and loader."""

import json
from pathlib import Path
from typing import Any

import pytest
from pydantic import ValidationError

from pharmparser.config import AppConfig, ConfigError, ExportSettings, PharmacyEntry, Profile, RequestConfig
from pharmparser.config.env import EnvOverrides
from pharmparser.config.loader import load_config, save_config

RAW: dict[str, Any] = {
    "profiles": {
        "Мой профиль": {
            "Аптека 1": "https://tabletka.by/pharmacies/111",
            "Аптека 2": "https://tabletka.by/pharmacies/222",
        }
    },
    "settings": {
        "green": "19CF1F",
        "red": "E81737",
        "title": "Тест",
        "fileName": "data.xlsx",
        "colWidth": 50,
        "cellWidth": 15,
        "diffWidth": 10,
    },
    "request": {
        "url": "https://tabletka.by/ajax-request/reload-pharmacy-price",
        "headers": {"Cookie": "PHPSESSID=abc; lim-result=5000"},
        "data": {"sort": "name", "_csrf": "token"},
    },
}


def write(tmp_path: Path, raw: dict[str, Any]) -> Path:
    path = tmp_path / "config.json"
    path.write_text(json.dumps(raw, ensure_ascii=False), encoding="utf-8")
    return path


# -- schema -------------------------------------------------------------------


def test_disk_mapping_becomes_named_profiles() -> None:
    config = AppConfig.model_validate(RAW)
    (profile,) = config.profiles
    assert profile.name == "Мой профиль"
    assert [entry.name for entry in profile.pharmacies] == ["Аптека 1", "Аптека 2"]
    assert [entry.pharmacy_id for entry in profile.pharmacies] == ["111", "222"]


def test_camel_case_settings_are_read_into_snake_case() -> None:
    settings = AppConfig.model_validate(RAW).settings
    assert settings.file_name == "data.xlsx"
    assert settings.col_width == 50
    assert settings.macro_file_name == "data.xlsm"


def test_settings_fall_back_to_defaults() -> None:
    config = AppConfig.model_validate({**RAW, "settings": {}})
    assert config.settings == ExportSettings()


def test_colours_accept_a_leading_hash() -> None:
    assert ExportSettings.model_validate({"green": "#19CF1F"}).green == "19CF1F"


def test_invalid_colour_is_rejected() -> None:
    with pytest.raises(ValidationError, match="green"):
        ExportSettings.model_validate({"green": "not-a-colour"})


def test_non_positive_widths_are_rejected() -> None:
    with pytest.raises(ValidationError):
        ExportSettings.model_validate({"colWidth": 0})


def test_unknown_settings_key_is_reported() -> None:
    with pytest.raises(ValidationError, match="colWith"):
        ExportSettings.model_validate({"colWith": 10})


def test_pharmacy_url_must_end_with_a_numeric_id() -> None:
    with pytest.raises(ValidationError, match="numeric pharmacy id"):
        PharmacyEntry(name="A", url="https://tabletka.by/pharmacies/not-a-number")


def test_blank_rows_are_allowed_but_not_scrapable() -> None:
    """The UI creates an empty row whenever "Add" is pressed."""
    profile = Profile(name="P", pharmacies=[PharmacyEntry(), PharmacyEntry(name="A", url="https://x.by/1")])
    assert len(profile.pharmacies) == 2
    assert [entry.name for entry in profile.scrapable] == ["A"]


def test_duplicate_pharmacy_names_in_a_profile_are_rejected() -> None:
    """B14: names label report columns, so they must be distinguishable."""
    with pytest.raises(ValidationError, match="duplicate pharmacy names"):
        Profile(
            name="P",
            pharmacies=[
                PharmacyEntry(name="Аптека", url="https://tabletka.by/pharmacies/1"),
                PharmacyEntry(name="Аптека", url="https://tabletka.by/pharmacies/2"),
            ],
        )


def test_request_requires_a_cookie() -> None:
    with pytest.raises(ValidationError, match="Cookie header"):
        RequestConfig(headers={})


def test_request_cookie_lookup_is_case_insensitive() -> None:
    assert RequestConfig(headers={"cookie": "x=1"}).cookie == "x=1"


# -- round trip ---------------------------------------------------------------


def test_serialisation_preserves_the_on_disk_shape() -> None:
    assert AppConfig.model_validate(RAW).model_dump() == RAW


def test_profile_names_survive_a_round_trip(tmp_path: Path) -> None:
    """A6: names used to be regenerated as "Profile 1..N" on every save."""
    path = write(tmp_path, RAW)
    save_config(load_config(path), path)
    assert list(json.loads(path.read_text(encoding="utf-8"))["profiles"]) == ["Мой профиль"]


def test_save_is_atomic(tmp_path: Path) -> None:
    path = write(tmp_path, RAW)
    save_config(load_config(path), path)
    # No stray temporary files left behind.
    assert [p.name for p in tmp_path.iterdir()] == ["config.json"]


# -- loader errors -------------------------------------------------------------


def test_missing_config_gives_an_actionable_message(tmp_path: Path) -> None:
    with pytest.raises(ConfigError, match=r"config\.json\.example"):
        load_config(tmp_path / "config.json")


def test_malformed_json_is_reported_as_such(tmp_path: Path) -> None:
    path = tmp_path / "config.json"
    path.write_text("{oh no", encoding="utf-8")
    with pytest.raises(ConfigError, match="not valid JSON"):
        load_config(path)


def test_validation_errors_name_the_offending_field(tmp_path: Path) -> None:
    broken = {**RAW, "settings": {**RAW["settings"], "colWidth": -5}}
    with pytest.raises(ConfigError, match=r"settings\.colWidth"):
        load_config(write(tmp_path, broken))


# -- environment overrides -----------------------------------------------------


def test_env_overrides_replace_the_cookie(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PHARMPARSER_COOKIE", "PHPSESSID=from-env")
    config = load_config(write(tmp_path, RAW))
    assert config.request.cookie == "PHPSESSID=from-env"


def test_env_overrides_replace_the_csrf_token(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PHARMPARSER_CSRF", "fresh")
    config = load_config(write(tmp_path, RAW))
    assert config.request.data["_csrf"] == "fresh"
    assert config.request.data["sort"] == "name"


def test_absent_env_leaves_the_file_untouched() -> None:
    config = AppConfig.model_validate(RAW)
    assert EnvOverrides(cookie=None, csrf=None, file_name=None).apply(config) == config
