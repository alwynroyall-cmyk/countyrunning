from importlib import import_module
from pathlib import Path

settings_module = import_module("league_scorer.settings")
from league_scorer.settings import LeagueSettings


def test_league_settings_save_and_load(tmp_path: Path, monkeypatch):
    settings_path = tmp_path / "wrrl_settings.json"
    monkeypatch.setattr(settings_module, "SETTINGS_FILE", settings_path)

    config = LeagueSettings()
    assert config.get("BEST_N") == 6
    assert config.get("MAX_RACES") == 8

    config.set("BEST_N", 4)
    assert config.get("BEST_N") == 4
    assert settings_path.exists()

    loaded = LeagueSettings()
    assert loaded.get("BEST_N") == 4


def test_league_settings_load_invalid_json(tmp_path: Path, monkeypatch):
    settings_path = tmp_path / "wrrl_settings.json"
    settings_path.write_text("{ not valid json }", encoding="utf-8")
    monkeypatch.setattr(settings_module, "SETTINGS_FILE", settings_path)

    config = LeagueSettings()
    assert config.get("BEST_N") == 6
