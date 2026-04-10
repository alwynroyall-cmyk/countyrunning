from league_scorer import rules
from league_scorer.rules import DEFAULT_SETTINGS


def test_getters_use_default_settings_when_missing(monkeypatch):
    monkeypatch.setattr(rules.settings, "get", lambda key: None)
    assert rules.get_best_n() == DEFAULT_SETTINGS["BEST_N"]
    assert rules.get_max_races() == DEFAULT_SETTINGS["MAX_RACES"]
    assert rules.get_team_size() == DEFAULT_SETTINGS["TEAM_SIZE"]
    assert rules.get_max_div_pts() == DEFAULT_SETTINGS["MAX_DIV_PTS"]
    assert rules.get_season_final_race() == DEFAULT_SETTINGS["SEASON_FINAL_RACE"]


def test_getters_fallback_on_malformed_values(monkeypatch):
    monkeypatch.setattr(rules.settings, "get", lambda key: "not-an-int")
    assert rules.get_best_n() == DEFAULT_SETTINGS["BEST_N"]
