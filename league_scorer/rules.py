"""Runtime accessors for league rules/settings.

Use these functions instead of hard-coded constants so the GUI-controlled
settings (via `league_scorer.settings`) are always the single source of
truth and callers don't need to import `settings` directly.
"""
from typing import Any

from .settings import settings, DEFAULT_SETTINGS


def _get_int(key: str, default: Any) -> int:
    try:
        return int(settings.get(key) if settings.get(key) is not None else default)
    except Exception:
        # Fallback to default if saved value is malformed
        return int(default)


def get_best_n() -> int:
    return _get_int("BEST_N", DEFAULT_SETTINGS["BEST_N"])


def get_max_races() -> int:
    return _get_int("MAX_RACES", DEFAULT_SETTINGS["MAX_RACES"])


def get_team_size() -> int:
    return _get_int("TEAM_SIZE", DEFAULT_SETTINGS["TEAM_SIZE"])


def get_max_div_pts() -> int:
    return _get_int("MAX_DIV_PTS", DEFAULT_SETTINGS["MAX_DIV_PTS"])


def get_season_final_race() -> int:
    return _get_int("SEASON_FINAL_RACE", DEFAULT_SETTINGS["SEASON_FINAL_RACE"])
