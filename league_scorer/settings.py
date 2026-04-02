"""
settings.py — Centralised league configuration and settings management.

This module holds all league rule constants and provides a simple interface for loading/saving settings.
"""
import json
from pathlib import Path
from typing import Any, Dict

SETTINGS_FILE = Path.home() / ".wrrl_settings.json"

# Default league settings (initial values from current hardcoded constants)
DEFAULT_SETTINGS = {
    "BEST_N": 6,              # Best N scores count
    "MAX_RACES": 8,           # Number of races in the season
    "TEAM_SIZE": 5,           # Top N scorers per team
    "MAX_DIV_PTS": 20,        # Max team points per race
    "SEASON_FINAL_RACE": 8,   # Race number for end-of-season narrative
}

class LeagueSettings:
    def __init__(self):
        self._settings = DEFAULT_SETTINGS.copy()
        self.load()

    def load(self):
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
                self._settings.update({k: data[k] for k in DEFAULT_SETTINGS if k in data})
            except Exception:
                pass  # Ignore errors, use defaults

    def save(self):
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(self._settings, f, indent=2)
        except OSError as exc:
            raise RuntimeError(
                f"Unable to save WRRL settings to '{SETTINGS_FILE}'. {exc}"
            ) from exc

    def get(self, key: str) -> Any:
        return self._settings.get(key, DEFAULT_SETTINGS.get(key))

    def set(self, key: str, value: Any):
        if key in DEFAULT_SETTINGS:
            self._settings[key] = value
            self.save()

    def as_dict(self) -> Dict[str, Any]:
        return self._settings.copy()

# Singleton instance
settings = LeagueSettings()
