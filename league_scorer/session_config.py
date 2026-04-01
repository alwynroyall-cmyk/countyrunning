"""
session_config.py — Centralised session configuration singleton.

Holds the current season year, data root, and derived input/output paths.
All GUI components read from and write to the shared `config` instance.

Directory convention:
    {data_root}/
        {year}/
            inputs/     ← race result files + events spreadsheet
            outputs/    ← generated reports
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path
from typing import Optional

# Persisted settings file lives in the user's home directory.
_PREFS_FILE = Path.home() / ".wrrl_prefs.json"


class SessionConfig:
    """Single source of truth for the active season and folder paths."""

    def __init__(self) -> None:
        self._data_root: Optional[Path] = None
        self._year: int = datetime.date.today().year
        self._events_path: Optional[Path] = None
        self._events_filename: Optional[str] = None
        self.load()  # restore previous session settings

    # ── year ──────────────────────────────────────────────────────────────────

    @property
    def year(self) -> int:
        return self._year

    @year.setter
    def year(self, value: int) -> None:
        self._year = int(value)
        self._sync_events_path()
        self.save()

    # ── data root ─────────────────────────────────────────────────────────────

    @property
    def data_root(self) -> Optional[Path]:
        return self._data_root

    @data_root.setter
    def data_root(self, value: Path) -> None:
        self._data_root = Path(value)
        self._sync_events_path()
        self.save()

    # ── derived paths ─────────────────────────────────────────────────────────

    @property
    def input_dir(self) -> Optional[Path]:
        if self._data_root is None:
            return None
        return self._data_root / str(self._year) / "inputs"

    @property
    def output_dir(self) -> Optional[Path]:
        if self._data_root is None:
            return None
        return self._data_root / str(self._year) / "outputs"

    # ── events file ───────────────────────────────────────────────────────────

    @property
    def events_path(self) -> Optional[Path]:
        return self._events_path

    @events_path.setter
    def events_path(self, value: Optional[Path]) -> None:
        self._events_path = Path(value) if value else None
        self._events_filename = self._events_path.name if self._events_path else None
        self.save()

    @property
    def events_filename(self) -> Optional[str]:
        return self._events_filename

    # ── state checks ──────────────────────────────────────────────────────────

    @property
    def is_configured(self) -> bool:
        """True once a data root has been selected."""
        return self._data_root is not None

    def ensure_dirs(self) -> None:
        """Create input and output directories if they do not yet exist."""
        if self.input_dir:
            self.input_dir.mkdir(parents=True, exist_ok=True)
        if self.output_dir:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        self._sync_events_path()
        self.save()

    def _sync_events_path(self) -> None:
        """Rebuild the events file path for the active season when possible."""
        if not self._events_filename:
            self._events_path = None
            return
        if self.input_dir is None:
            self._events_path = None
            return
        self._events_path = self.input_dir / self._events_filename

    # ── persistence ──────────────────────────────────────────────────────────

    def save(self) -> None:
        """Write year and data_root to the preferences file."""
        prefs: dict = {"year": self._year}
        if self._data_root is not None:
            prefs["data_root"] = str(self._data_root)
        if self._events_filename:
            prefs["events_filename"] = self._events_filename
        if self._events_path is not None:
            prefs["events_path"] = str(self._events_path)
        try:
            _PREFS_FILE.write_text(json.dumps(prefs, indent=2), encoding="utf-8")
        except OSError:
            pass  # non-fatal; preferences just won't persist

    def load(self) -> None:
        """Restore year and data_root from the preferences file (if it exists)."""
        if not _PREFS_FILE.exists():
            return
        try:
            prefs = json.loads(_PREFS_FILE.read_text(encoding="utf-8"))
            if "year" in prefs:
                self._year = int(prefs["year"])
            if "data_root" in prefs:
                candidate = Path(prefs["data_root"])
                if candidate.exists():
                    self._data_root = candidate
            if "events_filename" in prefs:
                self._events_filename = str(prefs["events_filename"])
            elif "events_path" in prefs:
                self._events_filename = Path(prefs["events_path"]).name
            self._sync_events_path()
        except (OSError, json.JSONDecodeError, ValueError):
            pass  # corrupt file; start fresh

    # ── available years ───────────────────────────────────────────────────────

    @staticmethod
    def available_years() -> list[int]:
        """Return years from 2020 to the current year + 2."""
        current = datetime.date.today().year
        return list(range(2020, current + 3))


# Module-level singleton — import `config` from any module that needs it
config = SessionConfig()
