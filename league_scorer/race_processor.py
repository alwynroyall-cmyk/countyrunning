"""Load a race Excel file, normalise all fields, and handle duplicates."""

import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .exceptions import RaceProcessingError
from .models import CategoryRecord, RunnerRaceEntry, UnrecognisedClub
from .normalisation import (
    find_time_column,
    normalise_category,
    normalise_gender,
    parse_time_to_seconds,
    time_display,
)

log = logging.getLogger(__name__)

_REQUIRED_COLS = {"Position", "Name", "Club", "Gender", "Category"}


# ─────────────────────────────────────────────────────── public interface ────

def extract_race_number(filename: str) -> Optional[int]:
    """Return the positive integer race number from a filename, or None."""
    m = re.search(r"\brace\s*(\d+)\b", filename, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return n if n > 0 else None
    return None


def process_race_file(
    filepath: Path,
    race_number: int,
    raw_to_preferred: Dict[str, str],
) -> Tuple[List[RunnerRaceEntry], List[CategoryRecord], List[UnrecognisedClub]]:
    """
    Load and validate a race file.

    Returns
    -------
    runners      : normalised runner entries (duplicates already resolved)
    cat_records  : category exception report rows
    unrec_clubs  : unrecognised club report rows

    Raises RaceProcessingError on fatal race-level conditions.
    """
    log.info("Loading Race %d: %s", race_number, filepath.name)

    try:
        # Do NOT force dtype=str — preserves native time / numeric types
        df = pd.read_excel(filepath, engine="openpyxl")
    except Exception as exc:
        raise RaceProcessingError(f"Cannot read '{filepath.name}': {exc}") from exc

    df.columns = [str(c).strip() for c in df.columns]

    # ── Required column detection (case-insensitive) ──
    col_lower = {c.lower(): c for c in df.columns}
    col_map: Dict[str, str] = {}
    missing_cols = []
    for req in _REQUIRED_COLS:
        found = col_lower.get(req.lower())
        if found is None:
            missing_cols.append(req)
        else:
            col_map[req] = found

    if missing_cols:
        raise RaceProcessingError(
            f"'{filepath.name}' missing required columns: {missing_cols}"
        )

    # ── Time column ──
    time_col = find_time_column(list(df.columns))
    if time_col is None:
        raise RaceProcessingError(
            f"'{filepath.name}' has no time-like column (need Chip Time / Gun Time / *time*)"
        )
    log.info("  Time column: '%s'", time_col)

    # accumulators for exception reports
    cat_tracker: Dict[str, Tuple[str, str, int]] = {}   # raw_lower → (norm, notes, count)
    unrec_tracker: Dict[str, Tuple[str, int]] = {}       # raw_lower → (display, count)

    runners: List[RunnerRaceEntry] = []
    skipped = 0

    for idx, row in df.iterrows():
        row_num = idx + 2  # 1-based with header

        # ── Name ──
        name_raw = row.get(col_map["Name"])
        name_s = "" if name_raw is None else str(name_raw).strip()
        if not name_s or name_s.lower() == "nan":
            skipped += 1
            continue
        name = name_s

        # ── Gender ──
        gender = normalise_gender(row.get(col_map["Gender"]))
        if gender is None:
            log.warning(
                "  Row %d (%s): invalid gender '%s' — row skipped",
                row_num, name, row.get(col_map["Gender"]),
            )
            skipped += 1
            continue

        # ── Time ──
        raw_time_val = row.get(time_col)
        time_seconds = parse_time_to_seconds(raw_time_val)
        if time_seconds is None:
            log.warning(
                "  Row %d (%s): invalid time '%s' — row skipped",
                row_num, name, raw_time_val,
            )
            skipped += 1
            continue
        disp_time = time_display(raw_time_val)

        # ── Club ──
        raw_club_val = row.get(col_map["Club"])
        raw_club = "" if raw_club_val is None else str(raw_club_val).strip()
        if raw_club.lower() == "nan":
            raw_club = ""

        preferred_club = raw_to_preferred.get(raw_club.lower()) if raw_club else None
        eligible = preferred_club is not None

        if not eligible:
            key = raw_club.lower()
            display_name, count = unrec_tracker.get(key, (raw_club or "(blank)", 0))
            unrec_tracker[key] = (display_name, count + 1)
            if raw_club:
                log.debug(
                    "  Row %d (%s): unrecognised club '%s' — excluded",
                    row_num, name, raw_club,
                )

        # ── Category ──
        raw_cat_val = row.get(col_map["Category"])
        raw_cat = (
            ""
            if (raw_cat_val is None or str(raw_cat_val).strip().lower() == "nan")
            else str(raw_cat_val).strip()
        )
        norm_cat, cat_notes = normalise_category(raw_cat)

        cat_key = raw_cat.lower()
        _, _, cat_count = cat_tracker.get(cat_key, ("", "", 0))
        cat_tracker[cat_key] = (norm_cat, cat_notes, cat_count + 1)

        runners.append(
            RunnerRaceEntry(
                name=name,
                raw_club=raw_club,
                preferred_club=preferred_club,
                gender=gender,
                raw_category=raw_cat,
                normalised_category=norm_cat,
                time_str=disp_time,
                time_seconds=time_seconds,
                race_number=race_number,
                eligible=eligible,
            )
        )

    log.info("  Parsed %d runner rows (%d skipped)", len(runners), skipped)

    # ── Deduplication ──
    runners = _deduplicate(runners, race_number)

    # ── Build report objects ──
    cat_records = [
        CategoryRecord(
            raw_category=k or "(blank)",
            normalised_category=v[0],
            count=v[2],
            notes=v[1],
        )
        for k, v in cat_tracker.items()
    ]

    unrec_clubs = [
        UnrecognisedClub(raw_club_name=v[0], occurrences=v[1])
        for v in unrec_tracker.values()
    ]

    return runners, cat_records, unrec_clubs


# ──────────────────────────────────────────────────────────── deduplication ──

def _deduplicate(
    runners: List[RunnerRaceEntry], race_number: int
) -> List[RunnerRaceEntry]:
    """
    Keep the fastest time when the same (Name, Preferred Club) appears more
    than once. Ineligible runners are de-duplicated by (Name, raw club).
    Logs a warning for every duplicate pair removed.
    """
    # key → best RunnerRaceEntry so far
    best: Dict[Tuple, RunnerRaceEntry] = {}

    for r in runners:
        # Use preferred_club for eligible runners, raw_club for ineligible
        club_key = r.preferred_club if r.preferred_club else f"__raw__{r.raw_club.lower()}"
        key = (r.name.lower(), club_key)

        if key not in best:
            best[key] = r
        else:
            existing = best[key]
            if r.time_seconds < existing.time_seconds:
                log.warning(
                    "  Race %d duplicate '%s' (%s): keeping %s, discarding %s",
                    race_number, r.name,
                    r.preferred_club or r.raw_club,
                    r.time_str, existing.time_str,
                )
                best[key] = r
            else:
                log.warning(
                    "  Race %d duplicate '%s' (%s): keeping %s, discarding %s",
                    race_number, r.name,
                    r.preferred_club or r.raw_club,
                    existing.time_str, r.time_str,
                )

    return list(best.values())
