"""Load a race Excel file, normalise all fields, and handle duplicates."""

import logging
import math
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .exceptions import RaceProcessingError
from .models import CategoryRecord, RaceIssue, RunnerRaceEntry, UnrecognisedClub
from .normalisation import (
    find_time_column,
    normalise_category,
    normalise_gender,
    parse_time_to_seconds,
    time_display,
)
from .race_validation import validate_race_schema
from .source_loader import load_race_dataframe

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────── public interface ────

def extract_race_number(filename: str) -> Optional[int]:
    """Return the positive integer race number from a filename, or None."""
    m = re.search(r"\brace\s*#?\s*(\d+)\b", filename, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return n if n > 0 else None
    return None


def process_race_file(
    filepath: Path,
    race_number: int,
    raw_to_preferred: Dict[str, str],
) -> Tuple[List[RunnerRaceEntry], List[CategoryRecord], List[UnrecognisedClub], List[RaceIssue]]:
    """
    Load and validate a race file.

    Returns
    -------
    runners      : normalised runner entries (duplicates already resolved)
    cat_records  : category exception report rows
    unrec_clubs  : unrecognised club report rows
    issue_notes  : race-specific issues captured for output summaries

    Raises RaceProcessingError on fatal race-level conditions.
    """
    log.info("Loading Race %d: %s", race_number, filepath.name)

    try:
        # Do NOT force dtype=str — preserves native time / numeric types
        df = load_race_dataframe(filepath)
    except Exception as exc:
        raise RaceProcessingError(f"Cannot read '{filepath.name}': {exc}") from exc

    df.columns = [str(c).strip() for c in df.columns]

    schema_validation = validate_race_schema(df, filepath)
    col_map = schema_validation.column_map
    time_col = schema_validation.time_column
    log.info("  Time column: '%s'", time_col)

    # accumulators for exception reports
    cat_tracker: Dict[str, Tuple[str, str, int]] = {}   # raw_lower → (norm, notes, count)
    unrec_tracker: Dict[str, Tuple[str, int]] = {}       # raw_lower → (display, count)

    runners: List[RunnerRaceEntry] = []
    issue_notes: List[RaceIssue] = list(schema_validation.issues)
    skipped = 0

    for idx, row in df.iterrows():
        row_num = idx + 2  # 1-based with header

        # ── Name ──
        name_raw = row.get(col_map["Name"])
        name_s = "" if name_raw is None else str(name_raw).strip()
        if not name_s or name_s.lower() == "nan":
            missing_time_seconds = parse_time_to_seconds(row.get(time_col))
            if missing_time_seconds is not None:
                raw_gender = row.get(col_map["Gender"])
                raw_gender_text = "" if raw_gender is None else str(raw_gender).strip()
                raw_club_val = row.get(col_map["Club"])
                raw_club_text = "" if raw_club_val is None else str(raw_club_val).strip()
                raw_cat_val = row.get(col_map["Category"])
                raw_cat_text = "" if raw_cat_val is None else str(raw_cat_val).strip()
                issue_notes.append(
                    RaceIssue(
                        "warning",
                        "Missing runner name with valid time",
                        row_num,
                        "AUD-ROW-005",
                        raw_club=raw_club_text,
                        gender=raw_gender_text,
                        raw_category=raw_cat_text,
                        time_str=time_display(row.get(time_col)),
                    )
                )
            skipped += 1
            continue
        name = name_s

        # ── Club / eligibility ──
        raw_club_val = row.get(col_map["Club"])
        raw_club = "" if raw_club_val is None else str(raw_club_val).strip()
        if raw_club.lower() == "nan":
            raw_club = ""

        preferred_club = raw_to_preferred.get(raw_club.lower()) if raw_club else None
        eligible = preferred_club is not None

        raw_gender_val = row.get(col_map["Gender"])
        raw_cat_val = row.get(col_map["Category"])
        raw_cat = (
            ""
            if (raw_cat_val is None or str(raw_cat_val).strip().lower() == "nan")
            else str(raw_cat_val).strip()
        )
        raw_time_val = row.get(time_col)

        if not eligible:
            key = raw_club.lower()
            display_name, count = unrec_tracker.get(key, (raw_club or "(blank)", 0))
            unrec_tracker[key] = (display_name, count + 1)
            if raw_club:
                log.debug(
                    "  Row %d (%s): unrecognised club '%s' — excluded",
                    row_num, name, raw_club,
                )

            gender_display = "" if raw_gender_val is None else str(raw_gender_val).strip()
            time_seconds = parse_time_to_seconds(raw_time_val)
            disp_time = time_display(raw_time_val) if time_seconds is not None else (
                "" if raw_time_val is None else str(raw_time_val).strip()
            )

            runners.append(
                RunnerRaceEntry(
                    name=name,
                    raw_club=raw_club,
                    preferred_club=None,
                    gender=gender_display,
                    raw_category=raw_cat,
                    normalised_category=raw_cat,
                    time_str=disp_time,
                    time_seconds=time_seconds if time_seconds is not None else math.inf,
                    race_number=race_number,
                    eligible=False,
                    source_row=row_num,
                )
            )
            continue

        # ── Gender ──
        gender = normalise_gender(raw_gender_val)
        if gender is None:
            issue_notes.append(
                RaceIssue(
                    "warning",
                    f"invalid gender '{raw_gender_val}'",
                    row_num,
                    "AUD-ROW-001",
                    runner_name=name,
                    raw_club=raw_club,
                    raw_category=raw_cat,
                )
            )
            log.warning(
                "  Row %d (%s): invalid gender '%s' — row skipped",
                row_num, name, raw_gender_val,
            )
            skipped += 1
            continue

        # ── Time ──
        time_seconds = parse_time_to_seconds(raw_time_val)
        if time_seconds is None:
            issue_notes.append(
                RaceIssue(
                    "warning",
                    f"invalid time '{raw_time_val}'",
                    row_num,
                    "AUD-ROW-002",
                    runner_name=name,
                    raw_club=raw_club,
                    gender=gender or "",
                    raw_category=raw_cat,
                )
            )
            log.warning(
                "  Row %d (%s): invalid time '%s' — row skipped",
                row_num, name, raw_time_val,
            )
            skipped += 1
            continue
        disp_time = time_display(raw_time_val)

        # ── Category ──
        norm_cat, cat_notes = normalise_category(raw_cat)
        row_warnings: List[str] = []
        if cat_notes:
            issue_notes.append(
                RaceIssue(
                    "warning",
                    cat_notes,
                    row_num,
                    "AUD-ROW-003" if raw_cat == "" else "AUD-ROW-004",
                    runner_name=name,
                    raw_club=raw_club,
                    gender=gender,
                    raw_category=raw_cat,
                    time_str=disp_time,
                )
            )
            row_warnings.append(cat_notes)

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
                source_row=row_num,
                warnings=row_warnings,
            )
        )

    log.info("  Parsed %d runner rows (%d skipped)", len(runners), skipped)

    # ── Deduplication ──
    runners = _deduplicate(runners, race_number, issue_notes)

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

    return runners, cat_records, unrec_clubs, issue_notes


# ──────────────────────────────────────────────────────────── deduplication ──

def _deduplicate(
    runners: List[RunnerRaceEntry], race_number: int, issue_notes: List[RaceIssue]
) -> List[RunnerRaceEntry]:
    """
    Keep the fastest time when the same (Name, Preferred Club) appears more
    than once.

    Ineligible runners are left unchanged, even if duplicate rows exist for the
    same name and raw club.

    Logs a warning only when an eligible duplicate pair is reduced.
    """
    # key → best RunnerRaceEntry so far
    best: Dict[Tuple, RunnerRaceEntry] = {}
    order_pos: Dict[Tuple, int] = {}
    ordered: List[RunnerRaceEntry] = []

    for r in runners:
        if not r.eligible:
            ordered.append(r)
            continue

        club_key = r.preferred_club
        key = (r.name.lower(), club_key)

        if key not in best:
            best[key] = r
            order_pos[key] = len(ordered)
            ordered.append(r)
        else:
            existing = best[key]
            if r.time_seconds < existing.time_seconds:
                warning_text = (
                    f"Duplicate runner: kept {r.time_str}, discarded {existing.time_str}"
                )
                code = "AUD-ROW-008"
                if _has_duplicate_attribute_conflict(existing, r):
                    code = "AUD-ROW-010"
                    warning_text = (
                        "Duplicate runner attribute conflict: "
                        f"kept {r.time_str} (row {r.source_row}), "
                        f"discarded {existing.time_str} (row {existing.source_row})"
                    )

                issue_notes.append(
                    RaceIssue(
                        "warning",
                        warning_text,
                        r.source_row,
                        code,
                        runner_name=r.name,
                        raw_club=r.raw_club,
                        gender=r.gender,
                        raw_category=r.raw_category,
                        time_str=r.time_str,
                    )
                )
                log.warning(
                    "  Race %d duplicate '%s' (%s): keeping %s, discarding %s",
                    race_number, r.name,
                    r.preferred_club or r.raw_club,
                    r.time_str, existing.time_str,
                )
                r.warnings.append(warning_text)
                best[key] = r
                ordered[order_pos[key]] = r
            else:
                code = "AUD-ROW-008"
                warning_text = (
                    f"Duplicate runner: kept {existing.time_str}, discarded {r.time_str}"
                )
                if _has_duplicate_attribute_conflict(existing, r):
                    code = "AUD-ROW-010"
                    warning_text = (
                        "Duplicate runner attribute conflict: "
                        f"kept {existing.time_str} (row {existing.source_row}), "
                        f"discarded {r.time_str} (row {r.source_row})"
                    )

                issue_notes.append(
                    RaceIssue(
                        "warning",
                        warning_text,
                        existing.source_row,
                        code,
                        runner_name=existing.name,
                        raw_club=existing.raw_club,
                        gender=existing.gender,
                        raw_category=existing.raw_category,
                        time_str=existing.time_str,
                    )
                )
                log.warning(
                    "  Race %d duplicate '%s' (%s): keeping %s, discarding %s",
                    race_number, r.name,
                    r.preferred_club or r.raw_club,
                    existing.time_str, r.time_str,
                )
                existing.warnings.append(warning_text)

    return ordered


def _has_duplicate_attribute_conflict(
    existing: RunnerRaceEntry, candidate: RunnerRaceEntry,
) -> bool:
    return any(
        (
            existing.gender != candidate.gender,
            existing.normalised_category != candidate.normalised_category,
            existing.preferred_club != candidate.preferred_club,
        )
    )

