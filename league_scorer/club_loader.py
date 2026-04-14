"""Load and validate clubs.xlsx."""

import logging
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd

from .exceptions import FatalError
from .models import ClubInfo

log = logging.getLogger(__name__)

_REQUIRED = {"Club", "Preferred name", "Team A", "Team B"}
_OPTIONAL_ALIAS_COLUMNS = {"ea club name", "ea club", "ea clubname"}


def load_clubs(filepath: Path) -> Tuple[Dict[str, str], Dict[str, ClubInfo]]:
    """
    Load clubs.xlsx.

    Returns
    -------
    raw_to_preferred : dict  lower(raw_name) → preferred_name
    club_info        : dict  preferred_name  → ClubInfo

    Raises FatalError on missing file or malformed content.
    """
    if not filepath.exists():
        raise FatalError(f"clubs.xlsx not found: '{filepath}'")

    try:
        df = pd.read_excel(filepath, engine="openpyxl", dtype=str)
    except Exception as exc:
        raise FatalError(f"Cannot read clubs.xlsx: {exc}") from exc

    df.columns = [str(c).strip() for c in df.columns]
    missing = _REQUIRED - set(df.columns)
    if missing:
        raise FatalError(f"clubs.xlsx missing columns: {missing}")

    alias_columns = [c for c in df.columns if str(c).strip().lower() in _OPTIONAL_ALIAS_COLUMNS]

    raw_to_preferred: Dict[str, str] = {}
    club_info: Dict[str, ClubInfo] = {}

    for idx, row in df.iterrows():
        raw = str(row["Club"]).strip()
        preferred = str(row["Preferred name"]).strip()

        if not raw or raw.lower() == "nan":
            log.warning("clubs.xlsx row %d: empty Club name — skipped", idx + 2)
            continue
        if not preferred or preferred.lower() == "nan":
            log.warning("clubs.xlsx row %d: empty Preferred name — skipped", idx + 2)
            continue

        try:
            div_a = int(float(str(row["Team A"]).strip()))
            div_b = int(float(str(row["Team B"]).strip()))
        except (ValueError, TypeError) as exc:
            raise FatalError(
                f"clubs.xlsx row {idx + 2}: non-integer division for '{preferred}': {exc}"
            ) from exc

        if div_a not in (1, 2):
            raise FatalError(
                f"clubs.xlsx row {idx + 2}: Team A must be 1 or 2, got {div_a} for '{preferred}'"
            )
        if div_b not in (1, 2):
            raise FatalError(
                f"clubs.xlsx row {idx + 2}: Team B must be 1 or 2, got {div_b} for '{preferred}'"
            )

        raw_to_preferred[raw.lower()] = preferred

        preferred_lower = preferred.lower()
        if preferred_lower in raw_to_preferred and raw_to_preferred[preferred_lower] != preferred:
            log.warning(
                "clubs.xlsx row %d: preferred name '%s' conflicts with existing club alias mapping",
                idx + 2,
                preferred,
            )
        raw_to_preferred[preferred_lower] = preferred

        for alias_column in alias_columns:
            alias = str(row[alias_column]).strip()
            if not alias or alias.lower() == "nan":
                continue
            alias_lower = alias.lower()
            if alias_lower in raw_to_preferred and raw_to_preferred[alias_lower] != preferred:
                log.warning(
                    "clubs.xlsx row %d: alias column '%s' value '%s' conflicts with existing mapping",
                    idx + 2,
                    alias_column,
                    alias,
                )
            raw_to_preferred[alias_lower] = preferred

        if preferred not in club_info:
            club_info[preferred] = ClubInfo(
                preferred_name=preferred,
                div_a=div_a,
                div_b=div_b,
            )
        else:
            # Multiple raw names → same preferred name: verify division consistency
            existing = club_info[preferred]
            if existing.div_a != div_a or existing.div_b != div_b:
                log.warning(
                    "clubs.xlsx row %d: division mismatch for '%s' — keeping first occurrence",
                    idx + 2, preferred,
                )

    if not club_info:
        raise FatalError("clubs.xlsx contains no valid club entries")

    log.info(
        "clubs.xlsx: %d name mappings → %d clubs",
        len(raw_to_preferred), len(club_info),
    )
    return raw_to_preferred, club_info
