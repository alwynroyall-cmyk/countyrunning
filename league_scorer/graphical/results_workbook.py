"""Shared helpers for locating and reading latest results workbooks."""

from __future__ import annotations

from pathlib import Path
import re

import pandas as pd
from ..race_processor import extract_race_number

from ..output_layout import build_output_paths


def find_latest_results_workbook(output_dir: Path | None) -> Path | None:
    """Return latest standings workbook from outputs/publish/standings."""
    if not output_dir or not output_dir.exists():
        return None

    standings_dir = build_output_paths(output_dir).publish_standings_dir
    if not standings_dir.exists():
        return None

    candidates: list[tuple[int, Path]] = []
    for path in standings_dir.glob("*.xlsx"):
        match = re.search(r"\bR(\d{1,2})\b", path.stem, re.IGNORECASE)
        if not match:
            continue
        race_number = int(match.group(1))
        candidates.append((race_number, path))

    if not candidates:
        return None

    return max(candidates, key=lambda item: item[0])[1]


def sorted_race_sheet_names(xl: pd.ExcelFile) -> list[str]:
    """Return race sheet names sorted by race number."""
    return sorted(
        [name for name in xl.sheet_names if name.startswith("Race ")],
        key=lambda sheet_name: extract_race_number(sheet_name) or 0,
    )


def display_race_sheet_name(sheet_name: str) -> str:
    """Return a user-facing race sheet label like 'R1' or 'R1 - Name'."""
    match = re.match(r"^Race\s+(\d+)(.*)$", sheet_name, flags=re.IGNORECASE)
    if not match:
        return sheet_name
    number = match.group(1)
    suffix = match.group(2) or ""
    return f"R{number}{suffix}"


def display_race_column_header(header: str) -> str:
    """Convert race column labels from 'Race N' to 'RN'."""
    if not isinstance(header, str):
        return header
    return re.sub(r"\bRace\s+(\d+)\b", r"R\1", header, flags=re.IGNORECASE)
