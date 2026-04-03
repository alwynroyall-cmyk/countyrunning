"""Shared helpers for locating and reading latest results workbooks."""

from __future__ import annotations

from pathlib import Path
import re

import pandas as pd

from ..output_layout import build_output_paths


def find_latest_results_workbook(output_dir: Path | None) -> Path | None:
    """Return latest standings workbook from outputs/publish/xlsx/standings."""
    if not output_dir or not output_dir.exists():
        return None

    standings_dir = build_output_paths(output_dir).publish_xlsx_standings_dir
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
