"""Shared helpers for locating and reading latest results workbooks."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from ..race_processor import extract_race_number


def find_latest_results_workbook(output_dir: Path | None) -> Path | None:
    """Return latest `Race N -- Results.xlsx` workbook from an output directory."""
    if not output_dir or not output_dir.exists():
        return None

    candidates: list[tuple[int, Path]] = []
    for path in output_dir.glob("*.xlsx"):
        if not path.name.lower().endswith("-- results.xlsx"):
            continue
        race_number = extract_race_number(path.stem)
        if race_number is None:
            continue
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
