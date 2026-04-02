"""Shared schema validation for race input dataframes."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .exceptions import RaceProcessingError
from .models import RaceIssue
from .normalisation import find_time_column

_REQUIRED_COLS = {
    "Position": ("position", "pos", "place"),
    "Name": ("name", "runner", "runner name"),
    "Club": ("club", "team", "affiliation"),
    "Gender": ("gender", "sex"),
    "Category": ("category", "cat", "age category"),
}


@dataclass
class RaceSchemaValidation:
    column_map: dict[str, str]
    time_column: str
    issues: list[RaceIssue] = field(default_factory=list)


def validate_race_schema(df: pd.DataFrame, filepath: Path) -> RaceSchemaValidation:
    """Validate and map race dataframe schema before row-level processing."""
    columns = [str(c).strip() for c in df.columns]
    normalised = [_normalise_column_name(c) for c in columns]

    issues: list[RaceIssue] = []

    duplicates = sorted({name for name in normalised if normalised.count(name) > 1 and name})
    if duplicates:
        issues.append(
            RaceIssue(
                "other",
                "Duplicate or equivalent headers detected after normalization: " + ", ".join(duplicates),
                code="AUD-RACE-008",
            )
        )

    col_lower = {_normalise_column_name(c): c for c in columns}
    col_map: dict[str, str] = {}
    missing_cols: list[str] = []
    for req, aliases in _REQUIRED_COLS.items():
        found = next((col_lower.get(alias) for alias in aliases if alias in col_lower), None)
        if found is None:
            missing_cols.append(req)
        else:
            col_map[req] = found

    if missing_cols:
        raise RaceProcessingError(
            f"'{filepath.name}' missing required columns: {missing_cols}"
        )

    time_col = find_time_column(columns)
    if time_col is None:
        raise RaceProcessingError(
            f"'{filepath.name}' has no time-like column (need Chip Time / Gun Time / *time*)"
        )

    blank_name_rows = 0
    for value in df[col_map["Name"]].tolist():
        text = "" if value is None else str(value).strip()
        if not text or text.lower() == "nan":
            blank_name_rows += 1
    if blank_name_rows:
        issues.append(
            RaceIssue(
                "warning",
                f"Schema pre-check found {blank_name_rows} row(s) with blank runner names.",
                code="AUD-RACE-009",
            )
        )

    return RaceSchemaValidation(column_map=col_map, time_column=time_col, issues=issues)


def _normalise_column_name(value: str) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower())
