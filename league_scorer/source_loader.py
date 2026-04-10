"""Helpers for discovering and loading race source files."""

import logging
import re
from pathlib import Path
from typing import Dict, Iterable

import pandas as pd

log = logging.getLogger(__name__)

_HEADER_TOKENS = {
    "pos",
    "position",
    "race no",
    "bib",
    "bib#",
    "name",
    "gun time",
    "net time",
    "chip time",
    "category",
    "cat pos",
    "gender",
    "sex",
    "gen pos",
    "club",
}


def discover_race_files(input_dir: Path, excluded_names: Iterable[str] = ()) -> Dict[int, Path]:
    """Return discovered race files keyed by race number.

    Preference order is .xlsx, .xlsm, .xls, then .csv so existing cleaned workbooks
    continue to win when both a raw and cleaned copy are present.
    """
    excluded = {name.lower() for name in excluded_names}
    found: Dict[int, tuple[int, Path]] = {}

    for rank, pattern in enumerate(("*.xlsx", "*.xlsm", "*.xls", "*.csv")):
        for filepath in sorted(input_dir.glob(pattern)):
            if filepath.name.lower() in excluded:
                continue

            race_num = _extract_race_number(filepath.stem)
            if race_num is None:
                continue

            existing = found.get(race_num)
            if existing is None:
                found[race_num] = (rank, filepath)
                continue

            current_rank, current_path = existing
            if rank < current_rank:
                log.warning(
                    "Duplicate race number %d — preferring '%s' over '%s'.",
                    race_num,
                    filepath.name,
                    current_path.name,
                )
                found[race_num] = (rank, filepath)
            else:
                log.warning(
                    "Duplicate race number %d — ignoring '%s' (keeping '%s').",
                    race_num,
                    filepath.name,
                    current_path.name,
                )

    return {race_num: filepath for race_num, (_, filepath) in sorted(found.items())}


def load_race_dataframe(filepath: Path) -> pd.DataFrame:
    """Load a race file from Excel, CSV, or HTML-table disguised as .xls."""
    suffix = filepath.suffix.lower()
    errors = []

    if suffix == ".csv":
        return _normalise_loaded_dataframe(pd.read_csv(filepath, dtype=str))

    engine = "openpyxl" if suffix in {".xlsx", ".xlsm"} else "xlrd"
    try:
        df = pd.read_excel(filepath, engine=engine)
        return _normalise_loaded_dataframe(df)
    except Exception as exc:
        errors.append(f"{engine}: {exc}")

    if not _looks_like_html_table(filepath):
        raise ValueError("; ".join(errors))

    for encoding in ("utf-16", "utf-8", "cp1252"):
        try:
            tables = pd.read_html(filepath, encoding=encoding)
            if tables:
                return _normalise_loaded_dataframe(tables[0])
        except Exception as exc:
            errors.append(f"html/{encoding}: {exc}")

    raise ValueError("; ".join(errors))


def _normalise_loaded_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    normalised = df.copy()
    normalised.columns = [str(column).strip() for column in normalised.columns]

    if _should_promote_first_row(normalised):
        promoted = [str(value).strip() for value in normalised.iloc[0].tolist()]
        normalised = normalised.iloc[1:].reset_index(drop=True)
        normalised.columns = promoted

    normalised.columns = [str(column).strip() for column in normalised.columns]
    return normalised


def _should_promote_first_row(df: pd.DataFrame) -> bool:
    columns = list(df.columns)
    generic_columns = all(
        isinstance(column, int)
        or str(column).startswith("Unnamed")
        or str(column).isdigit()
        for column in columns
    )
    if not generic_columns or df.empty:
        return False

    header_values = {
        str(value).strip().lower()
        for value in df.iloc[0].tolist()
        if str(value).strip() and str(value).strip().lower() != "nan"
    }
    header_matches = header_values & _HEADER_TOKENS
    return len(header_matches) >= 2


def _looks_like_html_table(filepath: Path) -> bool:
    sample = filepath.read_bytes()[:1024]
    if b"<table" in sample.lower() or b"<html" in sample.lower():
        return True

    for encoding in ("utf-16", "utf-8", "cp1252"):
        try:
            text = sample.decode(encoding, errors="ignore").lower()
        except Exception:
            continue
        if "<table" in text or "<html" in text:
            return True
    return False


def _extract_race_number(filename: str) -> int | None:
    match = re.search(r"\brace\s*#?\s*(\d+)\b", filename, re.IGNORECASE)
    if not match:
        return None

    race_num = int(match.group(1))
    return race_num if race_num > 0 else None