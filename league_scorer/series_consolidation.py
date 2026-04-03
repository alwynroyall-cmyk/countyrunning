"""Helpers for consolidating multi-file race series into one input workbook."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

import pandas as pd

from .source_loader import load_race_dataframe

_SERIES_PATTERN = re.compile(
    r"^race\s*#?\s*(\d+)\s*[-–]?\s*(.+?)\s+series\s*#\s*(\d+)\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class SeriesFileInfo:
    path: Path
    race_number: int
    series_name: str
    series_index: int


@dataclass(frozen=True)
class SeriesConsolidationResult:
    consolidated_path: Path
    round_number: int
    removed_previous: List[Path]
    club_warnings: List[str]


def consolidate_series_files(
    filepaths: Iterable[Path],
    *,
    series_dir: Path,
    raw_data_dir: Path,
) -> SeriesConsolidationResult:
    selected_paths = [Path(filepath) for filepath in filepaths]
    if len(selected_paths) < 2:
        raise ValueError("Select at least two series files to consolidate.")

    parsed_files = [_parse_series_file(path) for path in selected_paths]
    _validate_series_selection(parsed_files, series_dir)
    parsed_files.sort(key=lambda item: item.series_index)

    consolidated_df, club_warnings = _build_consolidated_dataframe(parsed_files)
    series_name = parsed_files[0].series_name
    race_number = parsed_files[0].race_number
    round_number = max(item.series_index for item in parsed_files)

    raw_data_dir.mkdir(parents=True, exist_ok=True)
    consolidated_path = raw_data_dir / f"Race #{race_number} {series_name} Round {round_number}.xlsx"

    removed_previous: List[Path] = []
    for existing in sorted(raw_data_dir.glob(f"Race #{race_number} {series_name} Round *.xlsx")):
        if existing.resolve() == consolidated_path.resolve():
            continue
        existing.unlink()
        removed_previous.append(existing)

    if consolidated_path.exists():
        consolidated_path.unlink()
    consolidated_df.to_excel(consolidated_path, index=False)

    return SeriesConsolidationResult(
        consolidated_path=consolidated_path,
        round_number=round_number,
        removed_previous=removed_previous,
        club_warnings=club_warnings,
    )


def _parse_series_file(path: Path) -> SeriesFileInfo:
    match = _SERIES_PATTERN.match(path.stem.strip())
    if not match:
        raise ValueError(
            f"'{path.name}' is not a recognised series file. Expected a name like 'Race #3 Westbury 5k Series #1'."
        )

    race_number = int(match.group(1))
    series_name = match.group(2).strip(" -")
    series_index = int(match.group(3))
    return SeriesFileInfo(
        path=path,
        race_number=race_number,
        series_name=series_name,
        series_index=series_index,
    )


def _validate_series_selection(parsed_files: List[SeriesFileInfo], input_dir: Path) -> None:
    first = parsed_files[0]
    for info in parsed_files:
        try:
            info.path.resolve().relative_to(input_dir.resolve())
        except ValueError as exc:
            raise ValueError(
                f"'{info.path.name}' is outside the active input folder and cannot be consolidated."
            ) from exc

        if info.race_number != first.race_number:
            raise ValueError("Selected files must all belong to the same race number.")
        if info.series_name.lower() != first.series_name.lower():
            raise ValueError("Selected files must all belong to the same series name.")

    seen_indexes: set[int] = set()
    for info in parsed_files:
        if info.series_index in seen_indexes:
            raise ValueError("Duplicate series leg numbers were selected for consolidation.")
        seen_indexes.add(info.series_index)


def _build_consolidated_dataframe(parsed_files: List[SeriesFileInfo]) -> tuple[pd.DataFrame, List[str]]:
    frames: List[pd.DataFrame] = []
    combined_columns: List[str] = []
    club_warnings: List[str] = []
    runner_clubs: dict[str, set[str]] = {}

    for info in parsed_files:
        frame = load_race_dataframe(info.path).copy()
        frame.columns = [str(column).strip() for column in frame.columns]
        frame = frame.loc[:, ~frame.columns.duplicated()].copy()
        _capture_runner_club_inconsistencies(frame, runner_clubs)
        frames.append(frame)
        for column in frame.columns:
            if column not in combined_columns:
                combined_columns.append(column)

    normalised_frames = [frame.reindex(columns=combined_columns) for frame in frames]
    for runner_key, clubs in sorted(runner_clubs.items()):
        clean = sorted(club for club in clubs if club)
        if len(clean) > 1:
            club_warnings.append(f"{runner_key}: {', '.join(clean)}")
    return pd.concat(normalised_frames, ignore_index=True), club_warnings


def _capture_runner_club_inconsistencies(
    frame: pd.DataFrame,
    runner_clubs: dict[str, set[str]],
) -> None:
    lower_columns = {str(col).strip().lower(): col for col in frame.columns}
    name_col = next((lower_columns[key] for key in lower_columns if "name" in key), None)
    club_col = next((lower_columns[key] for key in lower_columns if "club" in key), None)
    if name_col is None or club_col is None:
        return

    for _, row in frame.iterrows():
        name = str(row.get(name_col, "") or "").strip()
        club = str(row.get(club_col, "") or "").strip()
        if not name:
            continue
        key = name.lower()
        runner_clubs.setdefault(key, set())
        if club:
            runner_clubs[key].add(club)