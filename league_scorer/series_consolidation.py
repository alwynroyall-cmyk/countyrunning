"""Helpers for consolidating multi-file race series into one input workbook."""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

import pandas as pd

from .source_loader import load_race_dataframe

_SERIES_PATTERN = re.compile(
    r"^race\s*#?\s*(\d+)\s*[-–]?\s*(.+?)\s+series\s*#\s*(\d+)\s*$",
    re.IGNORECASE,
)
_INVALID_PATH_CHARS = re.compile(r'[<>:"/\\|?*]')


@dataclass(frozen=True)
class SeriesFileInfo:
    path: Path
    race_number: int
    series_name: str
    series_index: int


@dataclass(frozen=True)
class SeriesConsolidationResult:
    consolidated_path: Path
    archive_dir: Path
    moved_files: List[Path]


def consolidate_series_files(filepaths: Iterable[Path], input_dir: Path) -> SeriesConsolidationResult:
    selected_paths = [Path(filepath) for filepath in filepaths]
    if len(selected_paths) < 2:
        raise ValueError("Select at least two series files to consolidate.")

    parsed_files = [_parse_series_file(path) for path in selected_paths]
    _validate_series_selection(parsed_files, input_dir)
    parsed_files.sort(key=lambda item: item.series_index)

    consolidated_df = _build_consolidated_dataframe(parsed_files)
    series_name = parsed_files[0].series_name
    race_number = parsed_files[0].race_number

    consolidated_path = input_dir / f"Race #{race_number} {series_name} Consolidated.xlsx"
    archive_dir = input_dir / f"{_safe_path_name(series_name)} Series"
    archive_dir.mkdir(parents=True, exist_ok=True)

    if consolidated_path.exists():
        consolidated_path.unlink()
    consolidated_df.to_excel(consolidated_path, index=False)

    moved_files: List[Path] = []
    for info in parsed_files:
        destination = archive_dir / info.path.name
        if destination.exists() and destination.resolve() != info.path.resolve():
            raise FileExistsError(
                f"Cannot move '{info.path.name}' because '{destination.name}' already exists in '{archive_dir.name}'."
            )
        if info.path.resolve() != destination.resolve():
            shutil.move(str(info.path), str(destination))
        moved_files.append(destination)

    return SeriesConsolidationResult(
        consolidated_path=consolidated_path,
        archive_dir=archive_dir,
        moved_files=moved_files,
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


def _build_consolidated_dataframe(parsed_files: List[SeriesFileInfo]) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    combined_columns: List[str] = []

    for info in parsed_files:
        frame = load_race_dataframe(info.path).copy()
        frame.columns = [str(column).strip() for column in frame.columns]
        frame = frame.loc[:, ~frame.columns.duplicated()].copy()
        frames.append(frame)
        for column in frame.columns:
            if column not in combined_columns:
                combined_columns.append(column)

    normalised_frames = [frame.reindex(columns=combined_columns) for frame in frames]
    return pd.concat(normalised_frames, ignore_index=True)


def _safe_path_name(value: str) -> str:
    cleaned = _INVALID_PATH_CHARS.sub(" ", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "Series"