"""Helpers for the structured season input folder layout."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil

CONTROL_FILENAMES = {
    "clubs.xlsx",
    "name_corrections.xlsx",
    "wrrl_events.xlsx",
}

RACE_FILE_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".csv"}


@dataclass(frozen=True)
class InputPaths:
    input_dir: Path
    raw_data_dir: Path
    series_dir: Path
    control_dir: Path
    audited_dir: Path
    raw_data_archive_dir: Path

    @property
    def clubs_path(self) -> Path:
        return self.control_dir / "clubs.xlsx"

    @property
    def name_corrections_path(self) -> Path:
        return self.control_dir / "name_corrections.xlsx"


@dataclass(frozen=True)
class InputSortResult:
    moved_count: int
    skipped_count: int
    moved_files: dict[str, str]


def build_input_paths(input_dir: Path) -> InputPaths:
    input_dir = Path(input_dir)
    return InputPaths(
        input_dir=input_dir,
        raw_data_dir=input_dir / "raw_data",
        series_dir=input_dir / "series",
        control_dir=input_dir / "control",
        audited_dir=input_dir / "audited",
        raw_data_archive_dir=input_dir / "raw_data_archive",
    )


def ensure_input_subdirs(input_dir: Path) -> InputPaths:
    paths = build_input_paths(input_dir)
    paths.input_dir.mkdir(parents=True, exist_ok=True)
    paths.raw_data_dir.mkdir(parents=True, exist_ok=True)
    paths.series_dir.mkdir(parents=True, exist_ok=True)
    paths.control_dir.mkdir(parents=True, exist_ok=True)
    paths.audited_dir.mkdir(parents=True, exist_ok=True)
    paths.raw_data_archive_dir.mkdir(parents=True, exist_ok=True)
    return paths


def sort_existing_input_files(input_dir: Path) -> InputSortResult:
    """Move legacy flat input files into the structured subfolders."""
    paths = ensure_input_subdirs(input_dir)
    moved_files: dict[str, str] = {}
    moved_count = 0
    skipped_count = 0

    for item in sorted(paths.input_dir.iterdir()):
        if item.is_dir():
            continue

        destination_dir = _destination_for_file(item, paths)
        if destination_dir is None:
            skipped_count += 1
            continue

        destination = destination_dir / item.name
        if destination.exists():
            skipped_count += 1
            continue

        shutil.move(str(item), str(destination))
        moved_count += 1
        moved_files[item.name] = str(destination.relative_to(paths.input_dir))

    return InputSortResult(
        moved_count=moved_count,
        skipped_count=skipped_count,
        moved_files=moved_files,
    )


def _destination_for_file(file_path: Path, paths: InputPaths) -> Path | None:
    lower_name = file_path.name.lower()
    suffix = file_path.suffix.lower()

    if lower_name in CONTROL_FILENAMES:
        return paths.control_dir

    if suffix not in RACE_FILE_SUFFIXES:
        return None

    stem = file_path.stem.lower()
    if "(audited)" in stem:
        return paths.audited_dir

    if "series" in stem and _looks_like_round_file(stem):
        return paths.series_dir

    if _looks_like_race_file(stem) or lower_name == "race roster import history.csv":
        return paths.raw_data_dir

    return None


def _looks_like_race_file(stem: str) -> bool:
    return bool(re.search(r"\brace\s*#?\s*\d+\b", stem, re.IGNORECASE))


def _looks_like_round_file(stem: str) -> bool:
    return bool(re.search(r"\b(round|series)\s*#?\s*\d+\b", stem, re.IGNORECASE))
