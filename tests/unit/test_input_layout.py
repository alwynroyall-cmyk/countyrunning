from pathlib import Path

from league_scorer.input_layout import (
    InputPaths,
    InputSortResult,
    build_input_paths,
    ensure_input_subdirs,
    sort_existing_input_files,
)


def test_build_input_paths_returns_expected_paths(tmp_path: Path) -> None:
    input_dir = tmp_path / "inputs"
    paths = build_input_paths(input_dir)

    assert isinstance(paths, InputPaths)
    assert paths.input_dir == input_dir
    assert paths.raw_data_dir == input_dir / "raw_data"
    assert paths.series_dir == input_dir / "series"
    assert paths.control_dir == input_dir / "control"
    assert paths.audited_dir == input_dir / "audited"
    assert paths.raw_data_archive_dir == input_dir / "raw_data_archive"
    assert paths.clubs_path == input_dir / "control" / "clubs.xlsx"
    assert paths.name_corrections_path == input_dir / "control" / "name_corrections.xlsx"


def test_ensure_input_subdirs_creates_required_directories(tmp_path: Path) -> None:
    input_dir = tmp_path / "inputs"
    paths = ensure_input_subdirs(input_dir)

    assert input_dir.exists()
    assert paths.raw_data_dir.exists()
    assert paths.series_dir.exists()
    assert paths.control_dir.exists()
    assert paths.audited_dir.exists()
    assert paths.raw_data_archive_dir.exists()


def test_sort_existing_input_files_moves_legacy_files_to_structured_subdirs(tmp_path: Path) -> None:
    input_dir = tmp_path / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)

    raw_file = input_dir / "Race #1.xls"
    audited_file = input_dir / "Club Series 1 (audited).xlsx"
    series_file = input_dir / "Series #2.xlsx"
    control_file = input_dir / "clubs.xlsx"
    skip_file = input_dir / "readme.txt"

    raw_file.write_text("raw", encoding="utf-8")
    audited_file.write_text("audited", encoding="utf-8")
    series_file.write_text("series", encoding="utf-8")
    control_file.write_text("control", encoding="utf-8")
    skip_file.write_text("skip", encoding="utf-8")

    result = sort_existing_input_files(input_dir)

    assert isinstance(result, InputSortResult)
    assert result.moved_count == 4
    assert result.skipped_count == 1
    assert raw_file.exists() is False
    assert (input_dir / "raw_data" / raw_file.name).exists()
    assert (input_dir / "audited" / audited_file.name).exists()
    assert (input_dir / "series" / series_file.name).exists()
    assert (input_dir / "control" / control_file.name).exists()
    assert (input_dir / skip_file.name).exists()
    assert result.moved_files[raw_file.name] == str(Path("raw_data") / raw_file.name)
    assert result.moved_files[control_file.name] == str(Path("control") / "clubs.xlsx")
