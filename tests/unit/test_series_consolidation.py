from pathlib import Path

import pandas as pd

from league_scorer.series_consolidation import (
    SeriesFileInfo,
    _parse_series_file,
    _validate_series_selection,
    _build_consolidated_dataframe,
)


def test_parse_series_file_parses_valid_name():
    info = _parse_series_file(Path("Race #3 Westbury 5k Series #2.xlsx"))
    assert info.race_number == 3
    assert info.series_name == "Westbury 5k"
    assert info.series_index == 2


def test_parse_series_file_rejects_invalid_name():
    try:
        _parse_series_file(Path("Race 3 Westbury.xlsx"))
        assert False, "Expected ValueError"
    except ValueError as exc:
        assert "not a recognised series file" in str(exc)


def test_validate_series_selection_rejects_outside_input_dir(tmp_path: Path):
    parsed = [
        SeriesFileInfo(Path(tmp_path / "outside" / "Race #1 Test Series #1.xlsx"), 1, "Test", 1)
    ]
    input_dir = tmp_path / "inputs"
    input_dir.mkdir(parents=True)

    try:
        _validate_series_selection(parsed, input_dir)
        assert False, "Expected ValueError"
    except ValueError as exc:
        assert "outside the active input folder" in str(exc)


def test_build_consolidated_dataframe_merges_columns_and_warns_on_club_conflicts(monkeypatch):
    file1 = Path("Race #1 Test Series #1.xlsx")
    file2 = Path("Race #1 Test Series #2.xlsx")

    def fake_load(path):
        if path == file1:
            return pd.DataFrame([{"Name": "Alice", "Club": "Club A", "Time": "00:40:00"}])
        return pd.DataFrame([{"Name": "Alice", "Club": "Club B", "Time": "00:41:00", "Extra": "Data"}])

    monkeypatch.setattr("league_scorer.series_consolidation.load_race_dataframe", fake_load)
    parsed = [
        SeriesFileInfo(file1, 1, "Test", 1),
        SeriesFileInfo(file2, 1, "Test", 2),
    ]

    consolidated, warnings = _build_consolidated_dataframe(parsed)
    assert "Extra" in consolidated.columns
    assert len(consolidated) == 2
    assert warnings == ["alice: Club A, Club B"]
