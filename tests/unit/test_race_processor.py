from pathlib import Path

import pandas as pd

from league_scorer.race_processor import process_race_file


def test_process_race_file_skips_blank_name_with_valid_time(tmp_path):
    df = pd.DataFrame(
        {
            "Position": [1, 2],
            "Name": ["Alice", ""],
            "Club": ["Club A", "Club A"],
            "Gender": ["F", "F"],
            "Category": ["Sen", "Sen"],
            "Chip Time": ["00:40:00", "00:45:00"],
        }
    )
    filepath = tmp_path / "race1.xls"
    filepath.write_text(df.to_html(index=False), encoding="utf-8")

    runners, cat_records, unrec_clubs, issues = process_race_file(
        filepath, 1, {"club a": "Club A"}
    )

    assert len(runners) == 1
    assert runners[0].name == "Alice"
    assert any(issue.code == "AUD-ROW-005" for issue in issues)


def test_process_race_file_preserves_unrecognised_club(tmp_path):
    df = pd.DataFrame(
        {
            "Position": [1],
            "Name": ["Bob"],
            "Club": ["Unknown Club"],
            "Gender": ["M"],
            "Category": ["Sen"],
            "Chip Time": ["00:45:00"],
        }
    )
    filepath = tmp_path / "race2.xls"
    filepath.write_text(df.to_html(index=False), encoding="utf-8")

    runners, cat_records, unrec_clubs, issues = process_race_file(
        filepath, 2, {"club a": "Club A"}
    )

    assert len(runners) == 1
    assert not runners[0].eligible
    assert len(unrec_clubs) == 1
    assert unrec_clubs[0].raw_club_name == "Unknown Club"
    assert unrec_clubs[0].occurrences == 1


def test_process_race_file_deduplicates_same_runner(tmp_path):
    df = pd.DataFrame(
        {
            "Position": [1, 2],
            "Name": ["Charlie", "Charlie"],
            "Club": ["Club A", "Club A"],
            "Gender": ["M", "M"],
            "Category": ["Sen", "Sen"],
            "Chip Time": ["00:42:00", "00:39:00"],
        }
    )
    filepath = tmp_path / "race3.xls"
    filepath.write_text(df.to_html(index=False), encoding="utf-8")

    runners, cat_records, unrec_clubs, issues = process_race_file(
        filepath, 3, {"club a": "Club A"}
    )

    assert len(runners) == 1
    assert runners[0].time_str == "00:39:00"
    assert any(issue.code == "AUD-ROW-008" for issue in issues)
