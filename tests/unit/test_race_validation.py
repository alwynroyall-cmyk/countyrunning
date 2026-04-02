from pathlib import Path

import pandas as pd
import pytest

from league_scorer.exceptions import RaceProcessingError
from league_scorer.race_validation import validate_race_schema


def test_validate_race_schema_detects_required_columns_and_time():
    df = pd.DataFrame(
        {
            "Position": [1],
            "Name": ["A Runner"],
            "Club": ["Club A"],
            "Gender": ["M"],
            "Category": ["Sen"],
            "Chip Time": ["00:40:00"],
        }
    )

    result = validate_race_schema(df, Path("Race 1.xlsx"))
    assert result.column_map["Name"] == "Name"
    assert result.time_column == "Chip Time"


def test_validate_race_schema_raises_on_missing_required():
    df = pd.DataFrame({"Name": ["Runner"], "Chip Time": ["00:40:00"]})

    with pytest.raises(RaceProcessingError):
        validate_race_schema(df, Path("Race 1.xlsx"))


def test_validate_race_schema_flags_blank_name_rows_warning():
    df = pd.DataFrame(
        {
            "Position": [1, 2],
            "Name": ["Runner One", ""],
            "Club": ["Club A", "Club A"],
            "Gender": ["M", "F"],
            "Category": ["Sen", "Sen"],
            "Chip Time": ["00:40:00", "00:45:00"],
        }
    )

    result = validate_race_schema(df, Path("Race 1.xlsx"))
    codes = [issue.code for issue in result.issues]
    assert "AUD-RACE-009" in codes
