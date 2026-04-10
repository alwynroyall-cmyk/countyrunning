from pathlib import Path

import pytest
from openpyxl import Workbook

from league_scorer.club_loader import load_clubs
from league_scorer.exceptions import FatalError


def test_load_clubs_reads_mapping_and_club_info(tmp_path: Path):
    path = tmp_path / "clubs.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Club", "Preferred name", "Team A", "Team B"])
    ws.append(["Guest Runners", "Guest Runners", "2", "1"])
    ws.append(["Club A", "Club A", 1, 2])
    wb.save(path)

    raw_to_preferred, club_info = load_clubs(path)

    assert raw_to_preferred["guest runners"] == "Guest Runners"
    assert raw_to_preferred["club a"] == "Club A"
    assert club_info["Club A"].div_a == 1
    assert club_info["Club A"].div_b == 2
    assert club_info["Guest Runners"].preferred_name == "Guest Runners"


def test_load_clubs_skips_blank_rows_and_requires_columns(tmp_path: Path):
    path = tmp_path / "clubs.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Club", "Preferred name", "Team A", "Team B"])
    ws.append(["", "Club A", 1, 2])
    ws.append(["Club B", "", 2, 1])
    wb.save(path)

    with pytest.raises(FatalError, match="contains no valid club entries"):
        load_clubs(path)


def test_load_clubs_raises_on_missing_file(tmp_path: Path):
    path = tmp_path / "missing_clubs.xlsx"
    with pytest.raises(FatalError, match="clubs.xlsx not found"):
        load_clubs(path)
