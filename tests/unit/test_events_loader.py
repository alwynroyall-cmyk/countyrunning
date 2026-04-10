import datetime
from pathlib import Path

from openpyxl import Workbook

from league_scorer.events_loader import load_events


def test_load_events_parses_date_cells(tmp_path: Path):
    path = tmp_path / "WRRL_events.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Championship Events"
    ws.append([
        "RaceRef",
        "EventName",
        "Category",
        "Distance",
        "Location",
        "Organiser",
        "DateType",
        "ScheduledDates",
        "EligibilityWindow",
        "EntryFee",
        "ScoringBasis",
        "Notes",
        "Status",
    ])
    ws.append([
        "1",
        "Spring Series",
        "Seniors",
        "5k",
        "Oxford",
        "Club A",
        "Date",
        datetime.date(2025, 5, 17),
        "Open",
        "Free",
        "League",
        "Notes",
        "Confirmed",
    ])
    wb.save(path)

    schedule = load_events(path)

    assert len(schedule) == 1
    assert schedule.events[0].scheduled_dates == "17 May 2025"
