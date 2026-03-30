"""
Write all output Excel files.

Cumulative (single workbook per run):
  Race N -- Results.xlsx   — sheets: Summary | Div 1 | Div 2 | Male | Female

Per race:
  Race # -- categories.xlsx
  Race # -- unused clubs.xlsx
"""

import logging
from pathlib import Path
from typing import Dict, List

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import (
    CategoryRecord,
    RunnerRaceEntry,
    RunnerSeasonRecord,
    TeamRaceResult,
    TeamSeasonRecord,
    UnrecognisedClub,
)

log = logging.getLogger(__name__)

MAX_RACES = 8
_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_AGG_FILL    = PatternFill("solid", fgColor="D9E1F2")   # light blue tint for aggregate columns


# ──────────────────────────────────────────────────────────────── helpers ────

def _style_and_width(ws, df: pd.DataFrame) -> None:
    """Bold blue header + approximate column widths + aggregate column shading."""
    # Style the header row
    for cell in ws[1]:
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    # Light-shade every Aggregate data column so it stands out at a glance
    agg_col_indices = [
        col_idx
        for col_idx, col_name in enumerate(df.columns, start=1)
        if "Aggregate" in str(col_name)
    ]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.column in agg_col_indices:
                cell.fill = _AGG_FILL

    # Auto-width
    for col_idx, col_name in enumerate(df.columns, start=1):
        values = [str(col_name)] + [
            str(v) for v in df.iloc[:, col_idx - 1] if str(v) != ""
        ]
        width = min(max((len(v) for v in values), default=8) + 2, 50)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _write_df(df: pd.DataFrame, filepath: Path, sheet_name: str = "Data") -> None:
    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            _style_and_width(writer.sheets[sheet_name], df)
        log.debug("Written: %s", filepath.name)
    except Exception as exc:
        log.error("Failed to write '%s': %s", filepath, exc)


# ─────────────────────────────────────────────────────── individual table ────

def write_results_workbook(
    highest_race: int,
    male_records: List[RunnerSeasonRecord],
    female_records: List[RunnerSeasonRecord],
    div1_teams: List[TeamSeasonRecord],
    div2_teams: List[TeamSeasonRecord],
    all_race_runners: Dict[int, List[RunnerRaceEntry]],
    unrec_clubs_all: List[UnrecognisedClub],
    filepath: Path,
) -> None:
    """Write Summary, Div 1, Div 2, Male and Female into a single workbook."""

    # ── Summary data ────────────────────────────────────────────────────────
    clubs_scored = {r.preferred_club for r in male_records + female_records}
    unrec_names  = sorted({u.raw_club_name for u in unrec_clubs_all})
    overview_rows = [
        ("Highest Race Number",     highest_race),
        ("Race Files Processed",    len(all_race_runners)),
        ("Runners Scored (Male)",   len(male_records)),
        ("Runners Scored (Female)", len(female_records)),
        ("Clubs Scored",            len(clubs_scored)),
        ("Unidentified Clubs",      ", ".join(unrec_names) if unrec_names else "None"),
    ]
    df_summary = pd.DataFrame(overview_rows, columns=["Item", "Value"])

    # ── Individual table builder ─────────────────────────────────────────────
    def _ind_df(records: List[RunnerSeasonRecord]) -> pd.DataFrame:
        rows = []
        for rec in sorted(records, key=lambda r: (r.position, r.name)):
            last_race = max(rec.race_times) if rec.race_times else None
            row: dict = {
                "Position": rec.position,
                "Name":     rec.name,
                "Club":     rec.preferred_club,
                "Category": rec.category,
                "Time":     rec.race_times[last_race] if last_race else "",
            }
            for n in range(1, MAX_RACES + 1):
                row[f"Race {n} Points"] = rec.race_points.get(n, "")
            row["Total Points"] = rec.total_points
            rows.append(row)
        return pd.DataFrame(rows)

    # ── Club table builder ───────────────────────────────────────────────────
    def _club_df(teams: List[TeamSeasonRecord]) -> pd.DataFrame:
        rows = []
        for team in sorted(teams, key=lambda t: (t.position, t.preferred_club)):
            row: dict = {
                "Position": team.position,
                "Club":     team.display_name,
            }
            for n in range(1, MAX_RACES + 1):
                rr = team.race_results.get(n)
                row[f"Race {n} Men Score"]   = rr.men_score   if rr else ""
                row[f"Race {n} Women Score"] = rr.women_score if rr else ""
                row[f"Race {n} Aggregate"]   = (
                    (rr.men_score or 0) + (rr.women_score or 0) if rr else ""
                )
                row[f"Race {n} Team Points"] = rr.team_points if rr else ""
            row["Total Points"] = team.total_points
            rows.append(row)
        return pd.DataFrame(rows)

    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            for df, sheet in [
                (df_summary,              "Summary"),
                (_club_df(div1_teams),    "Div 1"),
                (_club_df(div2_teams),    "Div 2"),
                (_ind_df(male_records),   "Male"),
                (_ind_df(female_records), "Female"),
            ]:
                df.to_excel(writer, index=False, sheet_name=sheet)
                _style_and_width(writer.sheets[sheet], df)
        log.debug("Written: %s", filepath.name)
    except Exception as exc:
        log.error("Failed to write results workbook '%s': %s", filepath, exc)


# ─────────────────────────────────────────────────────── exception reports ───

def write_category_report(
    cat_records: List[CategoryRecord],
    filepath: Path,
) -> None:
    """Spec 13.1 Race # -- categories.xlsx"""
    rows = [
        {
            "Raw Category":        r.raw_category,
            "Normalised Category": r.normalised_category,
            "Count":               r.count,
            "Notes":               r.notes,
        }
        for r in sorted(cat_records, key=lambda r: r.raw_category)
    ]
    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["Raw Category", "Normalised Category", "Count", "Notes"]
    )
    _write_df(df, filepath, sheet_name="Categories")


def write_unrecognised_clubs(
    unrec_clubs: List[UnrecognisedClub],
    filepath: Path,
) -> None:
    """Spec 13.2 Race # -- unused clubs.xlsx"""
    rows = [
        {
            "Raw Club Name": u.raw_club_name,
            "Occurrences":   u.occurrences,
            "Action":        "Excluded",
        }
        for u in sorted(unrec_clubs, key=lambda u: u.raw_club_name)
    ]
    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["Raw Club Name", "Occurrences", "Action"]
    )
    _write_df(df, filepath, sheet_name="Unused Clubs")
