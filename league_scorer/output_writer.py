"""
Write all output Excel files.

Cumulative (single workbook per run):
    Race N -- Results.xlsx   — sheets: Race Summary | Div 1 | Div 2 | Male | Female | Race N

Per race:
  Race # -- unused clubs.xlsx
"""

import logging
from pathlib import Path
from typing import Dict, List

from .models import (
    RaceIssue,
    RunnerRaceEntry,
    RunnerSeasonRecord,
    TeamRaceResult,
    TeamSeasonRecord,
    UnrecognisedClub,
)

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

log = logging.getLogger(__name__)

from .settings import settings

MAX_RACES = settings.get("MAX_RACES")
_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_AGG_FILL    = PatternFill("solid", fgColor="D9E1F2")   # light blue tint for aggregate columns
_ALT_ROW_FILL = PatternFill("solid", fgColor="EEF3F8")


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

    if ws.title == "Race Summary":
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=0):
            use_alt_fill = row_idx % 2 == 1
            for cell in row:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if use_alt_fill:
                    cell.fill = _ALT_ROW_FILL

    issue_col_indices = [
        col_idx
        for col_idx, col_name in enumerate(df.columns, start=1)
        if str(col_name) in {"Warnings", "Other Issues"}
    ]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.column in issue_col_indices:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Auto-width
    for col_idx, col_name in enumerate(df.columns, start=1):
        values = [str(col_name)] + [
            str(v) for v in df.iloc[:, col_idx - 1] if str(v) != ""
        ]
        max_width = 80 if str(col_name) in {"Warnings", "Other Issues"} else 50
        width = min(max((len(v) for v in values), default=8) + 2, max_width)
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
    race_files: Dict[int, Path],
    all_unrec_clubs: Dict[int, List[UnrecognisedClub]],
    race_issues: Dict[int, List[RaceIssue]],
    filepath: Path,
) -> None:
    """Write Race Summary, Div 1, Div 2, Male, Female, and Race N sheets."""

    def _format_issue(issue: RaceIssue) -> str:
        if issue.source_row is not None:
            return f"Row {issue.source_row}: {issue.message}"
        return issue.message

    def _summarise_issue_bucket(messages: List[str]) -> str:
        if not messages:
            return ""
        return f"Count: {len(messages)}\n" + "\n".join(messages)

    # ── Individual table builder ─────────────────────────────────────────────
    def _ind_df(records: List[RunnerSeasonRecord]) -> pd.DataFrame:
        rows = []
        for rec in sorted(records, key=lambda r: (r.position, r.name)):
            row: dict = {
                "Position": rec.position,
                "Name":     rec.name,
                "Club":     rec.preferred_club,
                "Category": rec.category,
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

    def _race_df(race_num: int, runners: List[RunnerRaceEntry]) -> pd.DataFrame:
        rows = []
        sorted_runners = sorted(runners, key=lambda r: (r.time_seconds, r.name.lower()))
        gender_positions = {"M": 0, "F": 0}

        for overall_position, runner in enumerate(sorted_runners, start=1):
            gender_positions[runner.gender] = gender_positions.get(runner.gender, 0) + 1
            rows.append(
                {
                    "Overall Pos": overall_position,
                    "Gender Pos": gender_positions[runner.gender],
                    "Name": runner.name,
                    "Raw Club": runner.raw_club,
                    "Club": runner.preferred_club or "",
                    "Gender": runner.gender,
                    "Raw Category": runner.raw_category,
                    "Category": runner.normalised_category,
                    "Time": runner.time_str,
                    "Wiltshire Eligible": "Yes" if runner.eligible else "No",
                    "Points": runner.points if runner.points > 0 else "",
                    "Team": runner.team_id,
                    "Warnings": "\n".join(runner.warnings),
                }
            )

        return pd.DataFrame(
            rows,
            columns=[
                "Overall Pos",
                "Gender Pos",
                "Name",
                "Raw Club",
                "Club",
                "Gender",
                "Raw Category",
                "Category",
                "Time",
                "Wiltshire Eligible",
                "Points",
                "Team",
                "Warnings",
            ],
        )

    def _race_summary_df() -> pd.DataFrame:
        rows = []
        for race_num in sorted(race_files):
            issues = race_issues.get(race_num, [])
            warnings = [_format_issue(issue) for issue in issues if issue.kind == "warning"]
            other_issues = [_format_issue(issue) for issue in issues if issue.kind != "warning"]
            runners = all_race_runners.get(race_num, [])
            unrec = all_unrec_clubs.get(race_num, [])
            rows.append(
                {
                    "Race": race_num,
                    "Source File": race_files[race_num].name,
                    "Status": "Processed" if race_num in all_race_runners else "Skipped",
                    "Runner Rows": len(runners),
                    "Unrecognised Clubs": len(unrec),
                    "Warnings": _summarise_issue_bucket(warnings),
                    "Other Issues": _summarise_issue_bucket(other_issues),
                }
            )
        return pd.DataFrame(
            rows,
            columns=[
                "Race",
                "Source File",
                "Status",
                "Runner Rows",
                "Unrecognised Clubs",
                "Warnings",
                "Other Issues",
            ],
        )

    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            for df, sheet in [
                (_race_summary_df(),      "Race Summary"),
                (_club_df(div1_teams),    "Div 1"),
                (_club_df(div2_teams),    "Div 2"),
                (_ind_df(male_records),   "Male"),
                (_ind_df(female_records), "Female"),
            ]:
                df.to_excel(writer, index=False, sheet_name=sheet)
                _style_and_width(writer.sheets[sheet], df)

            for race_num in sorted(all_race_runners):
                race_df = _race_df(race_num, all_race_runners[race_num])
                sheet_name = f"Race {race_num}"
                race_df.to_excel(writer, index=False, sheet_name=sheet_name)
                _style_and_width(writer.sheets[sheet_name], race_df)
        log.debug("Written: %s", filepath.name)
    except Exception as exc:
        log.error("Failed to write results workbook '%s': %s", filepath, exc)


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
