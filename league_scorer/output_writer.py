"""
Write all output Excel files.

Cumulative (single workbook per run):
    Race N -- Results.xlsx   — sheets: Race Summary | Div 1 | Div 2 | Male | Female | Race N

Per race:
  Race # -- unused clubs.xlsx
"""

import logging
from collections import defaultdict
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
        msg = (
            f"Failed to write output file '{filepath}'. "
            "Please close the workbook if it is open and try again. "
            f"Original error: {exc}"
        )
        log.error(msg)
        raise RuntimeError(msg) from exc


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
    """Write Race Summary, Div 1, Div 2, Male, Female, Category Review, and Race N sheets."""
    max_races = settings.get("MAX_RACES")

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
            for n in range(1, max_races + 1):
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
            for n in range(1, max_races + 1):
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

    def _is_unknown_raw_category(raw_category: str) -> bool:
        compact = "".join(ch for ch in str(raw_category).strip().lower() if ch.isalnum())
        return compact in {"", "unknown", "qry", "na", "nacat", "uncategorised", "uncategorized"}

    def _category_review_df() -> pd.DataFrame:
        grouped: Dict[tuple, List[RunnerRaceEntry]] = defaultdict(list)
        for race_num, runners in all_race_runners.items():
            for runner in runners:
                key = (
                    runner.name.strip().lower(),
                    (runner.preferred_club or runner.raw_club or "").strip().lower(),
                    runner.gender,
                )
                grouped[key].append(runner)

        rows = []
        for grouped_rows in grouped.values():
            unknown_rows = [r for r in grouped_rows if _is_unknown_raw_category(r.raw_category)]
            if not unknown_rows:
                continue

            sample = sorted(grouped_rows, key=lambda r: r.race_number)[0]
            known_rows = [r for r in grouped_rows if not _is_unknown_raw_category(r.raw_category)]

            unknown_races = ", ".join(str(r.race_number) for r in sorted(unknown_rows, key=lambda r: r.race_number))
            unknown_values = ", ".join(
                sorted({str(r.raw_category).strip() or "(blank)" for r in unknown_rows})
            )

            if known_rows:
                evidence_parts = []
                for race_row in sorted(known_rows, key=lambda r: r.race_number):
                    evidence_parts.append(
                        f"R{race_row.race_number}: raw '{race_row.raw_category}' -> {race_row.normalised_category}"
                    )
                suggested = sorted({r.normalised_category for r in known_rows if r.normalised_category})
                suggested_text = ", ".join(suggested)
                status = "Review Suggested"
                next_step = "Check race sheets listed in Evidence and confirm category for unknown races."
            else:
                evidence_parts = []
                suggested_text = ""
                status = "No Evidence"
                next_step = "No known category found in other races yet; leave unresolved until more results exist."

            rows.append(
                {
                    "Name": sample.name,
                    "Club": sample.preferred_club or sample.raw_club,
                    "Gender": sample.gender,
                    "Unknown Races": unknown_races,
                    "Unknown Source Categories": unknown_values,
                    "Suggested Category": suggested_text,
                    "Evidence": "\n".join(evidence_parts),
                    "Status": status,
                    "Next Step": next_step,
                }
            )

        return pd.DataFrame(
            rows,
            columns=[
                "Name",
                "Club",
                "Gender",
                "Unknown Races",
                "Unknown Source Categories",
                "Suggested Category",
                "Evidence",
                "Status",
                "Next Step",
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
                (_category_review_df(),   "Category Review"),
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
        msg = (
            f"Failed to write results workbook '{filepath}'. "
            "Please close the workbook if it is open and try again. "
            f"Original error: {exc}"
        )
        log.error(msg)
        raise RuntimeError(msg) from exc


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
