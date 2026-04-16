import logging
import re
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

from .audit_writer import write_audit_workbook
from .club_loader import load_clubs
from .common_files import race_discovery_exclusions
from .exceptions import FatalError, RaceProcessingError
from .input_layout import build_input_paths, ensure_input_subdirs, sort_existing_input_files
from .output_layout import ensure_output_subdirs, sort_existing_output_files
from .models import RaceIssue, RunnerRaceEntry, UnrecognisedClub
from .race_processor import process_race_file
from .source_loader import discover_race_files, load_race_dataframe

log = logging.getLogger(__name__)


class LeagueAuditor:
    def __init__(self, input_dir: Path, output_dir: Path, year: int) -> None:
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.year = int(year)
        self.input_paths = build_input_paths(input_dir)

        self.raw_to_preferred: Dict[str, str] = {}
        self.preferred_clubs: List[str] = []

        self.all_race_runners: Dict[int, List[RunnerRaceEntry]] = {}
        self.all_unrec_clubs: Dict[int, List[UnrecognisedClub]] = {}
        self.all_race_issues: Dict[int, List[RaceIssue]] = {}
        self.selected_race_files: Dict[int, Path] = {}
        self.latest_actionable_issue_count: int = 0

    def run(self, race_files: "Dict[int, Path] | None" = None) -> Path:
        self._validate_paths()
        self._load_clubs()

        if race_files is None:
            race_files = self._discover_races()

        self.selected_race_files = dict(race_files)
        if not race_files:
            raise FatalError(f"No valid race files found in '{self.input_dir}'.")

        for race_num in sorted(race_files):
            self._process_race(race_num, race_files[race_num])

        output_paths = ensure_output_subdirs(self.output_dir)
        workbook_path = output_paths.audit_workbooks_dir / self._build_filename()
        sheets = self._build_sheets()
        write_audit_workbook(sheets, workbook_path)
        return workbook_path

    def _validate_paths(self) -> None:
        if self.input_dir.name != "inputs":
            raise FatalError(f"Input folder must end with 'inputs', got: {self.input_dir}")
        if self.output_dir.name != "outputs":
            raise FatalError(f"Output folder must end with 'outputs', got: {self.output_dir}")
        if self.input_dir.parent.name != str(self.year):
            raise FatalError(
                f"Configured season year does not match input directory layout: {self.input_dir}"
            )
        if self.output_dir.parent.name != str(self.year):
            raise FatalError(
                f"Configured season year does not match output directory layout: {self.output_dir}"
            )
        ensure_input_subdirs(self.input_dir)
        sort_existing_input_files(self.input_dir)
        ensure_output_subdirs(self.output_dir)
        sort_existing_output_files(self.output_dir)

    def _load_clubs(self) -> None:
        self.raw_to_preferred, club_info = load_clubs(self.input_paths.control_dir / "clubs.xlsx")
        self.preferred_clubs = sorted(club_info)

    def _discover_races(self) -> Dict[int, Path]:
        return discover_race_files(self.input_paths.audited_dir, excluded_names=race_discovery_exclusions())

    def _process_race(self, race_num: int, filepath: Path) -> None:
        try:
            runners, _, unrec, issue_notes = process_race_file(filepath, race_num, self.raw_to_preferred)
        except RaceProcessingError as exc:
            self.all_race_runners[race_num] = []
            self.all_unrec_clubs[race_num] = []
            self.all_race_issues[race_num] = [
                RaceIssue("other", f"Race skipped: {exc}", code="AUD-RACE-005")
            ]
            return

        self.all_race_runners[race_num] = runners
        self.all_unrec_clubs[race_num] = unrec
        self.all_race_issues[race_num] = list(issue_notes)

    def _build_filename(self) -> str:
        race_numbers = sorted(self.selected_race_files)
        if len(race_numbers) == 1:
            return f"Race {race_numbers[0]} - Audit.xlsx"
        return "Season Audit.xlsx"

    def _build_sheets(self) -> Dict[str, pd.DataFrame]:
        race_meta = self._build_race_metadata()
        row_df = self._build_row_audit_df(race_meta)
        runner_df = self._build_runner_audit_df(race_meta)
        club_df, unrec_df = self._build_club_audit_dfs()
        actionable_df = self._build_actionable_issues_df(row_df, runner_df, club_df)
        self.latest_actionable_issue_count = len(actionable_df)
        race_df = self._build_race_summary_df(race_meta, row_df, runner_df, club_df)
        return {
            "Race Audit Summary": race_df,
            "Actionable Issues": actionable_df,
            "Row Audit": row_df,
            "Runner Audit": runner_df,
            "Club Audit": club_df,
            "Unrecognised Club Summary": unrec_df,
        }

    def _build_actionable_issues_df(
        self,
        row_df: pd.DataFrame,
        runner_df: pd.DataFrame,
        club_df: pd.DataFrame,
    ) -> pd.DataFrame:
        """Return a compact list of high-signal issues that need manual correction."""
        row_codes = {"AUD-ROW-001", "AUD-ROW-002", "AUD-ROW-005", "AUD-ROW-010"}
        runner_codes = {"AUD-RUNNER-007", "AUD-RUNNER-008"}
        club_codes = {"AUD-CLUB-002", "AUD-CLUB-003"}

        rows: List[dict] = []

        if not row_df.empty:
            subset = row_df[row_df["Issue Code"].astype(str).isin(row_codes)]
            for _, item in subset.iterrows():
                rows.append(
                    {
                        "Type": "Row",
                        "Severity": item.get("Severity", ""),
                        "Issue Code": item.get("Issue Code", ""),
                        "Race": item.get("Race", ""),
                        "Source Row": item.get("Source Row", ""),
                        "Key": "",
                        "Name": item.get("Runner Name", ""),
                        "Club": item.get("Club", ""),
                        "Message": item.get("Message", ""),
                        "Next Step": item.get("Next Step", ""),
                    }
                )

        if not runner_df.empty:
            subset = runner_df[runner_df["Issue Code"].astype(str).isin(runner_codes)]
            for _, item in subset.iterrows():
                rows.append(
                    {
                        "Type": "Runner",
                        "Severity": item.get("Severity", ""),
                        "Issue Code": item.get("Issue Code", ""),
                        "Race": "",
                        "Source Row": "",
                        "Key": item.get("Runner Key", ""),
                        "Name": item.get("Display Name", ""),
                        "Club": item.get("Clubs Seen", ""),
                        "Message": item.get("Message", ""),
                        "Next Step": item.get("Next Step", ""),
                    }
                )

        if not club_df.empty:
            subset = club_df[club_df["Issue Code"].astype(str).isin(club_codes)]
            for _, item in subset.iterrows():
                rows.append(
                    {
                        "Type": "Club",
                        "Severity": item.get("Severity", ""),
                        "Issue Code": item.get("Issue Code", ""),
                        "Race": "",
                        "Source Row": "",
                        "Key": item.get("Raw Club", ""),
                        "Name": "",
                        "Club": item.get("Preferred Club", ""),
                        "Message": item.get("Message", ""),
                        "Next Step": item.get("Next Step", ""),
                    }
                )

        out = pd.DataFrame(
            rows,
            columns=[
                "Type",
                "Severity",
                "Issue Code",
                "Race",
                "Source Row",
                "Key",
                "Name",
                "Club",
                "Message",
                "Next Step",
            ],
        )
        if not out.empty:
            out = out.sort_values(by=["Type", "Issue Code", "Race", "Source Row", "Name"], ignore_index=True)
        return out

    def _build_race_metadata(self) -> Dict[int, dict]:
        metadata: Dict[int, dict] = {}
        for race_num, path in sorted(self.selected_race_files.items()):
            runners = self.all_race_runners.get(race_num, [])
            scheme, evidence = _classify_race_scheme(runners)
            metadata[race_num] = {
                "file": path,
                "scheme": scheme,
                "scheme_evidence": evidence,
            }
        return metadata

    def _build_row_audit_df(self, race_meta: Dict[int, dict]) -> pd.DataFrame:
        rows: List[dict] = []
        for race_num, meta in sorted(race_meta.items()):
            race_file = meta["file"].name
            runners = self.all_race_runners.get(race_num, [])
            issues = self.all_race_issues.get(race_num, [])
            runner_by_row = {runner.source_row: runner for runner in runners}
            is_series = _is_series_race(meta["file"])

            for issue in issues:
                if not issue.code.startswith("AUD-ROW"):
                    continue
                # Dedupe trace notes are expected and numerous in consolidated
                # series files; suppress them from the report there.
                if issue.code == "AUD-ROW-008" and is_series:
                    continue
                runner = runner_by_row.get(issue.source_row or -1)
                rows.append(
                    _build_row_entry(
                        race_num=race_num,
                        race_file=race_file,
                        issue=issue,
                        runner=runner,
                        is_series=is_series,
                    )
                )

            if meta["scheme"] == "EA 5-Year":
                for runner in runners:
                    if not runner.eligible:
                        continue
                    if runner.gender != "F":
                        continue
                    derived_category = _derived_audit_category(runner, meta["scheme"])
                    if derived_category != runner.normalised_category and derived_category:
                        issue = RaceIssue(
                            "warning",
                            f"Category derived from '{runner.raw_category}' to '{derived_category}'",
                            source_row=runner.source_row,
                            code="AUD-ROW-004",
                            runner_name=runner.name,
                            raw_club=runner.raw_club,
                            gender=runner.gender,
                            raw_category=runner.raw_category,
                            time_str=runner.time_str,
                        )
                        rows.append(
                            _build_row_entry(
                                race_num=race_num,
                                race_file=race_file,
                                issue=issue,
                                runner=runner,
                            )
                        )

            for runner in runners:
                if not runner.eligible:
                    continue
                ea_candidate = _build_ea_candidate_row(runner, race_num, race_file)
                if ea_candidate is None:
                    continue
                issue = RaceIssue(
                    "warning",
                    ea_candidate["Reason"],
                    source_row=runner.source_row,
                    code="AUD-ROW-011",
                    runner_name=runner.name,
                    raw_club=runner.raw_club,
                    gender=runner.gender,
                    raw_category=runner.raw_category,
                    time_str=runner.time_str,
                )
                rows.append(
                    _build_row_entry(
                        race_num=race_num,
                        race_file=race_file,
                        issue=issue,
                        runner=runner,
                    )
                )

        return pd.DataFrame(rows, columns=_ROW_AUDIT_COLUMNS)

    def _build_runner_audit_df(self, race_meta: Dict[int, dict]) -> pd.DataFrame:
        rows: List[dict] = []
        all_runners = [runner for runners in self.all_race_runners.values() for runner in runners]

        by_name: Dict[str, List[RunnerRaceEntry]] = defaultdict(list)
        by_name_club: Dict[Tuple[str, str], List[RunnerRaceEntry]] = defaultdict(list)
        by_identity: Dict[Tuple[str, str, str], List[RunnerRaceEntry]] = defaultdict(list)

        for runner in all_runners:
            if not runner.eligible:
                continue
            name_key = runner.name.strip().lower()
            club_key = (runner.preferred_club or runner.raw_club or "").strip().lower()
            by_name[name_key].append(runner)
            by_name_club[(name_key, club_key)].append(runner)
            if club_key:
                by_identity[(name_key, club_key, runner.gender.strip().upper())].append(runner)

        for name_key, grouped in sorted(by_name.items()):
            eligible_clubs = sorted({r.preferred_club for r in grouped if r.preferred_club})
            if len(eligible_clubs) > 1:
                sample = grouped[0]
                rows.append(
                    _build_runner_entry(
                        code="AUD-RUNNER-007",
                        severity="warning",
                        runner_key=f"{name_key}|collision",
                        display_name=sample.name,
                        clubs_seen=eligible_clubs,
                        sexes_seen=sorted({r.gender for r in grouped if r.gender}),
                        categories_seen=sorted({r.normalised_category for r in grouped if r.normalised_category}),
                        races_seen=sorted({r.race_number for r in grouped}),
                        status="Manual Review",
                        depends_on="Identity Review",
                        message="Exact runner name appears across different clubs; identity is ambiguous.",
                        next_step="Confirm whether the matching names are the same runner or different people.",
                    )
                )

        for (name_key, club_key), grouped in sorted(by_name_club.items()):
            sexes_seen = sorted({r.gender for r in grouped if r.gender})
            if len(sexes_seen) > 1:
                sample = grouped[0]
                rows.append(
                    _build_runner_entry(
                        code="AUD-RUNNER-008",
                        severity="warning",
                        runner_key=f"{name_key}|{club_key}|sex-conflict",
                        display_name=sample.name,
                        clubs_seen=sorted({_display_club(r) for r in grouped}),
                        sexes_seen=sexes_seen,
                        categories_seen=sorted({r.normalised_category for r in grouped if r.normalised_category}),
                        races_seen=sorted({r.race_number for r in grouped}),
                        status="Dependent",
                        depends_on="Identity Review",
                        message="Candidate same-runner records conflict on sex; automatic identity escalation has been stopped.",
                        next_step="Resolve identity first, then review downstream runner inconsistencies.",
                    )
                )

        for identity_key, grouped in sorted(by_identity.items()):
            sample = grouped[0]
            categories_by_race = {
                r.race_number: _derived_audit_category(r, race_meta.get(r.race_number, {}).get("scheme", "League Bands"))
                for r in sorted(grouped, key=lambda item: item.race_number)
            }
            categories_seen = list(dict.fromkeys(categories_by_race.values()))
            external_races = [race for race in categories_by_race if race_meta.get(race, {}).get("scheme") == "External Data Check"]
            if len(set(categories_seen)) > 1:
                baseline_race = min(categories_by_race)
                baseline_category = categories_by_race[baseline_race]
                later = ", ".join(
                    f"R{race}={category}"
                    for race, category in sorted(categories_by_race.items())
                    if race != baseline_race
                )
                rows.append(
                    _build_runner_entry(
                        code="AUD-RUNNER-002",
                        severity="warning",
                        runner_key="|".join(identity_key),
                        display_name=sample.name,
                        clubs_seen=[_display_club(sample)],
                        sexes_seen=sorted({r.gender for r in grouped if r.gender}),
                        categories_seen=categories_seen,
                        races_seen=sorted(categories_by_race),
                        status="Dependent" if external_races else "Open",
                        depends_on="External Category Check" if external_races else "None",
                        message=f"Baseline category is {baseline_category} from Race {baseline_race}; later categories: {later}.",
                        next_step="Review category progression and any race-level category scheme issues.",
                    )
                )

        for cluster in _find_name_variant_clusters(by_identity):
            rows.append(cluster)

        return pd.DataFrame(rows, columns=_RUNNER_AUDIT_COLUMNS)

    def _build_club_audit_dfs(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        # Structural issues in clubs.xlsx (duplicate aliases, inconsistent divisions)
        clubs_xlsx = self.input_paths.control_dir / "clubs.xlsx"
        club_rows = _inspect_club_lookup_file(clubs_xlsx)
        club_df = (
            pd.DataFrame(club_rows, columns=_CLUB_AUDIT_COLUMNS)
            if club_rows
            else pd.DataFrame(columns=_CLUB_AUDIT_COLUMNS)
        )

        # Aggregate unrecognised raw club names seen across all processed races
        aggregated: Dict[str, Dict] = {}
        for race_num, unrec_list in self.all_unrec_clubs.items():
            for unrec in unrec_list:
                key = unrec.raw_club_name.strip()
                if not key:
                    continue
                if key not in aggregated:
                    aggregated[key] = {"occurrences": 0, "races": set()}
                aggregated[key]["occurrences"] += unrec.occurrences
                aggregated[key]["races"].add(race_num)

        unrec_rows: List[dict] = []
        aggregated = self._aggregate_unrecognised_clubs_from_raw_data(aggregated)

        for raw_club, stats in sorted(aggregated.items(), key=lambda item: item[0].lower()):
            if raw_club in {"", "(blank)"}:
                continue
            best_match, confidence = _best_club_match(raw_club, self.preferred_clubs)
            races_seen = ", ".join(str(r) for r in sorted(stats["races"]))
            if best_match:
                message = (
                    f"Raw club '{raw_club}' is not in clubs.xlsx. "
                    f"Best match: '{best_match}' ({confidence}%)."
                )
            else:
                message = f"Raw club '{raw_club}' is not in clubs.xlsx and has no close match."
            unrec_rows.append(
                {
                    "Raw Club": raw_club,
                    "Best Match": best_match,
                    "Confidence": confidence if best_match else "",
                    "Occurrences": stats["occurrences"],
                    "Races Seen": races_seen,
                    "Status": "Manual Review",
                    "Message": message,
                }
            )

        summary_df = (
            pd.DataFrame(unrec_rows, columns=_UNREC_COLUMNS)
            if unrec_rows
            else pd.DataFrame(columns=_UNREC_COLUMNS)
        )
        return club_df, summary_df

    def _aggregate_unrecognised_clubs_from_raw_data(
        self,
        aggregated: dict[str, dict],
    ) -> dict[str, dict]:
        """Add raw_data unrecognised club names when audited rows have been blanked."""
        raw_data_dir = self.input_paths.raw_data_dir
        if not raw_data_dir.exists():
            return aggregated

        raw_race_files = discover_race_files(raw_data_dir, excluded_names=race_discovery_exclusions())
        for race_num, filepath in raw_race_files.items():
            try:
                df = load_race_dataframe(filepath)
            except Exception:
                continue

            club_columns = [c for c in df.columns if "club" in str(c).lower()]
            if not club_columns:
                continue

            club_col = club_columns[0]
            for raw_club in df[club_col].dropna().astype(str).map(str.strip):
                if not raw_club or raw_club.lower() == "nan":
                    continue
                raw_club_key = raw_club
                if raw_club_key.lower() in self.raw_to_preferred:
                    continue

                if raw_club_key not in aggregated:
                    aggregated[raw_club_key] = {"occurrences": 0, "races": set()}
                aggregated[raw_club_key]["occurrences"] += 1
                aggregated[raw_club_key]["races"].add(race_num)

        return aggregated

    def _build_ea_review_dfs(self, race_meta: Dict[int, dict]) -> Tuple[pd.DataFrame, pd.DataFrame]:
        candidate_rows: List[dict] = []
        checked_rows: List[dict] = []

        for race_num, meta in sorted(race_meta.items()):
            race_file = meta["file"].name
            for runner in self.all_race_runners.get(race_num, []):
                candidate = _build_ea_candidate_row(runner, race_num, race_file)
                if candidate is None:
                    continue

                candidate_rows.append(candidate)
                checked_rows.append(
                    {
                        "Name": candidate["Name"],
                        "Club": candidate["Club"],
                        "Race": candidate["Race"],
                        "Race File": candidate["Race File"],
                        "Source Category": candidate["Source Category"],
                        "Normalised Category": candidate["Normalised Category"],
                        "EA Category": "",
                        "League Category": "",
                        "Checked On": "",
                        "Notes": "",
                    }
                )

        candidates_df = pd.DataFrame(candidate_rows, columns=_EA_CANDIDATE_COLUMNS)
        checked_df = pd.DataFrame(checked_rows, columns=_EA_CHECKED_COLUMNS)
        return candidates_df, checked_df

    def _build_race_summary_df(
        self,
        race_meta: Dict[int, dict],
        row_df: pd.DataFrame,
        runner_df: pd.DataFrame,
        club_df: pd.DataFrame,
    ) -> pd.DataFrame:
        rows: List[dict] = []
        for race_num, meta in sorted(race_meta.items()):
            race_file = meta["file"].name
            race_issues = self.all_race_issues.get(race_num, [])
            race_rows = row_df[row_df["Race"] == race_num] if not row_df.empty else pd.DataFrame()
            race_runner_rows = runner_df[runner_df["Races Seen"].astype(str).str.contains(fr"\b{race_num}\b", regex=True)] if not runner_df.empty else pd.DataFrame()
            race_club_rows = club_df[club_df["Races Seen"].astype(str).str.contains(fr"\b{race_num}\b", regex=True)] if not club_df.empty and "Races Seen" in club_df.columns else pd.DataFrame()
            severities = list(race_rows.get("Severity", [])) + list(race_runner_rows.get("Severity", [])) + list(race_club_rows.get("Severity", []))
            for issue in race_issues:
                if issue.code.startswith("AUD-RACE"):
                    severities.append("error" if issue.code == "AUD-RACE-005" else "warning")
            errors = sum(1 for severity in severities if severity == "error")
            warnings = sum(1 for severity in severities if severity == "warning")
            infos = sum(1 for severity in severities if severity == "info")
            race_codes = sorted(set(race_rows.get("Issue Code", [])))
            race_codes.extend(issue.code for issue in race_issues if issue.code.startswith("AUD-RACE"))
            if meta["scheme"] == "EA 5-Year":
                race_codes.append("AUD-RACE-006")
            elif meta["scheme"] == "External Data Check":
                race_codes.append("AUD-RACE-007")
            race_codes = sorted(set(code for code in race_codes if code))

            status = "Open"
            depends_on = "None"
            summary = "Race audit generated."
            next_step = "Review the listed issues for this race."
            skipped_issue = next((issue for issue in race_issues if issue.code == "AUD-RACE-005"), None)
            if skipped_issue is not None:
                status = "Manual Review"
                depends_on = "Source Data Correction"
                summary = skipped_issue.message
                next_step = "Fix the race-level processing problem before relying on this race."
            if meta["scheme"] == "External Data Check":
                status = "Dependent"
                depends_on = "External Category Check"
                summary = f"Mixed female category evidence: {meta['scheme_evidence']}"
                next_step = "Confirm the race category scheme before trusting downstream category findings."
            elif meta["scheme"] == "EA 5-Year":
                summary = f"EA-style female category evidence: {meta['scheme_evidence']}"
                next_step = "Review derived categories and confirm the race used EA 5-year bands."

            if status == "Open" and any(code in race_codes for code in ("AUD-RACE-008", "AUD-RACE-009")):
                status = "Manual Review"
                depends_on = "Source Data Correction"
                summary = "Schema pre-check found header or required-field issues in this race workbook."
                next_step = "Review race-level schema warnings before relying on downstream row audit findings."

            rows.append(
                {
                    "Severity Max": _max_severity(severities),
                    "Race": race_num,
                    "Race File": race_file,
                    "Scheme Status": meta["scheme"],
                    "Error Count": errors,
                    "Warning Count": warnings,
                    "Info Count": infos,
                    "Issue Codes": ", ".join(race_codes),
                    "Status": status,
                    "Depends On": depends_on,
                    "Summary": summary,
                    "Next Step": next_step,
                }
            )

        return pd.DataFrame(rows, columns=_RACE_AUDIT_COLUMNS)


_RACE_AUDIT_COLUMNS = [
    "Severity Max",
    "Race",
    "Race File",
    "Scheme Status",
    "Error Count",
    "Warning Count",
    "Info Count",
    "Issue Codes",
    "Status",
    "Depends On",
    "Summary",
    "Next Step",
]

_ROW_AUDIT_COLUMNS = [
    "Severity",
    "Race",
    "Race File",
    "Source Row",
    "Issue Code",
    "Runner Name",
    "Time",
    "Club",
    "Sex",
    "Category",
    "Status",
    "Depends On",
    "Message",
    "Next Step",
]

_RUNNER_AUDIT_COLUMNS = [
    "Severity",
    "Runner Key",
    "Display Name",
    "Issue Code",
    "Clubs Seen",
    "Sexes Seen",
    "Categories Seen",
    "Races Seen",
    "Status",
    "Depends On",
    "Message",
    "Next Step",
]

_CLUB_AUDIT_COLUMNS = [
    "Severity",
    "Raw Club",
    "Issue Code",
    "Preferred Club",
    "Confidence",
    "Occurrences",
    "Races Seen",
    "Status",
    "Depends On",
    "Message",
    "Next Step",
]

_UNREC_COLUMNS = [
    "Raw Club",
    "Best Match",
    "Confidence",
    "Occurrences",
    "Races Seen",
    "Status",
    "Message",
]

_EA_CANDIDATE_COLUMNS = [
    "Name",
    "Club",
    "Race",
    "Race File",
    "Source Category",
    "Normalised Category",
    "Gender",
    "Reason",
]

_EA_CHECKED_COLUMNS = [
    "Name",
    "Club",
    "Race",
    "Race File",
    "Source Category",
    "Normalised Category",
    "EA Category",
    "League Category",
    "Checked On",
    "Notes",
]


def _build_row_entry(race_num: int, race_file: str, issue: RaceIssue, runner: Optional[RunnerRaceEntry], is_series: bool = False) -> dict:
    severity, status, depends_on, next_step = _status_for_code(issue.code, is_series=is_series)
    return {
        "Severity": severity,
        "Race": race_num,
        "Race File": race_file,
        "Source Row": issue.source_row or "",
        "Issue Code": issue.code,
        "Runner Name": issue.runner_name or (runner.name if runner else ""),
        "Time": issue.time_str or (runner.time_str if runner else ""),
        "Club": issue.raw_club or (_display_club(runner) if runner else ""),
        "Sex": issue.gender or (runner.gender if runner else ""),
        "Category": issue.raw_category or (runner.raw_category if runner else ""),
        "Status": status,
        "Depends On": depends_on,
        "Message": issue.message,
        "Next Step": next_step,
    }


def _build_runner_entry(
    code: str,
    severity: str,
    runner_key: str,
    display_name: str,
    clubs_seen: List[str],
    sexes_seen: List[str],
    categories_seen: List[str],
    races_seen: List[int],
    status: str,
    depends_on: str,
    message: str,
    next_step: str,
) -> dict:
    return {
        "Severity": severity,
        "Runner Key": runner_key,
        "Display Name": display_name,
        "Issue Code": code,
        "Clubs Seen": ", ".join(clubs_seen),
        "Sexes Seen": ", ".join(sexes_seen),
        "Categories Seen": ", ".join(categories_seen),
        "Races Seen": ", ".join(str(race) for race in races_seen),
        "Status": status,
        "Depends On": depends_on,
        "Message": message,
        "Next Step": next_step,
    }


def _status_for_code(code: str, is_series: bool = False) -> Tuple[str, str, str, str]:
    if code == "AUD-ROW-010" and is_series:
        return (
            "info",
            "Manual Review",
            "Duplicate Conflict Review",
            "Check conflicting club, sex, or category values in the source rows.",
        )
    mapping = {
        "AUD-ROW-001": ("warning", "Ready To Fix", "Source Data Correction", "Correct the gender value in the source workbook."),
        "AUD-ROW-002": ("warning", "Ready To Fix", "Source Data Correction", "Correct the time value in the source workbook."),
        "AUD-ROW-003": ("warning", "Open", "None", "Review the defaulted category and correct it in the source workbook if needed."),
        "AUD-ROW-004": ("warning", "Open", "None", "Review the category mapping and confirm the derived category is correct."),
        "AUD-ROW-005": ("warning", "Ready To Fix", "Source Data Correction", "Correct the missing runner name in the source workbook."),
        "AUD-ROW-006": ("warning", "Open", "Club Lookup", "Review the club suggestion and decide whether to add a club lookup conversion."),
        "AUD-ROW-008": ("warning", "Open", "None", "Review kept and discarded rows if traceability is needed."),
        "AUD-ROW-010": ("error", "Manual Review", "Duplicate Conflict Review", "Check conflicting club, sex, or category values in the source rows."),
        "AUD-ROW-011": ("warning", "Manual Review", "External Category Check", "Check the EA category workbook and confirm the league category for this female veteran row."),
    }
    return mapping.get(code, ("warning", "Open", "None", "Review the source data."))


def _classify_race_scheme(runners: Iterable[RunnerRaceEntry]) -> Tuple[str, str]:
    female_ages = sorted({age for age in (_extract_veteran_age(r.raw_category) for r in runners if r.gender == "F") if age is not None})
    if not female_ages:
        return "League Bands", "No female veteran categories detected"
    if all(age % 10 == 0 or age >= 70 for age in female_ages):
        return "League Bands", ", ".join(f"V{age}" if age < 70 else "V70+" for age in female_ages)
    if _has_consecutive_five_year_band(female_ages):
        return "EA 5-Year", ", ".join(f"V{age}" for age in female_ages)
    return "External Data Check", ", ".join(f"V{age}" for age in female_ages)


def _extract_veteran_age(raw_category: str) -> Optional[int]:
    if not raw_category:
        return None
    match = re.search(r"(?:^|[^\d])(\d{2})(?:\+)?", str(raw_category))
    if not match:
        return None
    age = int(match.group(1))
    return age if age >= 35 else None


def _has_consecutive_five_year_band(ages: List[int]) -> bool:
    age_set = set(ages)
    for age in age_set:
        if age + 5 in age_set:
            return True
    return False


def _derived_audit_category(runner: RunnerRaceEntry, scheme: str) -> str:
    if scheme != "EA 5-Year":
        return runner.normalised_category

    raw_age = _extract_veteran_age(runner.raw_category)
    if raw_age is None:
        return runner.normalised_category
    if raw_age < 40:
        return "Sen"
    if raw_age < 50:
        return "V40"
    if raw_age < 60:
        return "V50"
    if raw_age < 70:
        return "V60"
    return "V70+"


def _is_series_race(path: Path) -> bool:
    """Return True when the race file is a consolidated multi-event series."""
    stem = path.stem.lower()
    return "consolidated" in stem or "series" in stem


def _best_club_match(raw_club: str, preferred_clubs: List[str]) -> Tuple[str, int]:
    raw = raw_club.strip().lower()
    best_name = ""
    best_score = 0
    for preferred in preferred_clubs:
        score = int(round(SequenceMatcher(None, raw, preferred.lower()).ratio() * 100))
        if score > best_score:
            best_name = preferred
            best_score = score
    return best_name, best_score


def _inspect_club_lookup_file(filepath: Path) -> List[dict]:
    rows: List[dict] = []
    try:
        df = pd.read_excel(filepath, engine="openpyxl", dtype=str)
    except Exception:
        return rows

    df.columns = [str(col).strip() for col in df.columns]
    if not {"Club", "Preferred name", "Team A", "Team B"}.issubset(df.columns):
        return rows

    alias_map: Dict[str, set] = defaultdict(set)
    preferred_divisions: Dict[str, set] = defaultdict(set)
    for _, row in df.iterrows():
        raw = str(row.get("Club", "")).strip()
        preferred = str(row.get("Preferred name", "")).strip()
        if raw and raw.lower() != "nan" and preferred and preferred.lower() != "nan":
            alias_map[raw.lower()].add(preferred)
            preferred_divisions[preferred].add((str(row.get("Team A", "")).strip(), str(row.get("Team B", "")).strip()))

    for raw_alias, preferred_names in sorted(alias_map.items()):
        if len(preferred_names) > 1:
            rows.append(
                {
                    "Severity": "warning",
                    "Raw Club": raw_alias,
                    "Issue Code": "AUD-CLUB-003",
                    "Preferred Club": ", ".join(sorted(preferred_names)),
                    "Confidence": "",
                    "Occurrences": "",
                    "Races Seen": "",
                    "Status": "Manual Review",
                    "Depends On": "Club Lookup",
                    "Message": "The same raw club alias maps to more than one preferred club.",
                    "Next Step": "Clean up the club lookup so each raw alias resolves to one preferred club.",
                }
            )

    for preferred_name, divisions in sorted(preferred_divisions.items()):
        if len(divisions) > 1:
            rows.append(
                {
                    "Severity": "warning",
                    "Raw Club": "",
                    "Issue Code": "AUD-CLUB-002",
                    "Preferred Club": preferred_name,
                    "Confidence": "",
                    "Occurrences": "",
                    "Races Seen": "",
                    "Status": "Manual Review",
                    "Depends On": "Club Lookup",
                    "Message": "Preferred club has inconsistent Team A/Team B division values in the club lookup.",
                    "Next Step": "Review clubs.xlsx and keep a single consistent division assignment.",
                }
            )

    return rows


def _find_name_variant_clusters(
    by_identity: Dict[Tuple[str, str, str], List[RunnerRaceEntry]]
) -> List[dict]:
    keys = sorted(by_identity)
    rows: List[dict] = []
    seen_pairs: set[Tuple[Tuple[str, str, str], Tuple[str, str, str]]] = set()
    for idx, left_key in enumerate(keys):
        left_name, left_club, left_sex = left_key
        for right_key in keys[idx + 1:]:
            right_name, right_club, right_sex = right_key
            if left_club != right_club or left_sex != right_sex:
                continue
            ratio = SequenceMatcher(None, left_name, right_name).ratio()
            if ratio < 0.88 or left_name == right_name:
                continue
            pair = tuple(sorted((left_key, right_key)))
            if pair in seen_pairs:
                continue
            seen_pairs.add(pair)
            left_group = by_identity[left_key]
            right_group = by_identity[right_key]
            sample = left_group[0]
            rows.append(
                _build_runner_entry(
                    code="AUD-RUNNER-005",
                    severity="warning",
                    runner_key=f"{left_name}|{right_name}|variant",
                    display_name=f"{sample.name} / {right_group[0].name}",
                    clubs_seen=[_display_club(sample)],
                    sexes_seen=[left_sex],
                    categories_seen=sorted({r.normalised_category for r in left_group + right_group if r.normalised_category}),
                    races_seen=sorted({r.race_number for r in left_group + right_group}),
                    status="Manual Review",
                    depends_on="Identity Review",
                    message=f"Possible same-person name variant (confidence {int(round(ratio * 100))}%).",
                    next_step="Review the suggested name variant manually and correct source data if needed.",
                )
            )
    return rows


def _display_club(runner: Optional[RunnerRaceEntry]) -> str:
    if runner is None:
        return ""
    return runner.preferred_club or runner.raw_club


def _build_ea_candidate_row(
    runner: RunnerRaceEntry,
    race_num: int,
    race_file: str,
) -> Optional[dict]:
    if runner.gender != "F":
        return None

    candidate = _classify_ea_review_category(runner.raw_category)
    if candidate is None:
        return None

    return {
        "Name": runner.name,
        "Club": _display_club(runner),
        "Race": race_num,
        "Race File": race_file,
        "Source Category": runner.raw_category,
        "Normalised Category": candidate["normalised_category"],
        "Gender": runner.gender,
        "Reason": candidate["reason"],
    }


def _classify_ea_review_category(raw_category: str) -> Optional[dict]:
    stripped = _strip_category_sex_prefix(raw_category)
    stripped_lower = str(stripped).strip().lower()
    compact = _compact_category(stripped)
    if not compact:
        return None

    if _is_precise_ea_veteran_band(compact):
        return None

    ambiguous_reason = _ambiguous_veteran_reason(stripped_lower, compact)
    if ambiguous_reason:
        return {
            "normalised_category": _format_candidate_category(stripped),
            "reason": ambiguous_reason,
        }

    return None


def _strip_category_sex_prefix(raw_category: str) -> str:
    if not raw_category:
        return ""

    category = str(raw_category).strip()
    match = re.match(r"^([mfwlu])(?:\s*[-_ ]*)?(.*)$", category, flags=re.IGNORECASE)
    if not match:
        return category

    remainder = match.group(2).strip()
    if not remainder:
        return category

    lowered = remainder.lower()
    if lowered.startswith(("v", "vet", "u", "jun", "sen", "senior")) or re.match(r"^\d", lowered):
        return remainder
    return category


def _compact_category(value: str) -> str:
    return re.sub(r"[\s_\-.]", "", str(value).strip().lower())


def _is_precise_ea_veteran_band(compact: str) -> bool:
    prefixed_match = re.fullmatch(r"(?:v|vet)(\d{2})(\+)?", compact)
    if prefixed_match:
        return int(prefixed_match.group(1)) >= 35

    bare_age_match = re.fullmatch(r"(\d{2})", compact)
    if bare_age_match:
        return int(bare_age_match.group(1)) >= 35

    return False


def _ambiguous_veteran_reason(raw_lower: str, compact: str) -> str:
    if "vet" in compact or "veteran" in compact:
        if not _is_precise_ea_veteran_band(compact):
            return "Ambiguous veteran-style female category needs EA review."

    range_match = re.search(r"(\d{2})\s*(?:to|[-/])\s*(\d{2})", raw_lower)
    if range_match:
        low_age = int(range_match.group(1))
        high_age = int(range_match.group(2))
        if high_age - low_age == 9 and low_age >= 35:
            if low_age % 10 == 0 and high_age % 10 == 9:
                return ""
            return "Female veteran category uses a 5-4 boundary and needs EA review."
        if low_age >= 35 or high_age >= 35:
            return "Ambiguous veteran-style female category needs EA review."

    plus_match = re.search(r"(\d{2})\s*\+", raw_lower) or re.fullmatch(r"(\d{2})\+", compact)
    if plus_match:
        age = int(plus_match.group(1))
        if age >= 35:
            if age % 10 == 0:
                return ""
            return "Female veteran category uses a 5+ boundary and needs EA review."

    return ""


def _format_candidate_category(value: str) -> str:
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def _max_severity(severities: List[str]) -> str:
    if any(severity == "error" for severity in severities):
        return "error"
    if any(severity == "warning" for severity in severities):
        return "warning"
    if any(severity == "info" for severity in severities):
        return "info"
    return "info"