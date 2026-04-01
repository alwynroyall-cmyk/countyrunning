"""
Main orchestrator — implements the full rerun processing model.

Pipeline per run:
  1. Load clubs.xlsx (fatal if missing / malformed)
  2. Discover all Race # xlsx files
  3. For each race (ascending order):
       a. Load & normalise
       b. Assign individual points
       c. Build team scores
       d. Write per-race exception reports
  4. Aggregate season totals
  5. Write cumulative individual / club / summary outputs
"""

import logging
from pathlib import Path
from typing import Dict, List

from .club_loader import load_clubs
from .exceptions import FatalError, RaceProcessingError
from .individual_scoring import assign_individual_points
from .models import (
    CategoryRecord,
    ClubInfo,
    RaceIssue,
    RunnerRaceEntry,
    TeamRaceResult,
    UnrecognisedClub,
)
from .output_writer import (
    write_results_workbook,
    write_unrecognised_clubs,
)
from .report_writer import write_combined_report, write_race_report
from .race_processor import process_race_file
from .season_aggregation import build_individual_season, build_team_season
from .source_loader import discover_race_files
from .team_scoring import build_team_scores

log = logging.getLogger(__name__)


class LeagueScorer:
    """Orchestrates the complete Wiltshire League scoring pipeline."""

    def __init__(self, input_dir: Path, output_dir: Path, year: int) -> None:
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.year = int(year)

        self.raw_to_preferred: Dict[str, str] = {}
        self.club_info: Dict[str, ClubInfo] = {}

        # Accumulated by race number
        self.all_race_runners: Dict[int, List[RunnerRaceEntry]] = {}
        self.all_race_teams: Dict[int, List[TeamRaceResult]] = {}
        self.all_cat_records: Dict[int, List[CategoryRecord]] = {}
        self.all_unrec_clubs: Dict[int, List[UnrecognisedClub]] = {}
        self.all_race_issues: Dict[int, List[RaceIssue]] = {}
        self.selected_race_files: Dict[int, Path] = {}
        self.run_warnings: List[str] = []

    # ─────────────────────────────────────────────────────────── public ──────

    def run(self, race_files: "Dict[int, Path] | None" = None) -> List[str]:
        """Execute the full rerun pipeline.

        Parameters
        ----------
        race_files:
            Optional pre-selected mapping of race_number -> Path.
            If None, all race files in input_dir are discovered automatically.
        """
        self._validate_paths()
        self._load_clubs()

        if race_files is None:
            race_files = self._discover_races()

        self.selected_race_files = dict(race_files)

        if not race_files:
            log.warning("No valid race files found in '%s'.", self.input_dir)
            return []

        for race_num in sorted(race_files):
            self._process_race(race_num, race_files[race_num], len(race_files))

        if not self.all_race_runners:
            log.warning("No races were successfully processed — no output written.")
            return []

        self._write_cumulative_outputs()
        log.info("Done. Results in '%s'.", self.output_dir)
        return list(self.run_warnings)

    # ───────────────────────────────────────────────────────────── private ───

    def _validate_paths(self) -> None:
        """Validate the expected season directory layout before processing."""
        if self.input_dir.name != "inputs":
            raise FatalError(
                f"Input folder must end with 'inputs', got: {self.input_dir}"
            )
        if self.output_dir.name != "outputs":
            raise FatalError(
                f"Output folder must end with 'outputs', got: {self.output_dir}"
            )

        input_year = self.input_dir.parent.name
        output_year = self.output_dir.parent.name
        expected_year = str(self.year)

        if input_year != expected_year:
            raise FatalError(
                "Configured season year does not match input directory layout. "
                f"Expected .../{expected_year}/inputs, got: {self.input_dir}"
            )
        if output_year != expected_year:
            raise FatalError(
                "Configured season year does not match output directory layout. "
                f"Expected .../{expected_year}/outputs, got: {self.output_dir}"
            )

    def _load_clubs(self) -> None:
        self.raw_to_preferred, self.club_info = load_clubs(
            self.input_dir / "clubs.xlsx"
        )

    def _discover_races(self) -> Dict[int, Path]:
        """Find all supported race files with a valid Race # name."""
        races = discover_race_files(self.input_dir, excluded_names=("clubs.xlsx",))
        log.info("Discovered %d race file(s).", len(races))
        return races

    def _process_race(self, race_num: int, filepath: Path, total_races: int = 1) -> None:
        """Process one race file end-to-end."""
        try:
            runners, cat_recs, unrec, issue_notes = process_race_file(
                filepath, race_num, self.raw_to_preferred
            )
        except RaceProcessingError as exc:
            self.all_race_issues[race_num] = [RaceIssue("other", f"Race skipped: {exc}")]
            log.error("Race %d SKIPPED — %s", race_num, exc)
            return

        runners = assign_individual_points(runners)
        team_results, runners = build_team_scores(runners, self.club_info, race_num)

        self.all_race_runners[race_num] = runners
        self.all_race_teams[race_num] = team_results
        self.all_cat_records[race_num] = cat_recs
        self.all_unrec_clubs[race_num] = unrec
        self.all_race_issues[race_num] = list(issue_notes)

        # Write per-race exception reports immediately
        pfx = f"Race {race_num} -- "
        unrec_out = self.output_dir / f"{pfx}unused clubs.xlsx"
        write_unrecognised_clubs(unrec, unrec_out)
        if unrec:
            log.info(
                "Race %d: %d non-league club(s) excluded — see %s",
                race_num, len(unrec), unrec_out.name,
            )

        # Write branded per-race scoring card
        images_dir = Path(__file__).parent / "images"
        pdf_warning = write_race_report(
            race_num=race_num,
            total_races=total_races,
            runners=runners,
            team_results=team_results,
            images_dir=images_dir,
            year=self.year,
            filepath=self.output_dir / f"{pfx}race report",
            source_file=filepath,
        )
        if pdf_warning:
            self.all_race_issues.setdefault(race_num, []).append(RaceIssue("other", pdf_warning))
            self.run_warnings.append(pdf_warning)

        log.info("Race %d complete.", race_num)

    def _write_cumulative_outputs(self) -> None:
        """Aggregate all processed races and write cumulative output files."""
        highest = max(self.all_race_runners)
        pfx = f"Race {highest} -- "

        log.info("Aggregating results (highest race: %d)…", highest)

        male_recs, female_recs = build_individual_season(self.all_race_runners)
        div1_teams, div2_teams = build_team_season(self.all_race_teams, self.club_info)

        all_unrec = [u for lst in self.all_unrec_clubs.values() for u in lst]
        write_results_workbook(
            highest_race=highest,
            male_records=male_recs,
            female_records=female_recs,
            div1_teams=div1_teams,
            div2_teams=div2_teams,
            all_race_runners=self.all_race_runners,
            race_files=self.selected_race_files,
            all_unrec_clubs=self.all_unrec_clubs,
            race_issues=self.all_race_issues,
            filepath=self.output_dir / f"{pfx}Results.xlsx",
        )

        images_dir = Path(__file__).parent / "images"

        pdf_warning = write_combined_report(
            highest_race=highest,
            year=self.year,
            male_records=male_recs,
            female_records=female_recs,
            div1_teams=div1_teams,
            div2_teams=div2_teams,
            unrec_all=all_unrec,
            images_dir=images_dir,
            filepath=self.output_dir / f"{pfx}League Update Report",
        )
        if pdf_warning:
            self.run_warnings.append(pdf_warning)

        log.info("Cumulative outputs written successfully.")
