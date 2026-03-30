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
    RunnerRaceEntry,
    TeamRaceResult,
    UnrecognisedClub,
)
from .output_writer import (
    write_category_report,
    write_club_table,
    write_individual_table,
    write_summary,
    write_unrecognised_clubs,
)
from .race_processor import extract_race_number, process_race_file
from .season_aggregation import build_individual_season, build_team_season
from .team_scoring import build_team_scores

log = logging.getLogger(__name__)


class LeagueScorer:
    """Orchestrates the complete Wiltshire League scoring pipeline."""

    def __init__(self, input_dir: Path, output_dir: Path) -> None:
        self.input_dir = input_dir
        self.output_dir = output_dir

        self.raw_to_preferred: Dict[str, str] = {}
        self.club_info: Dict[str, ClubInfo] = {}

        # Accumulated by race number
        self.all_race_runners: Dict[int, List[RunnerRaceEntry]] = {}
        self.all_race_teams: Dict[int, List[TeamRaceResult]] = {}
        self.all_cat_records: Dict[int, List[CategoryRecord]] = {}
        self.all_unrec_clubs: Dict[int, List[UnrecognisedClub]] = {}

    # ─────────────────────────────────────────────────────────── public ──────

    def run(self) -> None:
        """Execute the full rerun pipeline."""
        self._load_clubs()

        race_files = self._discover_races()
        if not race_files:
            log.warning("No valid race files found in '%s'.", self.input_dir)
            return

        for race_num in sorted(race_files):
            self._process_race(race_num, race_files[race_num])

        if not self.all_race_runners:
            log.warning("No races were successfully processed — no output written.")
            return

        self._write_cumulative_outputs()
        log.info("Done. Results in '%s'.", self.output_dir)

    # ───────────────────────────────────────────────────────────── private ───

    def _load_clubs(self) -> None:
        self.raw_to_preferred, self.club_info = load_clubs(
            self.input_dir / "clubs.xlsx"
        )

    def _discover_races(self) -> Dict[int, Path]:
        """Find all xlsx files with a valid Race # name, sorted by race number."""
        races: Dict[int, Path] = {}
        for fp in sorted(self.input_dir.glob("*.xlsx")):
            if fp.name.lower() == "clubs.xlsx":
                continue
            n = extract_race_number(fp.stem)
            if n is None:
                log.warning("Ignoring '%s' — no valid race number in filename.", fp.name)
                continue
            if n in races:
                log.warning(
                    "Duplicate race number %d — ignoring '%s' (keeping '%s').",
                    n, fp.name, races[n].name,
                )
                continue
            races[n] = fp

        log.info("Discovered %d race file(s).", len(races))
        return races

    def _process_race(self, race_num: int, filepath: Path) -> None:
        """Process one race file end-to-end."""
        try:
            runners, cat_recs, unrec = process_race_file(
                filepath, race_num, self.raw_to_preferred
            )
        except RaceProcessingError as exc:
            log.error("Race %d SKIPPED — %s", race_num, exc)
            return

        runners = assign_individual_points(runners)
        team_results, runners = build_team_scores(runners, self.club_info, race_num)

        self.all_race_runners[race_num] = runners
        self.all_race_teams[race_num] = team_results
        self.all_cat_records[race_num] = cat_recs
        self.all_unrec_clubs[race_num] = unrec

        # Write per-race exception reports immediately
        pfx = f"Race {race_num} -- "
        write_category_report(cat_recs, self.output_dir / f"{pfx}categories.xlsx")
        write_unrecognised_clubs(unrec, self.output_dir / f"{pfx}unused clubs.xlsx")

        log.info("Race %d complete.", race_num)

    def _write_cumulative_outputs(self) -> None:
        """Aggregate all processed races and write cumulative output files."""
        highest = max(self.all_race_runners)
        pfx = f"Race {highest} -- "

        log.info("Aggregating results (highest race: %d)…", highest)

        male_recs, female_recs = build_individual_season(self.all_race_runners)
        div1_teams, div2_teams = build_team_season(self.all_race_teams, self.club_info)

        all_unrec = [u for lst in self.all_unrec_clubs.values() for u in lst]

        write_individual_table(
            male_recs,
            self.output_dir / f"{pfx}male_table.xlsx",
        )
        write_individual_table(
            female_recs,
            self.output_dir / f"{pfx}female_table.xlsx",
        )
        write_club_table(div1_teams, self.output_dir / f"{pfx}club_table_div1.xlsx")
        write_club_table(div2_teams, self.output_dir / f"{pfx}club_table_div2.xlsx")
        write_summary(
            highest,
            male_recs,
            female_recs,
            self.all_race_runners,
            all_unrec,
            self.output_dir / f"{pfx}summary.xlsx",
        )
        log.info
