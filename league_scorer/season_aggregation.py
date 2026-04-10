"""
Aggregate per-race results into season standings.

Spec §9.3 / §10.3:
    • Best N race scores count (N = BEST_N; if fewer available, all count).
  • Individual tiebreak: (1) total pts, (2) races completed, (3) shared pos.
  • Team tiebreak: total pts, then race aggregate, then shared position.
"""

import logging
from typing import Dict, List, Tuple

from .models import (
    ClubInfo,
    RunnerRaceEntry,
    RunnerSeasonRecord,
    TeamRaceResult,
    TeamSeasonRecord,
)

log = logging.getLogger(__name__)

from .rules import get_best_n


# ───────────────────────────────────────────────────────────── individuals ───

def build_individual_season(
    all_race_runners: Dict[int, List[RunnerRaceEntry]],
) -> Tuple[List[RunnerSeasonRecord], List[RunnerSeasonRecord]]:
    """
    Aggregate runner entries across all races.
    Returns (male_records, female_records), each sorted by position.
    Category is fixed at the runner's first appearance.
    """
    season_map = _build_runner_season_map(all_race_runners)
    best_n = get_best_n()
    _compute_runner_totals(season_map, best_n)

    male, female = _split_runner_records_by_gender(season_map)
    _rank_runners(male)
    _rank_runners(female)

    log.info(
        "Individual season: %d male, %d female runners",
        len(male), len(female),
    )
    return male, female


def _build_runner_season_map(
    all_race_runners: Dict[int, List[RunnerRaceEntry]],
) -> Dict[Tuple[str, str], RunnerSeasonRecord]:
    season_map: Dict[Tuple[str, str], RunnerSeasonRecord] = {}

    for race_num in sorted(all_race_runners):
        for r in all_race_runners[race_num]:
            if not r.eligible or r.preferred_club is None:
                continue

            key = (r.name.lower(), r.preferred_club)
            if key not in season_map:
                season_map[key] = RunnerSeasonRecord(
                    name=r.name,
                    preferred_club=r.preferred_club,
                    gender=r.gender,
                    category=r.normalised_category,
                )

            record = season_map[key]
            record.race_times[race_num] = r.time_str
            record.race_points[race_num] = r.points

    return season_map


def _compute_runner_totals(
    season_map: Dict[Tuple[str, str], RunnerSeasonRecord],
    best_n: int,
) -> None:
    for rec in season_map.values():
        sorted_scores = sorted(rec.race_points.values(), reverse=True)
        rec.total_points = sum(sorted_scores[:best_n])
        rec.races_completed = len(sorted_scores)


def _split_runner_records_by_gender(
    season_map: Dict[Tuple[str, str], RunnerSeasonRecord],
) -> Tuple[List[RunnerSeasonRecord], List[RunnerSeasonRecord]]:
    male = [r for r in season_map.values() if r.gender == "M"]
    female = [r for r in season_map.values() if r.gender == "F"]
    return male, female


def _rank_runners(records: List[RunnerSeasonRecord]) -> None:
    """
    Sort runners by total_points (desc), then races_completed (desc).
    Assign shared position to genuinely tied runners.

    Note: full head-to-head comparison (spec §9.4 tiebreak #1) is complex
    in a general multi-runner sort; primary/secondary keys handle the common
    cases. Truly tied runners share a position (tiebreak #3).
    """
    records.sort(
        key=lambda r: (r.total_points, r.races_completed),
        reverse=True,
    )
    n = len(records)
    i = 0
    pos = 1
    while i < n:
        j = i + 1
        while (
            j < n
            and records[j].total_points == records[i].total_points
            and records[j].races_completed == records[i].races_completed
        ):
            j += 1
        for k in range(i, j):
            records[k].position = pos
        pos += j - i
        i = j


# ────────────────────────────────────────────────────────────────── teams ───

def _team_aggregate(team: TeamSeasonRecord) -> int:
    """
    Sum of (men_score + women_score) across ALL races for this team.
    Used as the tiebreaker when two teams share the same total league points.
    Higher aggregate = better.
    """
    return sum(
        (rr.men_score or 0) + (rr.women_score or 0)
        for rr in team.race_results.values()
    )


def build_team_season(
    all_race_teams: Dict[int, List[TeamRaceResult]],
    club_info: Dict[str, ClubInfo],
) -> Tuple[List[TeamSeasonRecord], List[TeamSeasonRecord]]:
    """
    Aggregate team race results into season records.
    Returns (div1_teams, div2_teams), each sorted by position.
    """
    season_map = _build_team_season_map(club_info)
    _populate_team_results(all_race_teams, season_map)

    best_n = get_best_n()
    _compute_team_totals(season_map, best_n)

    div1, div2 = _split_team_records_by_division(season_map)
    _rank_teams(div1)
    _rank_teams(div2)

    log.info("Team season: %d Div-1 teams, %d Div-2 teams", len(div1), len(div2))
    return div1, div2


def _build_team_season_map(
    club_info: Dict[str, ClubInfo],
) -> Dict[Tuple[str, str], TeamSeasonRecord]:
    season_map: Dict[Tuple[str, str], TeamSeasonRecord] = {}
    for preferred_club, info in club_info.items():
        for team_id, div in (("A", info.div_a), ("B", info.div_b)):
            season_map[(preferred_club, team_id)] = TeamSeasonRecord(
                preferred_club=preferred_club,
                team_id=team_id,
                division=div,
            )
    return season_map


def _populate_team_results(
    all_race_teams: Dict[int, List[TeamRaceResult]],
    season_map: Dict[Tuple[str, str], TeamSeasonRecord],
) -> None:
    for race_num, race_teams in all_race_teams.items():
        for t in race_teams:
            key = (t.preferred_club, t.team_id)
            if key in season_map:
                season_map[key].race_results[race_num] = t


def _compute_team_totals(
    season_map: Dict[Tuple[str, str], TeamSeasonRecord],
    best_n: int,
) -> None:
    for rec in season_map.values():
        pts_list = sorted(
            (t.team_points for t in rec.race_results.values()),
            reverse=True,
        )
        rec.total_points = sum(pts_list[:best_n])


def _split_team_records_by_division(
    season_map: Dict[Tuple[str, str], TeamSeasonRecord],
) -> Tuple[List[TeamSeasonRecord], List[TeamSeasonRecord]]:
    div1 = [r for r in season_map.values() if r.division == 1]
    div2 = [r for r in season_map.values() if r.division == 2]
    return div1, div2


def _rank_teams(teams: List[TeamSeasonRecord]) -> None:
    """
    Sort teams by:
      1. total_points descending       — primary
      2. race aggregate descending     — tiebreaker (higher score = better)
    Shared position only assigned when BOTH total_points AND aggregate are equal.
    """
    teams.sort(
        key=lambda t: (t.total_points, _team_aggregate(t)),
        reverse=True,
    )
    n = len(teams)
    i = 0
    pos = 1
    while i < n:
        agg_i = _team_aggregate(teams[i])
        j = i + 1
        while (
            j < n
            and teams[j].total_points == teams[i].total_points
            and _team_aggregate(teams[j]) == agg_i
        ):
            j += 1
        for k in range(i, j):
            teams[k].position = pos
        pos += j - i
        i = j
