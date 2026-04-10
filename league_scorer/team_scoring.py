"""
Build A/B team scores per race and award division points.

Spec §10:
    • A Team = top-N men + top-N women (by individual race points), N = TEAM_SIZE.
    • B Team = next-N men + next-N women.
    • Fewer than N of a gender -> use all available.
  • Team Score = sum(men points) + sum(women points).
    • Each division ranked independently: 1st = MAX_DIV_PTS ... min 1 pt.
  • No runners → 0 pts (not 1).
"""

import logging
from typing import Dict, List, Tuple

from .models import ClubInfo, RunnerRaceEntry, TeamRaceResult

log = logging.getLogger(__name__)

from .rules import get_team_size, get_max_div_pts
MIN_DIV_PTS = 1


def build_team_scores(
    runners: List[RunnerRaceEntry],
    club_info: Dict[str, ClubInfo],
    race_number: int,
) -> Tuple[List[TeamRaceResult], List[RunnerRaceEntry]]:
    """
    Compute team scores and assign division points.

    Mutates runner.team_id ('A', 'B', or '').
    Returns (team_results, updated runners).
    """
    team_size = get_team_size()

    club_runners = _collect_club_runners(runners, club_info)
    _sort_club_runners(club_runners)

    team_results: List[TeamRaceResult] = []
    for preferred_club, info in club_info.items():
        team_results.extend(
            _build_team_results_for_club(
                preferred_club,
                info,
                club_runners[preferred_club],
                team_size,
                race_number,
            )
        )

    _assign_division_points(team_results, division=1)
    _assign_division_points(team_results, division=2)

    log.info(
        "  Team scores: %d teams across 2 divisions",
        len(team_results),
    )
    return team_results, runners


def _collect_club_runners(
    runners: List[RunnerRaceEntry],
    club_info: Dict[str, ClubInfo],
) -> Dict[str, Dict[str, List[RunnerRaceEntry]]]:
    club_runners: Dict[str, Dict[str, List[RunnerRaceEntry]]] = {
        club: {"M": [], "F": []} for club in club_info
    }

    for runner in runners:
        if runner.eligible and runner.preferred_club in club_runners:
            club_runners[runner.preferred_club][runner.gender].append(runner)

    return club_runners


def _sort_club_runners(club_runners: Dict[str, Dict[str, List[RunnerRaceEntry]]]) -> None:
    for club in club_runners:
        for gender in ("M", "F"):
            club_runners[club][gender].sort(key=lambda r: r.points, reverse=True)


def _build_team_results_for_club(
    preferred_club: str,
    info: ClubInfo,
    runners_by_gender: Dict[str, List[RunnerRaceEntry]],
    team_size: int,
    race_number: int,
) -> List[TeamRaceResult]:
    results: List[TeamRaceResult] = []
    for team_id, division in (("A", info.div_a), ("B", info.div_b)):
        offset = 0 if team_id == "A" else team_size
        men_slice = runners_by_gender["M"][offset: offset + team_size]
        women_slice = runners_by_gender["F"][offset: offset + team_size]

        _assign_runner_team_ids(men_slice + women_slice, team_id)

        men_score = sum(r.points for r in men_slice)
        women_score = sum(r.points for r in women_slice)
        team_score = men_score + women_score

        results.append(
            TeamRaceResult(
                preferred_club=preferred_club,
                team_id=team_id,
                division=division,
                race_number=race_number,
                men_score=men_score,
                women_score=women_score,
                team_score=team_score,
            )
        )

    return results


def _assign_runner_team_ids(runners: List[RunnerRaceEntry], team_id: str) -> None:
    for runner in runners:
        runner.team_id = team_id


def _assign_division_points(teams: List[TeamRaceResult], division: int) -> None:
    """
    Rank teams within one division and award 20→1 pts.
    Teams with team_score == 0 receive 0 pts (no runners).
    Ties share the same points (competition ranking).
    """
    div_teams = [t for t in teams if t.division == division]
    if not div_teams:
        return

    max_div_pts = get_max_div_pts()

    with_runners = sorted(
        [t for t in div_teams if t.team_score > 0],
        key=lambda t: t.team_score,
        reverse=True,
    )
    no_runners = [t for t in div_teams if t.team_score == 0]

    rank = 1
    i = 0
    while i < len(with_runners):
        j = i + 1
        while j < len(with_runners) and with_runners[j].team_score == with_runners[i].team_score:
            j += 1
        pts = max(MIN_DIV_PTS, max_div_pts - rank + 1)
        for k in range(i, j):
            with_runners[k].team_points = pts
        rank += j - i
        i = j

    for team in no_runners:
        team.team_points = 0
