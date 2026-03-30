"""
Build A/B team scores per race and award division points.

Spec §10:
  • A Team = top-5 men + top-5 women (by individual race points).
  • B Team = next-5 men + next-5 women (positions 6-10).
  • Fewer than 5 of a gender → use all available.
  • Team Score = sum(men points) + sum(women points).
  • Each division ranked independently: 1st = 20 pts … min 1 pt.
  • No runners → 0 pts (not 1).
"""

import logging
from typing import Dict, List, Tuple

from .models import ClubInfo, RunnerRaceEntry, TeamRaceResult

log = logging.getLogger(__name__)

TEAM_SIZE = 5
MAX_DIV_PTS = 20
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
    # Initialise empty gender groups for every club
    club_runners: Dict[str, Dict[str, List[RunnerRaceEntry]]] = {
        club: {"M": [], "F": []} for club in club_info
    }

    for r in runners:
        if r.eligible and r.preferred_club in club_runners:
            club_runners[r.preferred_club][r.gender].append(r)

    # Sort each group by individual points descending (best scorer first)
    for club in club_runners:
        for g in ("M", "F"):
            club_runners[club][g].sort(key=lambda r: r.points, reverse=True)

    team_results: List[TeamRaceResult] = []

    for preferred_club, info in club_info.items():
        for team_id, division in (("A", info.div_a), ("B", info.div_b)):
            offset = 0 if team_id == "A" else TEAM_SIZE
            men_slice = club_runners[preferred_club]["M"][offset: offset + TEAM_SIZE]
            women_slice = club_runners[preferred_club]["F"][offset: offset + TEAM_SIZE]

            # Tag runners with their team membership
            for r in men_slice:
                r.team_id = team_id
            for r in women_slice:
                r.team_id = team_id

            men_score = sum(r.points for r in men_slice)
            women_score = sum(r.points for r in women_slice)

            team_results.append(
                TeamRaceResult(
                    preferred_club=preferred_club,
                    team_id=team_id,
                    division=division,
                    race_number=race_number,
                    men_score=men_score,
                    women_score=women_score,
                    team_score=men_score + women_score,
                )
            )

    # Award division points independently per division
    _assign_division_points(team_results, division=1)
    _assign_division_points(team_results, division=2)

    log.info(
        "  Team scores: %d teams across 2 divisions",
        len(team_results),
    )
    return team_results, runners


def _assign_division_points(teams: List[TeamRaceResult], division: int) -> None:
    """
    Rank teams within one division and award 20→1 pts.
    Teams with team_score == 0 receive 0 pts (no runners).
    Ties share the same points (competition ranking).
    """
    div_teams = [t for t in teams if t.division == division]
    if not div_teams:
        return

    with_runners = sorted(
        [t for t in div_teams if t.team_score > 0],
        key=lambda t: t.team_score,
        reverse=True,
    )
    no_runners = [t for t in div_teams if t.team_score == 0]

    # Competition ranking
    n = len(with_runners)
    rank = 1
    i = 0
    while i < n:
        j = i + 1
        while j < n and with_runners[j].team_score == with_runners[i].team_score:
            j += 1
        pts = max(MIN_DIV_PTS, MAX_DIV_PTS - rank + 1)
        for k in range(i, j):
            with_runners[k].team_points = pts
        rank += j - i
        i = j

    for t in no_runners:
        t.team_points = 0
