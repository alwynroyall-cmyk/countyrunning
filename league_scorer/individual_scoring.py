"""
Assign per-race individual points.

Spec §9:
  • Gender-separated scoring.
  • 1st = 100 pts, minimum = 5 pts.
  • Ties → competition ranking (1,1,3,3,5…): tied runners share the same
    points; the next rank skips the tied positions.
  • Ineligible runners receive 0 pts.
"""

import logging
from typing import List

from .models import RunnerRaceEntry

log = logging.getLogger(__name__)

MAX_PTS = 100
MIN_PTS = 5


def assign_individual_points(runners: List[RunnerRaceEntry]) -> List[RunnerRaceEntry]:
    """Assign points in-place and return the runners list."""
    male_elig = [r for r in runners if r.gender == "M" and r.eligible]
    female_elig = [r for r in runners if r.gender == "F" and r.eligible]

    _score_group(male_elig)
    _score_group(female_elig)

    # Ineligible runners always get 0
    for r in runners:
        if not r.eligible:
            r.points = 0

    log.info(
        "  Individual points: %d male eligible, %d female eligible",
        len(male_elig), len(female_elig),
    )
    return runners


def _score_group(runners: List[RunnerRaceEntry]) -> None:
    """Sort by time, then apply competition-ranked points with tie sharing."""
    if not runners:
        return

    runners.sort(key=lambda r: r.time_seconds)

    n = len(runners)
    comp_pos = 1   # 1-based competition position
    i = 0

    while i < n:
        # Find the extent of the current tie group
        j = i + 1
        while j < n and runners[j].time_seconds == runners[i].time_seconds:
            j += 1

        pts = max(MIN_PTS, MAX_PTS - comp_pos + 1)
        for k in range(i, j):
            runners[k].points = pts

        # Competition ranking: next position skips past all tied runners
        comp_pos += (j - i)
        i = j
