from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class ClubInfo:
    preferred_name: str
    div_a: int  # 1 or 2
    div_b: int  # 1 or 2


@dataclass
class RunnerRaceEntry:
    name: str
    raw_club: str
    preferred_club: Optional[str]   # None = unrecognised club
    gender: str                      # 'M' or 'F'
    raw_category: str
    normalised_category: str
    time_str: str                    # display string (preserved or converted)
    time_seconds: float              # float seconds for comparison
    race_number: int
    eligible: bool
    points: int = 0                  # individual race points (0 if ineligible)
    team_id: str = ""                # 'A', 'B', or '' (not in any team)


@dataclass
class CategoryRecord:
    """Tracks a single raw→normalised mapping for the exception report."""
    raw_category: str
    normalised_category: str
    count: int = 0
    notes: str = ""


@dataclass
class UnrecognisedClub:
    """Tracks an unrecognised club name for the exception report."""
    raw_club_name: str
    occurrences: int = 0


@dataclass
class TeamRaceResult:
    preferred_club: str
    team_id: str      # 'A' or 'B'
    division: int
    race_number: int
    men_score: int = 0
    women_score: int = 0
    team_score: int = 0
    team_points: int = 0  # division points (20→1, or 0 for no runners)


@dataclass
class RunnerSeasonRecord:
    name: str
    preferred_club: str
    gender: str
    category: str       # fixed at first appearance
    race_times: Dict[int, str] = field(default_factory=dict)
    race_points: Dict[int, int] = field(default_factory=dict)
    total_points: int = 0
    races_completed: int = 0
    position: int = 0


@dataclass
class TeamSeasonRecord:
    preferred_club: str
    team_id: str        # 'A' or 'B'
    division: int
    race_results: Dict[int, TeamRaceResult] = field(default_factory=dict)
    total_points: int = 0
    position: int = 0

    @property
    def display_name(self) -> str:
        return f"{self.preferred_club} -- {self.team_id}"
