from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class ClubInfo:
    preferred_name: str
    div_a: int  # 1 or 2
    div_b: int  # 1 or 2

    def __post_init__(self) -> None:
        if self.div_a not in {1, 2}:
            raise ValueError(f"ClubInfo.div_a must be 1 or 2, got {self.div_a!r}")
        if self.div_b not in {1, 2}:
            raise ValueError(f"ClubInfo.div_b must be 1 or 2, got {self.div_b!r}")


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
    source_row: int = 0
    points: int = 0                  # individual race points (0 if ineligible)
    team_id: str = ""                # 'A', 'B', or '' (not in any team)
    warnings: List[str] = field(default_factory=list)


@dataclass
class RaceIssue:
    kind: str                        # 'warning' or 'other'
    message: str
    source_row: Optional[int] = None
    code: str = ""
    runner_name: str = ""
    raw_club: str = ""
    gender: str = ""
    raw_category: str = ""
    time_str: str = ""


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

    def __post_init__(self) -> None:
        if self.team_id not in {"A", "B"}:
            raise ValueError(f"TeamRaceResult.team_id must be 'A' or 'B', got {self.team_id!r}")
        if self.division not in {1, 2}:
            raise ValueError(f"TeamRaceResult.division must be 1 or 2, got {self.division!r}")


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
