from league_scorer.individual_scoring import assign_individual_points
from league_scorer.models import RunnerRaceEntry


def _runner(name: str, gender: str, time_s: float, eligible: bool = True) -> RunnerRaceEntry:
    return RunnerRaceEntry(
        name=name,
        raw_club="Club",
        preferred_club="Club" if eligible else None,
        gender=gender,
        raw_category="Sen",
        normalised_category="Sen",
        time_str="00:40:00",
        time_seconds=time_s,
        race_number=1,
        eligible=eligible,
    )


def test_assign_individual_points_ties_and_ineligible():
    runners = [
        _runner("A", "M", 100.0, True),
        _runner("B", "M", 100.0, True),
        _runner("C", "M", 110.0, True),
        _runner("D", "M", 120.0, False),
    ]

    assign_individual_points(runners)

    by_name = {r.name: r.points for r in runners}
    assert by_name["A"] == 100
    assert by_name["B"] == 100
    assert by_name["C"] == 98
    assert by_name["D"] == 0
