from league_scorer.models import RunnerRaceEntry, TeamRaceResult, ClubInfo
from league_scorer.season_aggregation import build_individual_season, build_team_season


def _runner(name: str, gender: str, race: int, points: int):
    return RunnerRaceEntry(
        name=name,
        raw_club="Club A",
        preferred_club="Club A",
        gender=gender,
        raw_category="Sen",
        normalised_category="Sen",
        time_str="00:40:00",
        time_seconds=100.0,
        race_number=race,
        eligible=True,
        points=points,
    )


def test_build_individual_season_uses_best_n(monkeypatch):
    from league_scorer import season_aggregation

    monkeypatch.setattr(season_aggregation.settings, "get", lambda key: 2 if key == "BEST_N" else None)

    all_races = {
        1: [_runner("Alex", "M", 1, 100)],
        2: [_runner("Alex", "M", 2, 90)],
        3: [_runner("Alex", "M", 3, 80)],
    }

    male, female = build_individual_season(all_races)
    assert len(female) == 0
    assert male[0].name == "Alex"
    assert male[0].total_points == 190


def test_build_team_season_uses_best_n(monkeypatch):
    from league_scorer import season_aggregation

    monkeypatch.setattr(season_aggregation.settings, "get", lambda key: 2 if key == "BEST_N" else None)

    club_info = {"Club A": ClubInfo(preferred_name="Club A", div_a=1, div_b=2)}
    all_teams = {
        1: [TeamRaceResult("Club A", "A", 1, 1, team_points=20, team_score=200)],
        2: [TeamRaceResult("Club A", "A", 1, 2, team_points=18, team_score=180)],
        3: [TeamRaceResult("Club A", "A", 1, 3, team_points=10, team_score=100)],
    }

    div1, div2 = build_team_season(all_teams, club_info)
    assert len(div2) == 1
    assert div1[0].total_points == 38
