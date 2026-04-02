from league_scorer.models import ClubInfo, RunnerRaceEntry
from league_scorer.team_scoring import build_team_scores


def _runner(name: str, gender: str, points: int) -> RunnerRaceEntry:
    return RunnerRaceEntry(
        name=name,
        raw_club="Club A",
        preferred_club="Club A",
        gender=gender,
        raw_category="Sen",
        normalised_category="Sen",
        time_str="00:40:00",
        time_seconds=100.0,
        race_number=1,
        eligible=True,
        points=points,
    )


def test_build_team_scores_uses_settings(monkeypatch):
    from league_scorer import team_scoring

    monkeypatch.setattr(team_scoring.settings, "get", lambda key: {"TEAM_SIZE": 2, "MAX_DIV_PTS": 20}[key])

    club_info = {"Club A": ClubInfo(preferred_name="Club A", div_a=1, div_b=2)}
    runners = [
        _runner("M1", "M", 100),
        _runner("M2", "M", 99),
        _runner("F1", "F", 98),
        _runner("F2", "F", 97),
        _runner("M3", "M", 96),
        _runner("F3", "F", 95),
    ]

    teams, updated = build_team_scores(runners, club_info, race_number=1)

    assert len(teams) == 2
    a_team = next(t for t in teams if t.team_id == "A")
    b_team = next(t for t in teams if t.team_id == "B")

    assert a_team.team_score == (100 + 99 + 98 + 97)
    assert b_team.team_score == (96 + 95)
    assert a_team.team_points >= 1
    assert b_team.team_points >= 0
    assert any(r.team_id == "A" for r in updated)
