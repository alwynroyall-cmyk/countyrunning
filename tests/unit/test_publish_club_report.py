from pathlib import Path
from types import SimpleNamespace
import json

from league_scorer.publish import club_report


class DummyTeam:
    def __init__(self, preferred_club: str, team_id: str, team_score: int, team_points: int, division: int, position: int, total_points: int):
        self.preferred_club = preferred_club
        self.team_id = team_id
        self.team_score = team_score
        self.team_points = team_points
        self.division = division
        self.position = position
        self.total_points = total_points
        self.race_results = {}


class DummyRunner:
    def __init__(self, name: str, preferred_club: str, gender: str, category: str, total_points: int, races_completed: int, race_points: dict[int, int], eligible: bool = True):
        self.name = name
        self.preferred_club = preferred_club
        self.gender = gender
        self.category = category
        self.total_points = total_points
        self.races_completed = races_completed
        self.race_points = race_points
        self.eligible = eligible
        self.race_results = {}
        self.position = 1


def test_generate_club_reports_success_writes_docx_and_report(monkeypatch, tmp_path: Path) -> None:
    data_root = tmp_path / "data"
    report_dir = tmp_path / "reports"
    data_root.mkdir(parents=True)
    report_dir.mkdir(parents=True)

    male_runner = DummyRunner(
        name="Alice",
        preferred_club="Test Club",
        gender="M",
        category="Sen",
        total_points=100,
        races_completed=5,
        race_points={1: 20, 2: 20, 3: 20, 4: 20, 5: 20},
    )
    female_runner = DummyRunner(
        name="Becky",
        preferred_club="Test Club",
        gender="F",
        category="Sen",
        total_points=90,
        races_completed=4,
        race_points={1: 22, 2: 22, 3: 23, 4: 23},
    )

    team_a = DummyTeam(
        preferred_club="Test Club",
        team_id="A",
        team_score=50,
        team_points=6,
        division=1,
        position=1,
        total_points=100,
    )
    team_a.race_results = {1: SimpleNamespace(men_score=50, women_score=0)}

    scorer = SimpleNamespace(
        all_race_runners={1: [male_runner, female_runner]},
        all_race_teams={1: [team_a]},
        club_info={"Test Club": "A"},
        run=lambda *args, **kwargs: [],
    )

    monkeypatch.setattr(club_report, "LeagueScorer", lambda *args, **kwargs: scorer)
    monkeypatch.setattr(club_report, "build_individual_season", lambda runners: ([male_runner], [female_runner]))
    monkeypatch.setattr(club_report, "build_team_season", lambda teams, club_info: ([team_a], []))

    result = club_report.generate_club_reports(year=1999, data_root=data_root, report_dir=report_dir)

    assert result == 0
    report_root = report_dir / "year-1999"
    assert (report_root / "club_reports.json").exists()
    assert (report_root / "club_reports.md").exists()

    payload = json.loads((report_root / "club_reports.json").read_text(encoding="utf-8"))
    assert payload["success"] is True
    assert payload["docx"].endswith("club_reports_1999.docx")
    assert Path(payload["docx"]).exists()