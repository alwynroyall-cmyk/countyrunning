from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from league_scorer.output.output_layout import (
    build_output_paths,
    ensure_output_subdirs,
    package_publish_artifacts,
)
from league_scorer.publish import club_report, publish


def test_generate_club_reports_failure_writes_error_report(monkeypatch, tmp_path: Path) -> None:
    data_root = tmp_path / "data"
    report_dir = tmp_path / "reports"
    data_root.mkdir(parents=True, exist_ok=True)

    year = 1999
    input_dir = data_root / str(year) / "inputs"
    output_dir = data_root / str(year) / "outputs"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Force a fatal error during scoring so the handler writes a failure report.
    class BrokenScorer:
        def __init__(self, *args, **kwargs):
            raise RuntimeError("simulated scorer failure")

    monkeypatch.setattr(club_report, "LeagueScorer", BrokenScorer)

    result = club_report.generate_club_reports(year=year, data_root=data_root, report_dir=report_dir)

    assert result == 1
    report_path = report_dir / f"year-{year}"
    error_md = report_path / "club_reports.md"
    assert error_md.exists()
    content = error_md.read_text(encoding="utf-8")
    assert "Error:" in content
    assert "simulated scorer failure" in content


def test_publish_package_artifacts_copies_publish_files(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    # Create sample publish files in each relevant location.
    sample_files = [
        output_paths.publish_docx_league_updates_dir / "Season Update.docx",
        output_paths.publish_docx_race_cards_dir / "Race 1 - Scoring Card.docx",
        output_paths.publish_standings_dir / "Standings.xlsx",
    ]
    for sample in sample_files:
        sample.write_text("placeholder", encoding="utf-8")

    # Also include a club report docx in the club-reports folder.
    club_reports_dir = output_paths.publish_docx_club_reports_dir
    club_reports_dir.mkdir(parents=True, exist_ok=True)
    club_report_file = club_reports_dir / "club_reports_1999.docx"
    club_report_file.write_text("placeholder", encoding="utf-8")

    package_dir = package_publish_artifacts(output_dir)

    expected = [
        package_dir / "Season Update.docx",
        package_dir / "Race 1 - Scoring Card.docx",
        package_dir / "Standings.xlsx",
        package_dir / "club_reports_1999.docx",
    ]

    for expected_file in expected:
        assert expected_file.exists()


def test_publish_results_generates_club_reports(tmp_path: Path, monkeypatch) -> None:
    data_root = tmp_path / "data"
    year = 1999
    year_root = data_root / str(year)
    audited_dir = year_root / "inputs" / "audited"
    audited_dir.mkdir(parents=True, exist_ok=True)
    (audited_dir / "Race 1 - audited.xlsx").write_text("placeholder", encoding="utf-8")

    output_dir = year_root / "outputs"
    ensure_output_subdirs(output_dir)

    report_dir = tmp_path / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)

    called = {"count": 0}

    def fake_generate_club_reports(year_arg, data_root_arg, report_dir_arg):
        assert year_arg == year
        assert data_root_arg == data_root
        assert report_dir_arg == report_dir
        called["count"] += 1
        return 0

    monkeypatch.setattr(club_report, "generate_club_reports", fake_generate_club_reports)

    result = publish.publish_results(year=year, data_root=data_root, report_dir=report_dir)

    assert result == 0
    assert called["count"] == 1


def test_publish_package_artifacts_preserves_nested_structure_when_not_flattened(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    sample_files = [
        output_paths.publish_docx_league_updates_dir / "Season Update.docx",
        output_paths.publish_docx_race_cards_dir / "Race 1 - Scoring Card.docx",
        output_paths.publish_standings_dir / "Standings.xlsx",
    ]
    for sample in sample_files:
        sample.write_text("placeholder", encoding="utf-8")

    club_reports_dir = output_paths.publish_docx_club_reports_dir
    club_reports_dir.mkdir(parents=True, exist_ok=True)
    club_report_file = club_reports_dir / "club_reports_1999.docx"
    club_report_file.write_text("placeholder", encoding="utf-8")

    package_dir = package_publish_artifacts(output_dir, flatten=False)

    expected = [
        package_dir / "docx" / "Season Update.docx",
        package_dir / "docx" / "Race 1 - Scoring Card.docx",
        package_dir / "standings" / "Standings.xlsx",
        package_dir / "club-reports" / "club_reports_1999.docx",
    ]

    for expected_file in expected:
        assert expected_file.exists()




def test_publish_results_failure_writes_error_report(monkeypatch, tmp_path: Path) -> None:
    data_root = tmp_path / "data"
    year = 1999
    year_root = data_root / str(year)
    audited_dir = year_root / "inputs" / "audited"
    audited_dir.mkdir(parents=True, exist_ok=True)

    # Create one audited file so stage 1 succeeds.
    audited_file = audited_dir / "Race 1 - audited.xlsx"
    audited_file.write_text("placeholder", encoding="utf-8")

    report_dir = tmp_path / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)

    # Ensure publish flow reaches the conversion step and then fails.
    monkeypatch.setattr(publish, "build_output_paths", lambda _output_dir: (_ for _ in ()).throw(RuntimeError("simulated publish failure")))
    monkeypatch.setattr(publish, "discover_race_files", lambda _input, excluded_names=(): {1: audited_file})

    result = publish.publish_results(year=year, data_root=data_root, report_dir=report_dir)

    assert result == 1
    report_path = report_dir / f"year-{year}"
    error_md = report_path / "publish_results.md"
    assert error_md.exists()
    content = error_md.read_text(encoding="utf-8")
    assert "Error" in content
    assert "simulated publish failure" in content
