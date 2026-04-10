from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from league_scorer.output.output_layout import (
    build_output_paths,
    ensure_output_subdirs,
    export_publish_pdfs,
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
    error_json = report_path / "club_reports.json"
    assert error_json.exists()

    payload = json.loads(error_json.read_text(encoding="utf-8"))
    assert payload["success"] is False
    assert payload["docx"] == ""
    assert "simulated scorer failure" in payload["error"]["message"]


def test_publish_package_artifacts_copies_publish_files(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    # Create sample publish files in each relevant location.
    sample_files = [
        output_paths.publish_docx_league_updates_dir / "Season Update.docx",
        output_paths.publish_docx_race_cards_dir / "Race 1 - Scoring Card.docx",
        output_paths.publish_pdf_league_updates_dir / "Season Update.pdf",
        output_paths.publish_pdf_race_cards_dir / "Race 1 - Scoring Card.pdf",
        output_paths.publish_standings_dir / "Standings.xlsx",
        output_paths.publish_review_packs_dir / "Category Review.xlsx",
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
        package_dir / "Season Update.pdf",
        package_dir / "Race 1 - Scoring Card.pdf",
        package_dir / "Standings.xlsx",
        package_dir / "Category Review.xlsx",
        package_dir / "club_reports_1999.docx",
    ]

    for expected_file in expected:
        assert expected_file.exists()


def test_export_publish_pdfs_copies_all_pdfs_to_single_folder(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    sample_pdfs = [
        output_paths.publish_pdf_league_updates_dir / "Season Update.pdf",
        output_paths.publish_pdf_race_cards_dir / "Race 1 - Scoring Card.pdf",
    ]
    for sample in sample_pdfs:
        sample.write_text("placeholder", encoding="utf-8")

    export_dir = tmp_path / "exported_pdfs"
    result_path = export_publish_pdfs(output_dir, export_dir, flatten=True)

    assert result_path == export_dir
    assert (export_dir / "Season Update.pdf").exists()
    assert (export_dir / "Race 1 - Scoring Card.pdf").exists()


def test_export_publish_pdfs_includes_standings_and_club_report_pdf(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    # Create published PDF files in the normal publish folders.
    (output_paths.publish_pdf_league_updates_dir / "Season Update.pdf").write_text("placeholder", encoding="utf-8")
    (output_paths.publish_pdf_race_cards_dir / "Race 1 - Scoring Card.pdf").write_text("placeholder", encoding="utf-8")

    # Create the club report PDF under the club-reports docx folder.
    club_pdf = output_paths.publish_docx_club_reports_dir / "club_reports_1999.pdf"
    club_pdf.write_text("placeholder", encoding="utf-8")

    # Create the season standings workbook.
    standings = output_paths.publish_standings_dir / "Season Standings R01 1999.xlsx"
    standings.write_text("placeholder", encoding="utf-8")

    export_dir = tmp_path / "exported_pdfs"
    result_path = export_publish_pdfs(output_dir, export_dir, flatten=True)

    assert result_path == export_dir
    assert (export_dir / "Season Update.pdf").exists()
    assert (export_dir / "Race 1 - Scoring Card.pdf").exists()
    assert (export_dir / "club_reports_1999.pdf").exists()
    assert (export_dir / "Season Standings R01 1999.xlsx").exists()


def test_publish_results_can_export_pdfs_to_a_single_folder(tmp_path: Path) -> None:
    data_root = tmp_path / "data"
    year = 1999
    year_root = data_root / str(year)
    audited_dir = year_root / "inputs" / "audited"
    audited_dir.mkdir(parents=True, exist_ok=True)
    (audited_dir / "Race 1 - audited.xlsx").write_text("placeholder", encoding="utf-8")

    output_dir = year_root / "outputs"
    output_paths = ensure_output_subdirs(output_dir)
    output_paths.publish_pdf_league_updates_dir.mkdir(parents=True, exist_ok=True)
    output_paths.publish_pdf_race_cards_dir.mkdir(parents=True, exist_ok=True)
    (output_paths.publish_pdf_league_updates_dir / "Season Update.pdf").write_text("placeholder", encoding="utf-8")
    (output_paths.publish_pdf_race_cards_dir / "Race 1 - Scoring Card.pdf").write_text("placeholder", encoding="utf-8")

    report_dir = tmp_path / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    export_pdf_dir = tmp_path / "exported_pdfs"

    result = publish.publish_results(
        year=year,
        data_root=data_root,
        report_dir=report_dir,
        export_pdf_dir=export_pdf_dir,
    )

    assert result == 0
    assert (export_pdf_dir / "Season Update.pdf").exists()
    assert (export_pdf_dir / "Race 1 - Scoring Card.pdf").exists()
    report_path = report_dir / f"year-{year}"
    assert (report_path / "publish_results.json").exists()


def test_publish_package_artifacts_preserves_nested_structure_when_not_flattened(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    sample_files = [
        output_paths.publish_docx_league_updates_dir / "Season Update.docx",
        output_paths.publish_docx_race_cards_dir / "Race 1 - Scoring Card.docx",
        output_paths.publish_pdf_league_updates_dir / "Season Update.pdf",
        output_paths.publish_pdf_race_cards_dir / "Race 1 - Scoring Card.pdf",
        output_paths.publish_standings_dir / "Standings.xlsx",
        output_paths.publish_review_packs_dir / "Category Review.xlsx",
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
        package_dir / "pdf" / "Season Update.pdf",
        package_dir / "pdf" / "Race 1 - Scoring Card.pdf",
        package_dir / "standings" / "Standings.xlsx",
        package_dir / "review-packs" / "Category Review.xlsx",
        package_dir / "club-reports" / "club_reports_1999.docx",
    ]

    for expected_file in expected:
        assert expected_file.exists()


def test_package_publish_cli_handles_no_flatten_and_creates_package(tmp_path: Path) -> None:
    year_root = tmp_path / "1999"
    output_dir = year_root / "outputs"
    output_paths = ensure_output_subdirs(output_dir)

    sample_files = [
        output_paths.publish_docx_league_updates_dir / "Season Update.docx",
        output_paths.publish_pdf_league_updates_dir / "Season Update.pdf",
    ]
    for sample in sample_files:
        sample.write_text("placeholder", encoding="utf-8")

    package_script = Path(__file__).resolve().parents[2] / "scripts" / "publish" / "package_publish.py"
    result = subprocess.run(
        [
            sys.executable,
            str(package_script),
            "--year",
            "1999",
            "--data-root",
            str(tmp_path),
            "--no-flatten",
        ],
        cwd=str(Path(__file__).resolve().parents[2]),
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
    package_dir = output_dir / "publish" / "package"
    assert package_dir.exists()
    assert (package_dir / "docx" / "Season Update.docx").exists()
    assert (package_dir / "pdf" / "Season Update.pdf").exists()


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
    error_json = report_path / "publish_results.json"
    assert error_json.exists()

    payload = json.loads(error_json.read_text(encoding="utf-8"))
    assert payload["success"] is False
    assert payload["error"]["message"] == "Publish failed: simulated publish failure"
    assert "RuntimeError: simulated publish failure" in payload["error"]["traceback"]
