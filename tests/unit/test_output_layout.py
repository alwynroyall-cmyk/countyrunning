from pathlib import Path

from league_scorer.output_layout import (
    OutputPaths,
    OutputSortResult,
    build_output_paths,
    ensure_output_subdirs,
    sort_existing_output_files,
)


def test_build_output_paths_returns_expected_directories(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    paths = build_output_paths(output_dir)

    assert isinstance(paths, OutputPaths)
    assert paths.output_dir == output_dir
    assert paths.publish_dir == output_dir / "publish"
    assert paths.publish_docx_dir == output_dir / "publish" / "docx"
    assert paths.publish_docx_club_reports_dir == output_dir / "publish" / "docx" / "club-reports"
    assert paths.audit_workbooks_dir == output_dir / "audit" / "workbooks"
    assert paths.audit_manual_changes_dir == output_dir / "audit" / "manual-changes"
    assert paths.quality_data_dir == output_dir / "quality" / "data-quality"
    assert paths.autopilot_runs_dir == output_dir / "autopilot" / "runs"


def test_ensure_output_subdirs_creates_all_directories(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    paths = ensure_output_subdirs(output_dir)

    assert output_dir.exists()
    assert paths.publish_dir.exists()
    assert paths.publish_docx_club_reports_dir.exists()
    assert paths.audit_workbooks_dir.exists()
    assert paths.audit_manual_changes_dir.exists()
    assert paths.quality_data_dir.exists()
    assert paths.quality_staged_checks_dir.exists()
    assert paths.autopilot_runs_dir.exists()
    assert paths.logs_dir.exists()
    assert paths.manifests_dir.exists()


def test_sort_existing_output_files_moves_legacy_files(tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_dir.mkdir(parents=True)
    (output_dir / "old_file.xlsx").write_text("legacy", encoding="utf-8")
    (output_dir / "legacy_audit.xlsx").write_text("legacy", encoding="utf-8")
    legacy_publish_xlsx = output_dir / "publish" / "xlsx"
    (legacy_publish_xlsx / "standings").mkdir(parents=True, exist_ok=True)
    (legacy_publish_xlsx / "review-packs").mkdir(parents=True, exist_ok=True)
    (legacy_publish_xlsx / "standings" / "Standings.xlsx").write_text("placeholder", encoding="utf-8")
    (legacy_publish_xlsx / "review-packs" / "Category Review.xlsx").write_text("placeholder", encoding="utf-8")

    result = sort_existing_output_files(output_dir)

    assert isinstance(result, OutputSortResult)
    assert result.moved_count >= 2
    assert (output_dir / "audit" / "workbooks" / "legacy_audit.xlsx").exists()
    assert (output_dir / "publish" / "standings" / "Standings.xlsx").exists()
    assert (output_dir / "publish" / "review-packs" / "Category Review.xlsx").exists()
