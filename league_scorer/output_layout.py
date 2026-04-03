"""Helpers for the structured season output folder layout and naming."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil


@dataclass(frozen=True)
class OutputPaths:
    output_dir: Path
    publish_dir: Path
    publish_docx_race_cards_dir: Path
    publish_docx_league_updates_dir: Path
    publish_pdf_race_cards_dir: Path
    publish_pdf_league_updates_dir: Path
    publish_xlsx_standings_dir: Path
    publish_xlsx_review_packs_dir: Path
    audit_workbooks_dir: Path
    audit_manual_changes_dir: Path
    quality_data_dir: Path
    quality_staged_checks_dir: Path
    autopilot_runs_dir: Path
    logs_dir: Path
    manifests_dir: Path


@dataclass(frozen=True)
class OutputSortResult:
    moved_count: int
    skipped_count: int
    moved_files: dict[str, str]


def build_output_paths(output_dir: Path) -> OutputPaths:
    output_dir = Path(output_dir)
    publish_dir = output_dir / "publish"
    quality_dir = output_dir / "quality"
    audit_dir = output_dir / "audit"
    return OutputPaths(
        output_dir=output_dir,
        publish_dir=publish_dir,
        publish_docx_race_cards_dir=publish_dir / "docx" / "race-cards",
        publish_docx_league_updates_dir=publish_dir / "docx" / "league-updates",
        publish_pdf_race_cards_dir=publish_dir / "pdf" / "race-cards",
        publish_pdf_league_updates_dir=publish_dir / "pdf" / "league-updates",
        publish_xlsx_standings_dir=publish_dir / "xlsx" / "standings",
        publish_xlsx_review_packs_dir=publish_dir / "xlsx" / "review-packs",
        audit_workbooks_dir=audit_dir / "workbooks",
        audit_manual_changes_dir=audit_dir / "manual-changes",
        quality_data_dir=quality_dir / "data-quality",
        quality_staged_checks_dir=quality_dir / "staged-checks",
        autopilot_runs_dir=output_dir / "autopilot" / "runs",
        logs_dir=output_dir / "logs",
        manifests_dir=output_dir / "manifests",
    )


def ensure_output_subdirs(output_dir: Path) -> OutputPaths:
    paths = build_output_paths(output_dir)
    paths.output_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_docx_race_cards_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_docx_league_updates_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_pdf_race_cards_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_pdf_league_updates_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_xlsx_standings_dir.mkdir(parents=True, exist_ok=True)
    paths.publish_xlsx_review_packs_dir.mkdir(parents=True, exist_ok=True)
    paths.audit_workbooks_dir.mkdir(parents=True, exist_ok=True)
    paths.audit_manual_changes_dir.mkdir(parents=True, exist_ok=True)
    paths.quality_data_dir.mkdir(parents=True, exist_ok=True)
    paths.quality_staged_checks_dir.mkdir(parents=True, exist_ok=True)
    paths.autopilot_runs_dir.mkdir(parents=True, exist_ok=True)
    paths.logs_dir.mkdir(parents=True, exist_ok=True)
    paths.manifests_dir.mkdir(parents=True, exist_ok=True)
    return paths


def sort_existing_output_files(output_dir: Path) -> OutputSortResult:
    """Move legacy flat output files/folders into the structured output layout."""
    paths = ensure_output_subdirs(output_dir)
    moved_files: dict[str, str] = {}
    moved_count = 0
    skipped_count = 0

    # Move legacy top-level files.
    for item in sorted(paths.output_dir.iterdir()):
        if item.is_dir():
            continue
        destination_dir = _destination_for_output_file(item, paths)
        if destination_dir is None:
            skipped_count += 1
            continue
        destination = destination_dir / item.name
        if destination.exists():
            skipped_count += 1
            continue
        shutil.move(str(item), str(destination))
        moved_count += 1
        moved_files[item.name] = str(destination.relative_to(paths.output_dir))

    # Lift legacy audit files if present.
    legacy_audit_dir = paths.output_dir / "audit"
    if legacy_audit_dir.exists() and legacy_audit_dir.is_dir():
        for workbook in sorted(legacy_audit_dir.glob("*.xlsx")):
            if workbook.name.lower() == "manual_data_audit.xlsx":
                target = paths.audit_manual_changes_dir / workbook.name
            else:
                target = paths.audit_workbooks_dir / workbook.name
            if target.exists():
                skipped_count += 1
                continue
            shutil.move(str(workbook), str(target))
            moved_count += 1
            moved_files[workbook.name] = str(target.relative_to(paths.output_dir))

    # Lift legacy staged/data-quality folders if present.
    moved_count += _move_tree_contents(paths.output_dir / "staged-checks", paths.quality_staged_checks_dir, moved_files)
    moved_count += _move_tree_contents(paths.output_dir / "data-quality", paths.quality_data_dir, moved_files)

    return OutputSortResult(moved_count=moved_count, skipped_count=skipped_count, moved_files=moved_files)


def standings_filename(highest_race: int, year: int) -> str:
    return f"Season Standings R{highest_race:02d} {year}.xlsx"


def category_review_filename(highest_race: int, year: int) -> str:
    return f"Season {year} - Category Review Through Race {highest_race}.xlsx"


def time_query_review_filename(highest_race: int, year: int) -> str:
    return f"Season {year} - Time Query Review Through Race {highest_race}.xlsx"


def league_update_basename(highest_race: int, year: int) -> str:
    return f"Season {year} - League Update Through Race {highest_race}"


def race_scoring_card_basename(race_num: int, race_name: str) -> str:
    race_name = re.sub(r"\s*\(audited\)\s*$", "", str(race_name), flags=re.IGNORECASE).strip()
    race_name = re.sub(r"^race\s*#?\s*\d+\s*[-–—]*\s*", "", race_name, flags=re.IGNORECASE).strip()
    clean_name = _sanitise_filename_part(race_name) or f"Race {race_num}"
    return f"Race {race_num:02d} - {clean_name} - Scoring Card"


def _sanitise_filename_part(value: str) -> str:
    cleaned = re.sub(r"[<>:\"/\\|?*]", " ", str(value))
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _destination_for_output_file(file_path: Path, paths: OutputPaths) -> Path | None:
    name = file_path.name.lower()
    suffix = file_path.suffix.lower()

    if suffix == ".xlsx":
        if "results" in name or "season standings" in name:
            return paths.publish_xlsx_standings_dir
        if "category" in name or "time qry" in name or "time query" in name:
            return paths.publish_xlsx_review_packs_dir
        if "audit" in name:
            return paths.audit_workbooks_dir
        return None

    if suffix == ".docx":
        if "race report" in name or "scoring card" in name:
            return paths.publish_docx_race_cards_dir
        if "league update" in name:
            return paths.publish_docx_league_updates_dir
        return None

    if suffix == ".pdf":
        if "race report" in name or "scoring card" in name:
            return paths.publish_pdf_race_cards_dir
        if "league update" in name:
            return paths.publish_pdf_league_updates_dir
        return None

    return None


def _move_tree_contents(source_dir: Path, target_dir: Path, moved_files: dict[str, str]) -> int:
    if not source_dir.exists() or not source_dir.is_dir():
        return 0
    target_dir.mkdir(parents=True, exist_ok=True)
    moved = 0
    for item in sorted(source_dir.rglob("*")):
        if item.is_dir():
            continue
        relative = item.relative_to(source_dir)
        destination = target_dir / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.exists():
            continue
        shutil.move(str(item), str(destination))
        moved += 1
        moved_files[item.name] = str(destination)
    return moved
