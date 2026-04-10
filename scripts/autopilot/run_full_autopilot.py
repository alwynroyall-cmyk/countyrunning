"""Run WRRL Admin Suite autopilot workflow for one season.

Autopilot flow:
1. Resolve season data root and input/output paths.
2. Run audit and load Actionable Issues.
3. Optionally apply safe auto-fixes for issues that do not require operator input.
4. Re-run audit to refresh actionable issue state.
5. Run staged checks and quality gate.
6. Write consolidated autopilot JSON/Markdown report.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from league_scorer.autopilot.audit import LeagueAuditor
from league_scorer.autopilot.audit_cleanser import create_cleansed_race_file
from league_scorer.autopilot.audit_data_service import ACTIONABLE_COLUMNS, load_actionable_issues
from league_scorer.autopilot.archive_service import ensure_archived_in_inputs
from league_scorer.input.club_loader import load_clubs
from league_scorer.input.common_files import race_discovery_exclusions
from league_scorer.input.input_layout import build_input_paths, ensure_input_subdirs, sort_existing_input_files
from league_scorer.output.output_layout import ensure_output_subdirs
from league_scorer.autopilot.issue_resolution_service import (
    apply_quick_fix_for_issue,
    quick_fix_requires_input,
    supports_quick_fix,
)
from league_scorer.process.race_processor import extract_race_number
from league_scorer.input.source_loader import discover_race_files
import shutil
from scripts.autopilot import run_staged_checks as staged_checks
from league_scorer.autopilot.series_consolidation import consolidate_series_files
import re


@dataclass
class AuditSnapshot:
    workbook: str
    race_file_count: int
    actionable_count: int
    issue_codes: list[str]


@dataclass
class FixSummary:
    attempted: int
    applied: int
    rows_updated: int
    files_touched: int
    skipped_unsupported: int
    skipped_requires_input: int
    failed: int
    failure_examples: list[str]


def _write_failure_report(
    *,
    report_dir: Path,
    year: int,
    generated_at: str,
    mode: str,
    reason: str,
    data_root: Path | None,
    input_dir: Path | None,
    output_dir: Path | None,
) -> tuple[Path, Path]:
    """Write a minimal autopilot report payload for early failures."""
    report_root = report_dir / f"year-{year}"
    report_root.mkdir(parents=True, exist_ok=True)

    payload: dict[str, Any] = {
        "generated_at": generated_at,
        "success": False,
        "settings": {
            "year": year,
            "mode": mode,
            "data_root": str(data_root) if data_root is not None else "",
            "input_dir": str(input_dir) if input_dir is not None else "",
            "output_dir": str(output_dir) if output_dir is not None else "",
        },
        "error": {
            "stage": "preflight",
            "message": reason,
        },
        "audit_before": {
            "workbook": "",
            "race_file_count": 0,
            "actionable_count": 0,
            "issue_codes": [],
        },
        "audit_after": {
            "workbook": "",
            "race_file_count": 0,
            "actionable_count": 0,
            "issue_codes": [],
        },
        "fix_summary": {
            "attempted": 0,
            "applied": 0,
            "rows_updated": 0,
            "files_touched": 0,
            "skipped_unsupported": 0,
            "skipped_requires_input": 0,
            "failed": 0,
            "failure_examples": [],
        },
        "staged_checks": {
            "success": False,
            "report_dir": "",
            "results": [],
        },
        "strict_condition": {
            "required": mode == "strict",
            "passed": False,
            "remaining_actionable_issues": 0,
        },
    }

    json_path = report_root / "autopilot_report.json"
    md_path = report_root / "autopilot_report.md"
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    md_path.write_text(_to_markdown(payload), encoding="utf-8")
    return json_path, md_path


def _resolve_data_root(explicit: Path | None) -> Path | None:
    # Reuse staged checks resolution to stay consistent with existing tooling.
    return staged_checks._resolve_data_root(explicit)


def _season_paths(data_root: Path, year: int) -> tuple[Path, Path]:
    year_root = data_root / str(year)
    return year_root / "inputs", year_root / "outputs"


def _run_audit_snapshot(
    *, input_dir: Path, output_dir: Path, year: int,
    race_files: list | None = None,
) -> AuditSnapshot:
    if race_files is None:
        audited_dir = build_input_paths(input_dir).audited_dir
        race_files = discover_race_files(audited_dir, excluded_names=race_discovery_exclusions())
    if not race_files:
        raise RuntimeError(
            f"No audited race files available for audit snapshot in '{input_dir}'. "
            "Generate audited files from raw_data first."
        )
    auditor = LeagueAuditor(input_dir=input_dir, output_dir=output_dir, year=year)
    workbook = auditor.run(race_files=race_files)

    actionable_df = load_actionable_issues(workbook)
    issue_codes = sorted(
        {
            str(value).strip()
            for value in actionable_df.get("Issue Code", [])
            if str(value).strip()
        }
    )
    return AuditSnapshot(
        workbook=str(workbook),
        race_file_count=len(race_files),
        actionable_count=len(actionable_df),
        issue_codes=issue_codes,
    )


def _attempt_safe_fixes(
    actionable_workbook: Path,
    *,
    input_dir: Path,
    max_attempts: int | None,
) -> FixSummary:
    actionable_df = load_actionable_issues(actionable_workbook)
    attempted = 0
    applied = 0
    rows_updated = 0
    files_touched = 0
    skipped_unsupported = 0
    skipped_requires_input = 0
    failed = 0
    failure_examples: list[str] = []

    for _, row in actionable_df.iterrows():
        issue = {col: str(row.get(col, "") or "").strip() for col in ACTIONABLE_COLUMNS}
        code = issue.get("Issue Code", "")

        if not supports_quick_fix(code):
            skipped_unsupported += 1
            continue

        if quick_fix_requires_input(code):
            skipped_requires_input += 1
            continue

        if max_attempts is not None and attempted >= max_attempts:
            break

        attempted += 1
        result = apply_quick_fix_for_issue(issue, input_dir=input_dir, target_value=None)
        if result.success:
            applied += 1
            rows_updated += int(result.updated_rows)
            files_touched += int(result.touched_files)
        else:
            failed += 1
            if len(failure_examples) < 10:
                failure_examples.append(f"{code}: {result.message}")

    return FixSummary(
        attempted=attempted,
        applied=applied,
        rows_updated=rows_updated,
        files_touched=files_touched,
        skipped_unsupported=skipped_unsupported,
        skipped_requires_input=skipped_requires_input,
        failed=failed,
        failure_examples=failure_examples,
    )


def _race_names_for_progress(race_files: dict[int, Path]) -> list[str]:
    """Return ordered race file names for GUI progress updates."""
    return [race_files[race_num].name for race_num in sorted(race_files)]


def _generate_audited_race_files(input_dir: Path, *, overwrite_existing: bool) -> dict[int, Path]:
    """Generate audited race files from raw_data and return {race_num: audited_path}."""
    paths = build_input_paths(input_dir)

    # Clear out existing audited files to avoid false-negative stale issues
    audited_dir = paths.audited_dir
    if audited_dir.exists():
        for f in sorted(audited_dir.iterdir()):
            if f.is_file() and f.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
                try:
                    f.unlink()
                except Exception:
                    # ignore failures to delete; will be reported later if they cause errors
                    pass

    # Remove any residual series round workbooks that ended up in raw_data (these
    # should not be present; consolidated workbooks belong in raw_data instead).
    series_pattern = re.compile(r"\bseries\b\s*#\s*\d+", re.IGNORECASE)
    raw_data_dir = paths.raw_data_dir
    if raw_data_dir.exists():
        for f in sorted(raw_data_dir.iterdir()):
            if not f.is_file():
                continue
            name = f.stem
            if series_pattern.search(name):
                try:
                    f.unlink()
                except Exception:
                    pass

    # Discover candidate source files. Prefer series files (which may contain
    # operator edits) over raw_data files when both exist for the same race.
    raw_race_files = discover_race_files(paths.raw_data_dir, excluded_names=race_discovery_exclusions())
    series_race_files = discover_race_files(paths.series_dir, excluded_names=race_discovery_exclusions())

    # Merge: series files take precedence over raw_data for the same race number.
    if not raw_race_files and not series_race_files:
        return {}

    # Start with raw files then overlay series files
    merged_sources: dict[int, Path] = dict(raw_race_files)

    # Build merged_sources by preferring consolidated/raw sources in this order:
    # 1) If there are multiple series round files for a race, consolidate them
    #    into a single consolidated workbook in `raw_data` and use that.
    # 2) If a single series file exists for the race, prefer that.
    # 3) Otherwise use the raw_data file if present.
    series_pattern = re.compile(r"^race\s*#?\s*(\d+)\b", re.IGNORECASE)
    all_numbers = set(list(raw_race_files.keys()) + list(series_race_files.keys()))
    for num in sorted(all_numbers):
        # gather candidate series files for this race number
        candidates: list[Path] = []
        if paths.series_dir.exists():
            for p in sorted(paths.series_dir.iterdir()):
                if not p.is_file() or p.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
                    continue
                m = series_pattern.match(p.stem.strip())
                if not m:
                    continue
                try:
                    rn = int(m.group(1))
                except Exception:
                    continue
                if rn == num:
                    candidates.append(p)

        if len(candidates) >= 2:
            # Consolidate multi-round series into raw_data and prefer that file
            try:
                result = consolidate_series_files(candidates, series_dir=paths.series_dir, raw_data_dir=paths.raw_data_dir)
                merged_sources[num] = result.consolidated_path
                continue
            except Exception:
                # fall through to try single series or raw
                pass

        # single series file present?
        if num in series_race_files:
            merged_sources[num] = series_race_files[num]
        elif num in raw_race_files:
            merged_sources[num] = raw_race_files[num]

    raw_to_preferred, club_info = load_clubs(paths.control_dir / "clubs.xlsx")
    preferred_clubs = sorted(club_info)

    audited_race_files: dict[int, Path] = {}
    for race_num in sorted(merged_sources):
        raw_path = merged_sources[race_num]
        ensure_archived_in_inputs(raw_path, input_dir)
        audited_path = create_cleansed_race_file(
            raw_path,
            raw_to_preferred,
            preferred_clubs,
            audited_dir=paths.audited_dir,
            control_dir=paths.control_dir,
            overwrite_existing=overwrite_existing,
        )
        discovered_num = extract_race_number(audited_path.stem)
        if discovered_num is None:
            raise RuntimeError(
                f"Generated audited file does not contain a valid race number: {audited_path.name}"
            )
        audited_race_files[discovered_num] = audited_path

    return audited_race_files


def _build_stage_args(args: argparse.Namespace, data_root: Path) -> SimpleNamespace:
    return SimpleNamespace(
        year=args.year,
        data_root=data_root,
        report_dir=args.staged_report_dir,
        baseline_file=args.baseline_file,
        write_baseline=args.write_baseline,
        quality_gate_threshold=args.quality_gate_threshold,
        data_quality_output_dir=args.data_quality_output_dir,
        allow_missing_data=args.allow_missing_data,
    )


def _print_preflight_summary(input_dir: Path) -> None:
    """Print a concise preflight summary of structured input folders."""
    paths = build_input_paths(input_dir)

    raw_files = [
        p
        for p in sorted(paths.raw_data_dir.iterdir())
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
    ] if paths.raw_data_dir.exists() else []
    series_files = [
        p
        for p in sorted(paths.series_dir.iterdir())
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
    ] if paths.series_dir.exists() else []
    audited_files = [
        p
        for p in sorted(paths.audited_dir.iterdir())
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
    ] if paths.audited_dir.exists() else []

    raw_races = discover_race_files(paths.raw_data_dir, excluded_names=race_discovery_exclusions())
    audited_races = discover_race_files(paths.audited_dir, excluded_names=race_discovery_exclusions())

    clubs_ok = (paths.control_dir / "clubs.xlsx").exists()
    events_ok = (paths.control_dir / "wrrl_events.xlsx").exists()
    name_lookup_ok = (paths.control_dir / "name_corrections.xlsx").exists()

    print("PRECHECK:INPUT_DIR:" + str(paths.input_dir), flush=True)
    print(
        "PRECHECK:COUNTS:"
        f"raw_files={len(raw_files)};series_files={len(series_files)};audited_files={len(audited_files)};"
        f"raw_races={len(raw_races)};audited_races={len(audited_races)}",
        flush=True,
    )
    print(
        "PRECHECK:CONTROL:"
        f"clubs.xlsx={'yes' if clubs_ok else 'no'};"
        f"wrrl_events.xlsx={'yes' if events_ok else 'no'};"
        f"name_corrections.xlsx={'yes' if name_lookup_ok else 'no'}",
        flush=True,
    )


def _to_markdown(payload: dict[str, Any]) -> str:
    settings = payload["settings"]
    before = payload["audit_before"]
    after = payload["audit_after"]
    fixes = payload["fix_summary"]

    lines = [
        "# WRRL Admin Suite Autopilot Report",
        "",
        f"Generated: {payload['generated_at']}",
        f"Year: {settings['year']}",
        f"Mode: {settings['mode']}",
        f"Overall Success: {payload['success']}",
        "",
        "## Audit Snapshot",
        "",
        f"- Before actionable issues: {before['actionable_count']}",
        f"- After actionable issues: {after['actionable_count']}",
        f"- Race files discovered: {after['race_file_count']}",
        "",
        "## Safe Fix Pass",
        "",
        f"- Attempts: {fixes['attempted']}",
        f"- Applied: {fixes['applied']}",
        f"- Rows updated: {fixes['rows_updated']}",
        f"- Files touched: {fixes['files_touched']}",
        f"- Skipped (unsupported): {fixes['skipped_unsupported']}",
        f"- Skipped (requires input): {fixes['skipped_requires_input']}",
        f"- Failed: {fixes['failed']}",
        "",
    ]

    error = payload.get("error")
    if isinstance(error, dict) and error.get("message"):
        lines.extend([
            "## Error",
            "",
            f"- Stage: {error.get('stage', '')}",
            f"- Message: {error.get('message', '')}",
            "",
        ])

    lines.extend([
        "## Staged Checks",
        "",
    ])

    for stage in payload["staged_checks"]["results"]:
        lines.append(f"- Stage {stage['stage']} {stage['name']}: {stage['status']} - {stage['message']}")

    if fixes["failure_examples"]:
        lines.extend([
            "",
            "## Fix Failures",
            "",
        ])
        for item in fixes["failure_examples"]:
            lines.append(f"- {item}")

    lines.append("")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run WRRL Admin Suite autopilot for one season."
    )
    parser.add_argument("--year", type=int, default=1999)
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument(
        "--mode",
        choices=("dry-run", "apply-safe-fixes", "strict"),
        default="apply-safe-fixes",
        help="dry-run: no fix attempts; apply-safe-fixes: apply safe no-input quick fixes; strict: apply safe fixes and fail if actionable issues remain.",
    )
    parser.add_argument(
        "--report-dir",
        type=Path,
        default=Path("output") / "autopilot" / "runs",
        help="Folder for autopilot JSON/Markdown output.",
    )
    parser.add_argument(
        "--staged-report-dir",
        type=Path,
        default=Path("output") / "quality" / "staged-checks",
        help="Folder for staged-check output generated during autopilot.",
    )
    parser.add_argument(
        "--data-quality-output-dir",
        type=Path,
        default=Path("output") / "quality" / "data-quality",
    )
    parser.add_argument(
        "--quality-gate-threshold",
        type=float,
        default=80.0,
    )
    parser.add_argument(
        "--baseline-file",
        type=Path,
        default=Path("tests") / "baselines" / "season_1999_results_baseline.json",
    )
    parser.add_argument("--write-baseline", action="store_true")
    parser.add_argument("--allow-missing-data", action="store_true")
    parser.add_argument(
        "--max-fix-attempts",
        type=int,
        default=0,
        help="Optional cap on safe-fix attempts (0 means no cap).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    generated_at = datetime.now(timezone.utc).isoformat()
    data_root = _resolve_data_root(args.data_root)
    if data_root is None:
        reason = "No data root configured. Set --data-root or configure WRRL data root in app settings."
        json_path, md_path = _write_failure_report(
            report_dir=args.report_dir,
            year=args.year,
            generated_at=generated_at,
            mode=args.mode,
            reason=reason,
            data_root=None,
            input_dir=None,
            output_dir=None,
        )
        print(reason)
        print(f"Wrote: {json_path}")
        print(f"Wrote: {md_path}")
        return 1

    input_dir, output_dir = _season_paths(data_root, args.year)
    if not input_dir.exists():
        reason = f"Input directory not found: {input_dir}"
        json_path, md_path = _write_failure_report(
            report_dir=args.report_dir,
            year=args.year,
            generated_at=generated_at,
            mode=args.mode,
            reason=reason,
            data_root=data_root,
            input_dir=input_dir,
            output_dir=output_dir,
        )
        print(reason)
        print(f"Wrote: {json_path}")
        print(f"Wrote: {md_path}")
        return 1

    ensure_input_subdirs(input_dir)
    sort_existing_input_files(input_dir)
    ensure_output_subdirs(output_dir)
    _print_preflight_summary(input_dir)

    # Build audited files from raw_data so fresh seasons can run end-to-end.
    paths = build_input_paths(input_dir)
    archive_before = set(p.name for p in paths.raw_data_archive_dir.glob("*")) if paths.raw_data_archive_dir.exists() else set()
    race_files = _generate_audited_race_files(input_dir, overwrite_existing=True)
    archive_after = set(p.name for p in paths.raw_data_archive_dir.glob("*")) if paths.raw_data_archive_dir.exists() else set()
    archive_added = max(0, len(archive_after) - len(archive_before))
    print(
        f"INFO: raw_data archive updated: added={archive_added}; total={len(archive_after)}",
        flush=True,
    )
    if not race_files:
        existing_audited = discover_race_files(paths.audited_dir, excluded_names=race_discovery_exclusions())
        if existing_audited:
            print(
                "INFO: Existing audited files were detected but are intentionally ignored "
                "until rebuilt from raw_data in this run.",
                flush=True,
            )
            print(f"INFO: Existing audited race count: {len(existing_audited)}", flush=True)
        reason = (
            "No raw race files found in raw_data. Place race workbooks in inputs/raw_data and re-run. "
            "Existing audited files are intentionally ignored to preserve provenance."
        )
        print(f"ERROR: {reason}", flush=True)
        print(f"ERROR: raw_data path: {paths.raw_data_dir}", flush=True)
        json_path, md_path = _write_failure_report(
            report_dir=args.report_dir,
            year=args.year,
            generated_at=generated_at,
            mode=args.mode,
            reason=f"{reason} raw_data path: {paths.raw_data_dir}",
            data_root=data_root,
            input_dir=input_dir,
            output_dir=output_dir,
        )
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    # Emit names for live progress tracking in the GUI.
    print(f"PROGRESS:RACES:{'|'.join(_race_names_for_progress(race_files))}", flush=True)

    print("PROGRESS:STAGE:1:Auditing season races", flush=True)
    audit_before = _run_audit_snapshot(
        input_dir=input_dir, output_dir=output_dir, year=args.year, race_files=race_files,
    )
    print("PROGRESS:STAGE_DONE:1", flush=True)

    fix_summary = FixSummary(
        attempted=0,
        applied=0,
        rows_updated=0,
        files_touched=0,
        skipped_unsupported=0,
        skipped_requires_input=0,
        failed=0,
        failure_examples=[],
    )

    if args.mode in {"apply-safe-fixes", "strict"}:
        print("PROGRESS:STAGE:2:Applying safe fixes", flush=True)
        cap = args.max_fix_attempts if args.max_fix_attempts > 0 else None
        fix_summary = _attempt_safe_fixes(
            Path(audit_before.workbook),
            input_dir=input_dir,
            max_attempts=cap,
        )
        print("PROGRESS:STAGE_DONE:2", flush=True)
    else:
        print("PROGRESS:STAGE:2:Skipping fixes (dry-run)", flush=True)
        print("PROGRESS:STAGE_DONE:2", flush=True)

    # Rebuild audited files after potential raw_data edits from quick fixes.
    race_files = _generate_audited_race_files(input_dir, overwrite_existing=True)

    # Always refresh actionable state after optional fix pass.
    audit_after = _run_audit_snapshot(
        input_dir=input_dir, output_dir=output_dir, year=args.year, race_files=race_files,
    )

    print("PROGRESS:STAGE:3:Running quality checks and scoring regression", flush=True)
    stage_args = _build_stage_args(args, data_root)
    previous_disable_pdf = os.environ.get("WRRL_DISABLE_PDF")
    os.environ["WRRL_DISABLE_PDF"] = "1"
    try:
        staged_success, staged_results = staged_checks.run_checks(
            stage_args,
            progress_cb=lambda msg: print(f"PROGRESS:SUBSTEP:3:{msg}", flush=True),
        )
    finally:
        if previous_disable_pdf is None:
            os.environ.pop("WRRL_DISABLE_PDF", None)
        else:
            os.environ["WRRL_DISABLE_PDF"] = previous_disable_pdf
    staged_checks.write_report(staged_results, args.staged_report_dir, staged_success)
    print("PROGRESS:STAGE_DONE:3", flush=True)

    strict_ok = True
    if args.mode == "strict" and audit_after.actionable_count > 0:
        strict_ok = False

    success = staged_success and strict_ok

    report_root = args.report_dir / f"year-{args.year}"
    report_root.mkdir(parents=True, exist_ok=True)

    payload: dict[str, Any] = {
        "generated_at": generated_at,
        "success": success,
        "settings": {
            "year": args.year,
            "mode": args.mode,
            "data_root": str(data_root),
            "input_dir": str(input_dir),
            "output_dir": str(output_dir),
            "quality_gate_threshold": args.quality_gate_threshold,
            "max_fix_attempts": args.max_fix_attempts,
        },
        "audit_before": asdict(audit_before),
        "audit_after": asdict(audit_after),
        "fix_summary": asdict(fix_summary),
        "staged_checks": {
            "success": staged_success,
            "report_dir": str(args.staged_report_dir),
            "results": [asdict(item) for item in staged_results],
        },
        "strict_condition": {
            "required": args.mode == "strict",
            "passed": strict_ok,
            "remaining_actionable_issues": audit_after.actionable_count,
        },
    }

    json_path = report_root / "autopilot_report.json"
    md_path = report_root / "autopilot_report.md"
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    md_path.write_text(_to_markdown(payload), encoding="utf-8")

    print(f"Autopilot success: {success}")
    print(f"Audit actionable issues before/after: {audit_before.actionable_count} -> {audit_after.actionable_count}")
    print(f"Safe fixes applied: {fix_summary.applied}")
    print(f"Wrote: {json_path}")
    print(f"Wrote: {md_path}")

    # Clear dirty flag so UI knows data are aligned
    try:
        dirty_flag = Path(output_dir) / "autopilot" / "dirty"
        if dirty_flag.exists():
            dirty_flag.unlink()
            print(f"INFO: Cleared dirty flag: {dirty_flag}", flush=True)
        # also clear RAES-specific dirty marker if present
        try:
            raes_flag = Path(output_dir) / "raes" / "dirty"
            if raes_flag.exists():
                raes_flag.unlink()
                print(f"INFO: Cleared RAES dirty flag: {raes_flag}", flush=True)
        except Exception:
            pass
    except Exception:
        pass

    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())
