"""Run staged validation checks for the WRRL operational chain.

Stages:
1. Raw ingest validation
2. Raw -> audited consolidation validation
3. Audit generation validation
4. Main scoring regression validation
"""

from __future__ import annotations

import argparse
import json
import hashlib
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from league_scorer.audit import LeagueAuditor
from league_scorer.common_files import race_discovery_exclusions
from league_scorer.main import LeagueScorer
from league_scorer.race_processor import extract_race_number
from league_scorer.race_validation import validate_race_schema
from league_scorer.source_loader import discover_race_files, load_race_dataframe


@dataclass
class StageResult:
    stage: int
    name: str
    status: str
    message: str
    details: dict[str, Any] = field(default_factory=dict)


def _resolve_data_root(explicit: Path | None) -> Path | None:
    if explicit is not None:
        return explicit

    prefs = Path.home() / ".wrrl_prefs.json"
    if prefs.exists():
        try:
            data = json.loads(prefs.read_text(encoding="utf-8"))
            value = data.get("data_root")
            if value:
                return Path(value)
        except Exception:
            return None
    return None


def _find_raw_dir(input_dir: Path) -> Path | None:
    candidates = [
        input_dir / "Raw Data",
        input_dir / "Raw",
        input_dir / "raw data",
        input_dir / "raw",
    ]
    for candidate in candidates:
        if candidate.exists() and candidate.is_dir():
            return candidate

    for child in input_dir.iterdir() if input_dir.exists() else []:
        if child.is_dir() and child.name.strip().lower() in {"raw data", "raw"}:
            return child
    return None


def _discover_raw_race_files(raw_dir: Path) -> dict[int, list[Path]]:
    by_race: dict[int, list[Path]] = {}
    for file in sorted(raw_dir.rglob("*.xlsx")):
        race_num = extract_race_number(file.stem)
        if race_num is None:
            continue
        by_race.setdefault(race_num, []).append(file)
    return by_race


def _workbook_fingerprint(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets: dict[str, Any] = {}
        for ws in wb.worksheets:
            digest = hashlib.sha256()
            row_count = 0
            col_count = 0
            for row in ws.iter_rows(values_only=True):
                row_count += 1
                col_count = max(col_count, len(row))
                joined = "|".join("" if v is None else str(v) for v in row)
                digest.update(joined.encode("utf-8", errors="ignore"))
                digest.update(b"\n")
            sheets[ws.title] = {
                "rows": row_count,
                "cols": col_count,
                "sha256": digest.hexdigest(),
            }
        return {
            "workbook_name": path.name,
            "sheets": sheets,
        }
    finally:
        wb.close()


def run_checks(args: argparse.Namespace) -> tuple[bool, list[StageResult]]:
    data_root = _resolve_data_root(args.data_root)
    if data_root is None:
        message = "No data root configured. Set --data-root or configure WRRL data root in app settings."
        if args.allow_missing_data:
            return True, [StageResult(0, "Preflight", "skipped", message)]
        return False, [StageResult(0, "Preflight", "failed", message)]

    input_dir = data_root / str(args.year) / "inputs"
    output_dir = data_root / str(args.year) / "outputs"

    if not input_dir.exists():
        message = f"Input directory not found: {input_dir}"
        if args.allow_missing_data:
            return True, [StageResult(0, "Preflight", "skipped", message)]
        return False, [StageResult(0, "Preflight", "failed", message)]

    results: list[StageResult] = []

    # Stage 1
    raw_dir = _find_raw_dir(input_dir)
    if raw_dir is None:
        results.append(
            StageResult(
                1,
                "Raw Ingest Validation",
                "failed",
                "Raw data folder not found under inputs.",
                {"input_dir": str(input_dir)},
            )
        )
        return False, results

    raw_races = _discover_raw_race_files(raw_dir)
    schema_ok = 0
    schema_fail = 0
    schema_warnings = 0
    stage1_errors: list[str] = []

    for race_num, files in sorted(raw_races.items()):
        for file in files:
            try:
                df = load_race_dataframe(file)
                validation = validate_race_schema(df, file)
                schema_ok += 1
                schema_warnings += len(validation.issues)
            except Exception as exc:
                schema_fail += 1
                stage1_errors.append(f"Race {race_num} {file.name}: {exc}")

    status = "passed" if schema_fail == 0 else "failed"
    results.append(
        StageResult(
            1,
            "Raw Ingest Validation",
            status,
            f"Validated {schema_ok + schema_fail} raw workbook(s); failures={schema_fail}, warnings={schema_warnings}.",
            {
                "raw_dir": str(raw_dir),
                "races_detected": sorted(raw_races),
                "errors": stage1_errors[:50],
            },
        )
    )
    if status == "failed":
        return False, results

    # Stage 2
    audited_race_files = discover_race_files(input_dir, excluded_names=race_discovery_exclusions())
    raw_race_nums = set(raw_races)
    audited_race_nums = set(audited_race_files)

    missing_from_audited = sorted(raw_race_nums - audited_race_nums)
    status = "passed" if not missing_from_audited else "failed"
    results.append(
        StageResult(
            2,
            "Raw To Audited Consolidation",
            status,
            f"Raw races={len(raw_race_nums)}, audited races={len(audited_race_nums)}.",
            {
                "missing_from_audited": missing_from_audited,
                "audited_files": {k: str(v) for k, v in audited_race_files.items()},
            },
        )
    )
    if status == "failed":
        return False, results

    # Stage 3
    auditor = LeagueAuditor(input_dir=input_dir, output_dir=output_dir, year=args.year)
    audit_path = auditor.run(race_files=audited_race_files)

    issue_count = sum(len(v) for v in auditor.all_race_issues.values())
    results.append(
        StageResult(
            3,
            "Audit Generation Validation",
            "passed",
            f"Audit workbook generated with {issue_count} issue row(s).",
            {
                "audit_workbook": str(audit_path),
                "issue_count": issue_count,
            },
        )
    )

    # Stage 4
    scorer = LeagueScorer(input_dir=input_dir, output_dir=output_dir, year=args.year)
    warnings = scorer.run(race_files=audited_race_files)

    from league_scorer.graphical.results_workbook import find_latest_results_workbook

    latest_results = find_latest_results_workbook(output_dir)
    if latest_results is None:
        results.append(
            StageResult(4, "Main Scoring Regression", "failed", "No results workbook was generated.")
        )
        return False, results

    fingerprint = _workbook_fingerprint(latest_results)
    baseline_path = Path(args.baseline_file)
    baseline_exists = baseline_path.exists()

    if args.write_baseline or not baseline_exists:
        baseline_path.parent.mkdir(parents=True, exist_ok=True)
        baseline_path.write_text(json.dumps(fingerprint, indent=2), encoding="utf-8")
        results.append(
            StageResult(
                4,
                "Main Scoring Regression",
                "passed",
                "Baseline written from current results workbook.",
                {
                    "baseline": str(baseline_path),
                    "warnings": warnings,
                    "results_workbook": str(latest_results),
                },
            )
        )
        return True, results

    expected = json.loads(baseline_path.read_text(encoding="utf-8"))
    status = "passed" if expected.get("sheets") == fingerprint.get("sheets") else "failed"
    message = "Results fingerprint matches baseline." if status == "passed" else "Results fingerprint differs from baseline."

    details = {
        "baseline": str(baseline_path),
        "results_workbook": str(latest_results),
        "warnings": warnings,
    }

    if status == "failed":
        expected_sheets = expected.get("sheets", {})
        current_sheets = fingerprint.get("sheets", {})
        changed = sorted(
            set(expected_sheets) | set(current_sheets)
        )
        diffs: list[dict[str, Any]] = []
        for sheet in changed:
            if expected_sheets.get(sheet) != current_sheets.get(sheet):
                diffs.append(
                    {
                        "sheet": sheet,
                        "expected": expected_sheets.get(sheet),
                        "current": current_sheets.get(sheet),
                    }
                )
        details["diffs"] = diffs[:50]

    results.append(StageResult(4, "Main Scoring Regression", status, message, details))
    return status == "passed", results


def write_report(results: list[StageResult], report_dir: Path, success: bool) -> None:
    report_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "success": success,
        "stages": [asdict(item) for item in results],
    }
    (report_dir / "staged_checks_report.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")

    lines = [
        "# WRRL Staged Checks Report",
        "",
        f"Generated: {payload['generated_at']}",
        f"Success: {success}",
        "",
    ]
    for item in results:
        lines.append(f"## Stage {item.stage} - {item.name}")
        lines.append(f"Status: {item.status}")
        lines.append("")
        lines.append(item.message)
        lines.append("")
    (report_dir / "staged_checks_report.md").write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run WRRL staged operational checks.")
    parser.add_argument("--year", type=int, default=1999)
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument(
        "--report-dir",
        type=Path,
        default=Path("output") / "staged-checks",
    )
    parser.add_argument(
        "--baseline-file",
        type=Path,
        default=Path("tests") / "baselines" / "season_1999_results_baseline.json",
    )
    parser.add_argument("--write-baseline", action="store_true")
    parser.add_argument("--allow-missing-data", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    success, results = run_checks(args)
    write_report(results, args.report_dir, success)

    for item in results:
        print(f"[stage {item.stage}] {item.name}: {item.status} - {item.message}")

    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())
