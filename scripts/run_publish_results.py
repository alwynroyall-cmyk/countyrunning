"""Publish final WRRL results from audited files.

Flow:
1. Resolve season paths.
2. Discover audited race files.
3. Run scoring pipeline from audited files with PDF generation enabled.
4. Write a short publish summary report.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from league_scorer.main import LeagueScorer
from league_scorer.common_files import race_discovery_exclusions
from league_scorer.input_layout import build_input_paths
from league_scorer.output_layout import build_output_paths, ensure_output_subdirs, sort_existing_output_files
from league_scorer.source_loader import discover_race_files
from scripts.run_full_autopilot import _race_names_for_progress, _resolve_data_root, _season_paths


def _to_markdown(payload: dict[str, Any]) -> str:
    lines = [
        "# WRRL League AI Publish Results",
        "",
        f"Generated: {payload['generated_at']}",
        f"Year: {payload['settings']['year']}",
        f"Success: {payload['success']}",
        "",
        "## Summary",
        "",
        f"- Audited race files published: {payload['summary']['audited_race_count']}",
        f"- Results workbook: {payload['summary']['results_workbook'] or 'Not written'}",
        f"- Warnings: {payload['summary']['warning_count']}",
        "",
        "## Notes",
        "",
        "- This path publishes from audited files and includes PDF outputs.",
    ]

    warnings = payload['summary'].get('warnings', [])
    if warnings:
        lines.extend(["", "## Pipeline Warnings", ""])
        for item in warnings:
            lines.append(f"- {item}")

    error = payload.get('error')
    if isinstance(error, dict) and error.get('message'):
        lines.extend(["", "## Error", "", f"- {error['message']}"])

    lines.append("")
    return "\n".join(lines)


def _write_report(report_dir: Path, year: int, payload: dict[str, Any]) -> tuple[Path, Path]:
    report_root = report_dir / f"year-{year}"
    report_root.mkdir(parents=True, exist_ok=True)
    json_path = report_root / "publish_results.json"
    md_path = report_root / "publish_results.md"
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    md_path.write_text(_to_markdown(payload), encoding="utf-8")
    return json_path, md_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Publish WRRL results from audited files.")
    parser.add_argument("--year", type=int, default=1999)
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument(
        "--report-dir",
        type=Path,
        default=Path("output") / "autopilot" / "runs",
        help="Folder for publish-results JSON/Markdown output.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    generated_at = datetime.now(timezone.utc).isoformat()

    data_root = _resolve_data_root(args.data_root)
    if data_root is None:
        payload = {
            "generated_at": generated_at,
            "success": False,
            "settings": {"year": args.year, "data_root": ""},
            "summary": {"audited_race_count": 0, "results_workbook": "", "warning_count": 0, "warnings": []},
            "error": {"message": "No data root configured. Set Data Root before publishing results."},
        }
        json_path, md_path = _write_report(args.report_dir, args.year, payload)
        print(payload["error"]["message"], flush=True)
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    input_dir, output_dir = _season_paths(data_root, args.year)
    ensure_output_subdirs(output_dir)
    sort_existing_output_files(output_dir)

    print("PROGRESS:STAGE:1:Loading audited race files", flush=True)
    audited_dir = build_input_paths(input_dir).audited_dir
    race_files = discover_race_files(audited_dir, excluded_names=race_discovery_exclusions())
    print(f"PROGRESS:RACES:{'|'.join(_race_names_for_progress(race_files))}", flush=True)
    print("PROGRESS:STAGE_DONE:1", flush=True)

    if not race_files:
        payload = {
            "generated_at": generated_at,
            "success": False,
            "settings": {"year": args.year, "data_root": str(data_root), "input_dir": str(input_dir), "output_dir": str(output_dir)},
            "summary": {"audited_race_count": 0, "results_workbook": "", "warning_count": 0, "warnings": []},
            "error": {"message": "No audited race files found in inputs/audited. Run audit/autopilot first."},
        }
        json_path, md_path = _write_report(args.report_dir, args.year, payload)
        print(payload["error"]["message"], flush=True)
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    print("PROGRESS:STAGE:2:Publishing final results", flush=True)
    scorer = LeagueScorer(input_dir=input_dir, output_dir=output_dir, year=args.year)

    previous_disable_pdf = os.environ.get("WRRL_DISABLE_PDF")
    if previous_disable_pdf is not None:
        os.environ.pop("WRRL_DISABLE_PDF", None)
    try:
        warnings = scorer.run(race_files=race_files)
    except Exception as exc:
        if previous_disable_pdf is not None:
            os.environ["WRRL_DISABLE_PDF"] = previous_disable_pdf
        payload = {
            "generated_at": generated_at,
            "success": False,
            "settings": {"year": args.year, "data_root": str(data_root), "input_dir": str(input_dir), "output_dir": str(output_dir)},
            "summary": {"audited_race_count": len(race_files), "results_workbook": "", "warning_count": 0, "warnings": []},
            "error": {"message": str(exc)},
        }
        json_path, md_path = _write_report(args.report_dir, args.year, payload)
        print(f"ERROR: {exc}", flush=True)
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    if previous_disable_pdf is not None:
        os.environ["WRRL_DISABLE_PDF"] = previous_disable_pdf

    print("PROGRESS:STAGE_DONE:2", flush=True)

    print("PROGRESS:STAGE:3:Writing publish summary", flush=True)
    results_workbook = None
    publish_dir = build_output_paths(output_dir).publish_xlsx_standings_dir
    if publish_dir.exists():
        candidates = sorted(publish_dir.glob("*.xlsx"), key=lambda item: item.stat().st_mtime, reverse=True)
        if candidates:
            results_workbook = candidates[0]

    payload = {
        "generated_at": generated_at,
        "success": True,
        "settings": {
            "year": args.year,
            "data_root": str(data_root),
            "input_dir": str(input_dir),
            "output_dir": str(output_dir),
        },
        "summary": {
            "audited_race_count": len(race_files),
            "results_workbook": str(results_workbook) if results_workbook is not None else "",
            "warning_count": len(warnings),
            "warnings": warnings,
        },
        "error": None,
    }
    json_path, md_path = _write_report(args.report_dir, args.year, payload)
    print("PROGRESS:STAGE_DONE:3", flush=True)
    print("Publish results success: True", flush=True)
    print(f"Warnings: {len(warnings)}", flush=True)
    print(f"Wrote: {json_path}", flush=True)
    print(f"Wrote: {md_path}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
