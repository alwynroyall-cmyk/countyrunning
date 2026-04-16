"""Publish results functionality extracted from the CLI script.

This module provides `publish_results(year, data_root, report_dir)` which
performs the same steps as the prior `scripts/run_publish_results.py` script
but is callable from other code paths (GUI/tests) without forking a new
process.
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import traceback

from league_scorer.input.common_files import race_discovery_exclusions
from league_scorer.input.input_layout import build_input_paths
from league_scorer.output.output_layout import (
    build_output_paths,
    ensure_output_subdirs,
    export_publish_pdfs,
    package_publish_artifacts,
    sort_existing_output_files,
)
from league_scorer.publish import club_report
from league_scorer.input.source_loader import discover_race_files

# Reuse helper functions from the scripts package
from scripts.autopilot.run_full_autopilot import _race_names_for_progress, _resolve_data_root, _season_paths


def _to_markdown(payload: dict[str, Any]) -> str:
    lines = [
        "# WRRL Admin Suite Publish Results",
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
        "- This path publishes from audited files and includes PDF outputs and club reports.",
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


def publish_results(
    year: int,
    data_root: Path | None,
    report_dir: Path,
    export_pdf_dir: Path | None = None,
) -> int:
    """Publish final results for `year` using `data_root` and write reports to `report_dir`.

    If *export_pdf_dir* is provided, all published PDF files are copied into that
    folder after conversion and packaging.

    This workflow now also generates club reports as part of the publish path.

    Returns an exit code (0 success, non-zero failure). Prints progress lines
    and summary messages to stdout/stderr to preserve the original CLI behaviour.
    """
    generated_at = datetime.now(timezone.utc).isoformat()

    data_root_resolved = _resolve_data_root(data_root)
    if data_root_resolved is None:
        payload = {
            "generated_at": generated_at,
            "success": False,
            "settings": {"year": year, "data_root": ""},
            "summary": {"audited_race_count": 0, "results_workbook": "", "warning_count": 0, "warnings": []},
            "error": {"message": "No data root configured. Set Data Root before publishing results."},
        }
        json_path, md_path = _write_report(report_dir, year, payload)
        print(payload["error"]["message"], flush=True)
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    input_dir, output_dir = _season_paths(data_root_resolved, year)
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
            "settings": {"year": year, "data_root": str(data_root_resolved), "input_dir": str(input_dir), "output_dir": str(output_dir)},
            "summary": {"audited_race_count": 0, "results_workbook": "", "warning_count": 0, "warnings": []},
            "error": {"message": "No audited race files found in inputs/audited. Run audit/autopilot first."},
        }
        json_path, md_path = _write_report(report_dir, year, payload)
        print(payload["error"]["message"], flush=True)
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    try:
        # Publish for GUI should *not* re-run the scorer — it should only
        # ensure DOCX publish artefacts are converted to PDF. This routine
        # scans the publish DOCX folders and converts any missing/stale PDFs.
        print("PROGRESS:STAGE:2:Converting DOCX -> PDF", flush=True)

        output_paths = build_output_paths(output_dir)

        conversion_pairs = [
            (output_paths.publish_docx_league_updates_dir, output_paths.publish_pdf_league_updates_dir),
            (output_paths.publish_docx_race_cards_dir, output_paths.publish_pdf_race_cards_dir),
        ]

        converted = 0
        skipped = 0
        warnings: list[str] = []

        # Ensure PDF conversion runs even if WRRL_DISABLE_PDF is set in the
        # environment elsewhere — temporarily clear it for this conversion
        # process and restore the previous value afterwards.
        previous_disable_pdf = os.environ.get("WRRL_DISABLE_PDF")
        if previous_disable_pdf is not None:
            os.environ.pop("WRRL_DISABLE_PDF", None)

        try:
            from docx2pdf import convert  # type: ignore
        except Exception:
            convert = None  # type: ignore

        for docx_dir, pdf_dir in conversion_pairs:
            try:
                if not docx_dir.exists():
                    continue
                pdf_dir.mkdir(parents=True, exist_ok=True)
                for docx in sorted(docx_dir.glob("*.docx")):
                    target_pdf = pdf_dir / docx.with_suffix(".pdf").name
                    try:
                        # Skip conversion if PDF is up-to-date
                        if target_pdf.exists() and target_pdf.stat().st_mtime >= docx.stat().st_mtime:
                            skipped += 1
                            continue
                        if convert is None:
                            warnings.append(f"docx2pdf not available: {docx.name}")
                            continue
                        convert(str(docx), str(target_pdf))
                        converted += 1
                    except Exception as exc:
                        warnings.append(f"{docx.name}: PDF conversion skipped ({exc})")
            except Exception:
                # Non-fatal: continue to next folder
                continue
        if previous_disable_pdf is None:
            os.environ.pop("WRRL_DISABLE_PDF", None)
        else:
            os.environ["WRRL_DISABLE_PDF"] = previous_disable_pdf
    except Exception as exc:
        error_message = f"Publish failed: {exc}"
        tb = traceback.format_exc()
        print(error_message, flush=True)
        print(tb, flush=True)
        payload = {
            "generated_at": generated_at,
            "success": False,
            "settings": {
                "year": year,
                "data_root": str(data_root_resolved),
                "input_dir": str(input_dir),
                "output_dir": str(output_dir),
            },
            "summary": {
                "audited_race_count": len(race_files),
                "results_workbook": "",
                "warning_count": 0,
                "warnings": [],
            },
            "error": {"message": error_message, "traceback": tb},
        }
        json_path, md_path = _write_report(report_dir, year, payload)
        print(f"Wrote: {json_path}", flush=True)
        print(f"Wrote: {md_path}", flush=True)
        return 1

    club_result = -1
    try:
        print("PROGRESS:STAGE:3:Generating club reports", flush=True)
        club_result = club_report.generate_club_reports(year, data_root_resolved, report_dir)
        if club_result != 0:
            warnings.append("Club report generation failed. See club_reports.json for details.")
        print("PROGRESS:STAGE_DONE:3", flush=True)
    except Exception as exc:
        warnings.append(f"Club report generation skipped: {exc}")

    try:
        package_publish_artifacts(output_dir)
    except Exception as exc:
        warnings.append(f"Publish package creation skipped: {exc}")

    payload = {
        "generated_at": generated_at,
        "success": True,
        "settings": {
            "year": year,
            "data_root": str(data_root_resolved),
            "input_dir": str(input_dir),
            "output_dir": str(output_dir),
        },
        "summary": {
            "audited_race_count": len(race_files),
            "results_workbook": "",
            "converted_pdf_count": converted,
            "skipped_pdf_count": skipped,
            "club_reports_generated": club_result == 0,
            "warning_count": len(warnings),
            "warnings": warnings,
        },
        "error": None,
    }

    exported_count = 0
    if export_pdf_dir is not None:
        try:
            export_path = export_publish_pdfs(output_dir, export_pdf_dir, flatten=True)
            exported_count = sum(1 for _ in export_path.glob("**/*.pdf"))
        except Exception as exc:
            warnings.append(f"Publish PDF export skipped: {exc}")

    payload["summary"]["exported_pdf_count"] = exported_count

    json_path, md_path = _write_report(report_dir, year, payload)
    print("PROGRESS:STAGE_DONE:2", flush=True)
    print(f"Converted PDFs: {converted}", flush=True)
    print(f"Skipped PDFs: {skipped}", flush=True)
    if export_pdf_dir is not None:
        print(f"Exported PDFs: {exported_count}", flush=True)
    if warnings:
        print(f"Warnings: {len(warnings)}", flush=True)
    print(f"Wrote: {json_path}", flush=True)
    print(f"Wrote: {md_path}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit("This module is a library; import and call `publish_results()` instead.")
