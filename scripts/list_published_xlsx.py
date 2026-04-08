"""List published .xlsx files for a season and write a Markdown report.

Usage: python -m scripts.list_published_xlsx --year 2026 --data-root <path>

Writes `xlsx_files.md` to `output/autopilot/runs/year-<YEAR>/xlsx_files.md` by default.
"""
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.run_full_autopilot import _resolve_data_root, _season_paths
from league_scorer.output_layout import build_output_paths


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="List published XLSX files and write a Markdown report.")
    p.add_argument("--year", type=int, default=1999)
    p.add_argument("--data-root", type=Path, default=None)
    p.add_argument("--report-dir", type=Path, default=Path("output") / "autopilot" / "runs")
    return p.parse_args()


def _format_file_line(path: Path) -> str:
    stat = path.stat()
    mtime = datetime.fromtimestamp(stat.st_mtime).isoformat(sep=" ")
    size_kb = stat.st_size / 1024.0
    return f"- {path.as_posix()} — {size_kb:.1f} KB, modified {mtime}"


def main() -> int:
    args = parse_args()
    data_root = _resolve_data_root(args.data_root)
    if data_root is None:
        print("No data root configured. Set Data Root before listing published files.")
        return 1

    input_dir, output_dir = _season_paths(data_root, args.year)
    out = build_output_paths(output_dir)

    # Search recursively under the output_dir for .xlsx files
    xlsx_files = sorted(output_dir.rglob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)

    report_root = args.report_dir / f"year-{args.year}"
    report_root.mkdir(parents=True, exist_ok=True)
    md_path = report_root / "xlsx_files.md"

    lines = [f"# XLSX files for year {args.year}", ""]
    lines.append(f"Generated: {datetime.utcnow().isoformat()}Z")
    lines.append("")

    if not xlsx_files:
        lines.append("No .xlsx files found under the season output folder.")
    else:
        lines.append(f"Found {len(xlsx_files)} .xlsx file{'s' if len(xlsx_files) != 1 else ''}:")
        lines.append("")
        for p in xlsx_files:
            lines.append(_format_file_line(p))

    lines.append("")
    md_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote: {md_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
