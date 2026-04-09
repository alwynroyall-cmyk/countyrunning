"""Thin CLI wrapper that delegates to `league_scorer.publish.club_report`.

This keeps backward compatibility for scripts and the GUI (which invokes the
script as a subprocess) while moving the implementation into a reusable
library location.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import sys


# When this script is executed directly from the `scripts/` folder by the
# GUI (which runs the script path), Python's import search path may not
# include the repository root. Ensure the repo root is on `sys.path` so
# `import league_scorer` resolves to the local package.
_repo_root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_repo_root))

# Now import the local club report function
try:
    from league_scorer.publish.club_report import generate_club_reports  # noqa: E402
except Exception:
    import importlib.util
    import traceback
    print("Failed normal import of league_scorer.publish.club_report; falling back to direct file import")
    traceback.print_exc()
    club_path = _repo_root / "league_scorer" / "publish" / "club_report.py"
    if not club_path.exists():
        raise FileNotFoundError(f"club_report.py not found at {club_path}")
    spec = importlib.util.spec_from_file_location("league_scorer.publish.club_report", str(club_path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore
    generate_club_reports = getattr(mod, "generate_club_reports")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate club-level reports for a season.")
    parser.add_argument("--year", type=int, default=1999)
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument(
        "--report-dir",
        type=Path,
        default=Path("output") / "publish" / "club_reports",
        help="Folder for club-report JSON/Markdown output.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    return generate_club_reports(year=args.year, data_root=args.data_root, report_dir=args.report_dir)


if __name__ == "__main__":
    raise SystemExit(main())
