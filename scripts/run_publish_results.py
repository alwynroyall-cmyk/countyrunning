"""Thin CLI wrapper that delegates to `league_scorer.publish.publish_results`.

This keeps backward compatibility for scripts and the GUI (which invokes the
script as a subprocess) while moving the implementation into a reusable
library location.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any
import sys


# When this script is executed directly from the `scripts/` folder by the
# GUI (which runs the script path), Python's import search path may not
# include the repository root. Ensure the repo root is on `sys.path` so
# `import league_scorer` resolves to the local package.
_repo_root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_repo_root))

# Now import the local publish function
from league_scorer.publish.publish import publish_results  # noqa: E402


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
    return publish_results(year=args.year, data_root=args.data_root, report_dir=args.report_dir)


if __name__ == "__main__":
    raise SystemExit(main())
