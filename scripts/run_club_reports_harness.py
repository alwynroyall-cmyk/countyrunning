"""Test harness: create minimal season data and run club report generator.

Creates a temporary `data/<year>/inputs` layout with a `clubs.xlsx` and one
audited race workbook, then calls the `generate_club_reports` entrypoint.
This is intended for local testing and CI harnessing when refining the
club report output.
"""
from __future__ import annotations

import tempfile
from pathlib import Path
import pandas as pd
import shutil
import sys

# Ensure repo root on sys.path when run from scripts/
_repo_root = Path(__file__).resolve().parents[1]
if str(_repo_root) not in sys.path:
    sys.path.insert(0, str(_repo_root))

# Diagnostic info to help debug import issues when run in different environments
print("[harness-file] running file:", Path(__file__).resolve())
print("[harness] sys.executable:", sys.executable)
print("[harness] cwd:", Path.cwd())
print("[harness] repo_root:", _repo_root)
print("[harness] sys.path[0:3]:", sys.path[0:3])

import os
print("[harness] repo_root exists:", _repo_root.exists())
try:
    print("[harness] repo_root listing:", sorted(os.listdir(str(_repo_root))))
except Exception as e:
    print("[harness] repo_root listing error:", e)

# Also try absolute path insert variants
rp = str(_repo_root)
if rp not in sys.path:
    sys.path.insert(0, rp)
if str(Path.cwd()) not in sys.path:
    sys.path.insert(0, str(Path.cwd()))

print("[harness] sys.path head after inserts:", sys.path[0:6])

from importlib import util
spec = util.find_spec('league_scorer')
print("[harness] find_spec('league_scorer'):", spec)

# Delay importing the generator until after the environment / paths are configured
# to avoid ModuleNotFoundError when the script is executed from varied contexts.

def make_clubs_xlsx(path: Path) -> None:
    df = pd.DataFrame(
        {
            "Club": ["Club A", "Club B"],
            "Preferred name": ["Club A", "Club B"],
            "Team A": [1, 1],
            "Team B": [2, 2],
        }
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False)


def make_audited_race(path: Path) -> None:
    df = pd.DataFrame(
        {
            "Position": [1, 2],
            "Name": ["Alice Runner", "Bob Runner"],
            "Club": ["Club A", "Club B"],
            "Gender": ["F", "M"],
            "Category": ["Sen", "Sen"],
            "Chip Time": ["00:40:00", "00:42:00"],
        }
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(path, index=False)


def main() -> int:
    # Use 1999 for test season by default
    year = 1999
    workspace = Path(__file__).resolve().parents[1]

    tmp = workspace / "test_club_report_data"
    if tmp.exists():
        shutil.rmtree(tmp)
    data_root = tmp

    year_root = data_root / str(year)
    inputs = year_root / "inputs"
    outputs = year_root / "outputs"

    control = inputs / "control"
    audited = inputs / "audited"

    # clubs.xlsx
    clubs_path = control / "clubs.xlsx"
    make_clubs_xlsx(clubs_path)

    # one audited race
    audited_race = audited / "Race 1 - audited.xlsx"
    make_audited_race(audited_race)

    # Run generator
    report_dir = outputs / "publish" / "club_reports"
    import importlib, traceback
    try:
        club_path = _repo_root / "league_scorer" / "publish" / "club_report.py"
        if club_path.exists():
            spec = util.spec_from_file_location("league_scorer.publish.club_report", str(club_path))
            mod = util.module_from_spec(spec)
            spec.loader.exec_module(mod)  # type: ignore
            generate_club_reports = getattr(mod, "generate_club_reports")
        else:
            mod = importlib.import_module("league_scorer.publish.club_report")
            generate_club_reports = getattr(mod, "generate_club_reports")
    except Exception:
        traceback.print_exc()
        print("[harness] failed to import generate_club_reports")
        return 2

    if "generate_club_reports" not in locals():
        print("[harness] generate_club_reports not defined after import; aborting")
        return 3

    rc = generate_club_reports(year=year, data_root=data_root, report_dir=report_dir)
    print(f"Club report generator returned: {rc}")
    print(f"DOCX output (if any): {outputs}")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
