from __future__ import annotations

from argparse import Namespace
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from league_scorer.models import RaceIssue
from league_scorer.race_validation import RaceSchemaValidation
from scripts import run_staged_checks as staged


def test_run_checks_allows_missing_data(monkeypatch, tmp_path):
    monkeypatch.setattr(staged, "_resolve_data_root", lambda _explicit: None)

    args = Namespace(
        year=1999,
        data_root=None,
        report_dir=tmp_path / "reports",
        baseline_file=tmp_path / "baseline.json",
        write_baseline=False,
        allow_missing_data=True,
    )

    success, results = staged.run_checks(args)

    assert success is True
    assert len(results) == 1
    assert results[0].status == "skipped"


def test_run_checks_baseline_write_and_compare(monkeypatch, tmp_path):
    data_root = tmp_path / "data"
    year_root = data_root / "1999"
    input_dir = year_root / "inputs"
    output_dir = year_root / "outputs"
    raw_dir = input_dir / "Raw Data"

    raw_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    raw_file = raw_dir / "Race 1 - raw.xlsx"
    raw_file.write_text("placeholder", encoding="utf-8")

    audited_file = input_dir / "Race 1 - audited.xlsx"
    audited_file.write_text("placeholder", encoding="utf-8")

    results_path = output_dir / "Race 8 -- Results.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Race Summary"
    ws.append(["Race", "Value"])
    ws.append([1, 123])
    wb.save(results_path)
    wb.close()

    monkeypatch.setattr(staged, "_resolve_data_root", lambda _explicit: data_root)
    monkeypatch.setattr(staged, "_find_raw_dir", lambda _input: raw_dir)
    monkeypatch.setattr(staged, "_discover_raw_race_files", lambda _raw: {1: [raw_file]})

    def fake_load_race_dataframe(_path):
        return pd.DataFrame(
            {
                "Position": [1],
                "Name": ["Runner"],
                "Club": ["Club A"],
                "Gender": ["M"],
                "Category": ["Sen"],
                "Chip Time": ["00:40:00"],
            }
        )

    monkeypatch.setattr(staged, "load_race_dataframe", fake_load_race_dataframe)
    monkeypatch.setattr(
        staged,
        "validate_race_schema",
        lambda df, _path: RaceSchemaValidation(
            column_map={"Name": "Name", "Club": "Club", "Gender": "Gender", "Category": "Category", "Position": "Position"},
            time_column="Chip Time",
            issues=[],
        ),
    )
    monkeypatch.setattr(staged, "discover_race_files", lambda _input, excluded_names=(): {1: audited_file})

    class FakeAuditor:
        def __init__(self, input_dir: Path, output_dir: Path, year: int):
            self.input_dir = input_dir
            self.output_dir = output_dir
            self.year = year
            self.all_race_issues = {1: [RaceIssue("warning", "sample")]} 

        def run(self, race_files=None):
            audit_path = self.output_dir / "audit" / "Season Audit.xlsx"
            audit_path.parent.mkdir(parents=True, exist_ok=True)
            wb = Workbook()
            wb.save(audit_path)
            wb.close()
            return audit_path

    class FakeScorer:
        def __init__(self, input_dir: Path, output_dir: Path, year: int):
            self.input_dir = input_dir
            self.output_dir = output_dir
            self.year = year

        def run(self, race_files=None):
            return []

    monkeypatch.setattr(staged, "LeagueAuditor", FakeAuditor)
    monkeypatch.setattr(staged, "LeagueScorer", FakeScorer)
    monkeypatch.setattr(
        staged,
        "analyse_season",
        lambda year, data_root, output_dir: (
            {
                "audited_summary": {
                    "blank_category_pct": 0.0,
                    "blank_name_pct": 0.0,
                    "blank_gender_pct": 0.0,
                    "invalid_time_pct": 0.0,
                },
                "hotspots": {"audited_blank_category": []},
            },
            output_dir / f"year-{year}" / "data_quality_report.json",
            output_dir / f"year-{year}" / "data_quality_report.md",
        ),
    )

    from league_scorer.graphical import results_workbook

    monkeypatch.setattr(results_workbook, "find_latest_results_workbook", lambda _out: results_path)

    baseline_file = tmp_path / "baseline.json"

    args_write = Namespace(
        year=1999,
        data_root=None,
        report_dir=tmp_path / "reports-write",
        baseline_file=baseline_file,
        write_baseline=True,
        quality_gate_threshold=80.0,
        allow_missing_data=False,
    )

    success_write, _results_write = staged.run_checks(args_write)
    assert success_write is True
    assert baseline_file.exists()

    args_compare = Namespace(
        year=1999,
        data_root=None,
        report_dir=tmp_path / "reports-compare",
        baseline_file=baseline_file,
        write_baseline=False,
        quality_gate_threshold=80.0,
        allow_missing_data=False,
    )

    success_compare, results_compare = staged.run_checks(args_compare)
    assert success_compare is True
    stage4 = next(r for r in results_compare if r.stage == 4)
    assert stage4.status == "passed"


def test_run_checks_stage2_quality_gate_blocks(monkeypatch, tmp_path):
    data_root = tmp_path / "data"
    input_dir = data_root / "1999" / "inputs"
    raw_dir = input_dir / "Raw Data"
    raw_dir.mkdir(parents=True, exist_ok=True)

    raw_file = raw_dir / "Race 1 - raw.xlsx"
    raw_file.write_text("placeholder", encoding="utf-8")
    audited_file = input_dir / "Race 1 - audited.xlsx"
    audited_file.parent.mkdir(parents=True, exist_ok=True)
    audited_file.write_text("placeholder", encoding="utf-8")

    monkeypatch.setattr(staged, "_resolve_data_root", lambda _explicit: data_root)
    monkeypatch.setattr(staged, "_find_raw_dir", lambda _input: raw_dir)
    monkeypatch.setattr(staged, "_discover_raw_race_files", lambda _raw: {1: [raw_file]})
    monkeypatch.setattr(staged, "discover_race_files", lambda _input, excluded_names=(): {1: audited_file})
    monkeypatch.setattr(
        staged,
        "load_race_dataframe",
        lambda _path: pd.DataFrame(
            {
                "Position": [1],
                "Name": ["Runner"],
                "Club": ["Club A"],
                "Gender": ["M"],
                "Category": ["Sen"],
                "Chip Time": ["00:40:00"],
            }
        ),
    )
    monkeypatch.setattr(
        staged,
        "validate_race_schema",
        lambda df, _path: RaceSchemaValidation(
            column_map={"Name": "Name", "Club": "Club", "Gender": "Gender", "Category": "Category", "Position": "Position"},
            time_column="Chip Time",
            issues=[],
        ),
    )
    monkeypatch.setattr(
        staged,
        "analyse_season",
        lambda year, data_root, output_dir: (
            {
                "audited_summary": {
                        "blank_category_pct": 50.0,
                        "blank_name_pct": 20.0,
                        "blank_gender_pct": 20.0,
                        "invalid_time_pct": 20.0,
                },
                    "hotspots": {"audited_blank_category": [{"file": "race1.xlsx", "blank_category_pct": 50.0}]},
            },
            output_dir / f"year-{year}" / "data_quality_report.json",
            output_dir / f"year-{year}" / "data_quality_report.md",
        ),
    )

    args = Namespace(
        year=1999,
        data_root=None,
        report_dir=tmp_path / "reports",
        baseline_file=tmp_path / "baseline.json",
        write_baseline=False,
        quality_gate_threshold=80.0,
        allow_missing_data=False,
    )

    success, results = staged.run_checks(args)

    assert success is False
    stage2 = next(r for r in results if r.stage == 2)
    assert stage2.status == "failed"
    assert stage2.details["data_quality"]["quality_gate_passed"] is False
