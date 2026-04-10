from __future__ import annotations

from argparse import Namespace
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from league_scorer.race_validation import RaceSchemaValidation
from scripts import analyse_data_quality as quality


def test_analyse_season_supports_xls_and_csv_raw_files(monkeypatch, tmp_path: Path):
    data_root = tmp_path / "data"
    year_root = data_root / "1999"
    input_dir = year_root / "inputs"
    raw_dir = input_dir / "raw_data"
    audited_dir = input_dir / "audited"
    raw_dir.mkdir(parents=True)
    audited_dir.mkdir(parents=True)

    (raw_dir / "Race #1.xls").write_text("ignored", encoding="utf-8")
    (raw_dir / "Race #2.csv").write_text("Pos,Bib,Name,Club,Gender,Category,Chip Time\n1,12,Alice,Club A,F,Sen,00:42:00\n", encoding="utf-8")
    audited_file = audited_dir / "Race #1 audited.xlsx"
    audited_file.write_text("ignored", encoding="utf-8")

    monkeypatch.setattr(quality, "load_race_dataframe", lambda path: pd.DataFrame({
        "Position": [1],
        "Name": ["Runner"],
        "Club": ["Club A"],
        "Gender": ["M"],
        "Category": ["Sen"],
        "Chip Time": ["00:40:00"],
    }))
    monkeypatch.setattr(
        quality,
        "validate_race_schema",
        lambda df, path: RaceSchemaValidation(
            column_map={"Name": "Name", "Club": "Club", "Gender": "Gender", "Category": "Category", "Position": "Position"},
            time_column="Chip Time",
            issues=[],
        ),
    )

    payload, json_path, md_path = quality.analyse_season(1999, data_root, tmp_path / "output")

    assert payload["raw_summary"]["rows"] >= 1
    assert json_path.exists()
    assert md_path.exists()
