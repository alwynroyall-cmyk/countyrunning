from pathlib import Path

from scripts import run_staged_checks as staged


def test_discover_raw_race_files_supports_non_xlsx_extensions(tmp_path: Path):
    raw_dir = tmp_path / "inputs" / "raw_data"
    raw_dir.mkdir(parents=True)
    (raw_dir / "Race #1.xls").write_text("", encoding="utf-8")
    (raw_dir / "Race #2.csv").write_text("Pos,Bib,Name\n1,12,Alice\n", encoding="utf-8")

    results = staged._discover_raw_race_files(raw_dir)

    assert set(results) == {1, 2}
    assert any(str(path).endswith("Race #1.xls") for path in results[1])
    assert any(str(path).endswith("Race #2.csv") for path in results[2])
