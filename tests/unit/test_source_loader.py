from pathlib import Path

import pandas as pd

from league_scorer.source_loader import discover_race_files, load_race_dataframe


def test_discover_race_files_prefers_highest_priority_extension(tmp_path: Path):
    (tmp_path / "Race #1.xlsx").write_text("", encoding="utf-8")
    (tmp_path / "Race #1.xlsm").write_text("", encoding="utf-8")
    (tmp_path / "Race #1.xls").write_text("", encoding="utf-8")
    (tmp_path / "Race #1.csv").write_text("Pos,Bib,Name\n1,12,Alice\n", encoding="utf-8")

    race_files = discover_race_files(tmp_path)

    assert race_files == {1: tmp_path / "Race #1.xlsx"}


def test_discover_race_files_includes_csv_after_excel_priority(tmp_path: Path):
    (tmp_path / "Race #1.csv").write_text("Pos,Bib,Name\n1,12,Alice\n", encoding="utf-8")
    race_files = discover_race_files(tmp_path)

    assert race_files == {1: tmp_path / "Race #1.csv"}


def test_discover_race_files_excludes_control_names(tmp_path: Path):
    (tmp_path / "clubs.xlsx").write_text("", encoding="utf-8")
    race_files = discover_race_files(tmp_path, excluded_names=["clubs.xlsx"])

    assert race_files == {}


def test_load_race_dataframe_reads_csv(tmp_path: Path):
    path = tmp_path / "Race #1.csv"
    path.write_text("Pos,Bib,Name\n1,12,Alice\n", encoding="utf-8")

    df = load_race_dataframe(path)

    assert list(df.columns) == ["Pos", "Bib", "Name"]
    assert df.iloc[0]["Name"] == "Alice"


def test_load_race_dataframe_reads_html_disguised_xls(tmp_path: Path):
    html = """
    <html>
      <body>
        <table>
          <tr><td>Pos</td><td>Name</td></tr>
          <tr><td>1</td><td>Alice</td></tr>
        </table>
      </body>
    </html>
    """
    path = tmp_path / "Race #1.xls"
    path.write_text(html, encoding="utf-8")

    df = load_race_dataframe(path)

    assert list(df.columns) == ["Pos", "Name"]
    assert df.iloc[0]["Name"] == "Alice"
