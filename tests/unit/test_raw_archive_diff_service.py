from pathlib import Path

import pandas as pd

from league_scorer.raw_archive_diff_service import (
    build_side_by_side_diff,
    list_comparable_file_pairs,
    load_comparable_lines,
)


def test_list_comparable_file_pairs_matches_same_filenames(tmp_path: Path):
    input_dir = tmp_path / "inputs"
    raw_dir = input_dir / "raw_data"
    archive_dir = input_dir / "raw_data_archive"
    raw_dir.mkdir(parents=True)
    archive_dir.mkdir(parents=True)

    (raw_dir / "Race 1 - Example.csv").write_text("name,club\nAlice,Team A\n", encoding="utf-8")
    (archive_dir / "Race 1 - Example.csv").write_text("name,club\nAlice,Team A\n", encoding="utf-8")
    (raw_dir / "Race 2 - New.csv").write_text("name,club\nBob,Team B\n", encoding="utf-8")

    pairs = list_comparable_file_pairs(input_dir)

    assert [pair.filename for pair in pairs] == ["Race 1 - Example.csv"]


def test_load_comparable_lines_renders_excel_sheets_to_text_lines(tmp_path: Path):
    workbook_path = tmp_path / "Race 1 - Example.xlsx"
    with pd.ExcelWriter(workbook_path) as writer:
        pd.DataFrame({"Name": ["Alice"], "Club": ["Team A"]}).to_excel(writer, sheet_name="Race 1", index=False)

    lines = load_comparable_lines(workbook_path)

    assert lines[0] == "=== Sheet: Race 1 ==="
    assert lines[1] == "Name | Club"
    assert lines[2] == "Alice | Team A"


def test_build_side_by_side_diff_marks_replaced_deleted_and_inserted_rows():
    rows = build_side_by_side_diff(
        ["header", "alice", "charlie"],
        ["header", "bob", "charlie", "delta"],
    )

    assert rows[0].status == "same"
    assert rows[1].status == "replace"
    assert rows[1].left_text == "alice"
    assert rows[1].right_text == "bob"
    assert rows[2].status == "same"
    assert rows[3].status == "insert"
    assert rows[3].left_text == ""
    assert rows[3].right_text == "delta"


def test_load_comparable_lines_handles_utf16_html_xls_exports(tmp_path: Path):
        legacy_xls = tmp_path / "Race 1 - Legacy.xls"
        html = """
        <table>
            <tr><th>Name</th><th>Club</th></tr>
            <tr><td>Alice</td><td>Team A</td></tr>
        </table>
        """.strip()
        legacy_xls.write_bytes(html.encode("utf-16"))

        lines = load_comparable_lines(legacy_xls)

        assert lines[0] == "=== Table: 1 ==="
        assert "Name" in lines[1]
        assert "Alice" in lines[2]