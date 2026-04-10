from pathlib import Path

import pandas as pd

from league_scorer import manual_data_audit
from league_scorer.manual_data_audit import (
    _machine_id,
    _safe_rel_path,
    log_manual_data_changes,
)


def test_log_manual_data_changes_fails_when_output_not_configured(monkeypatch):
    monkeypatch.setattr(manual_data_audit.session_config, "_data_root", None)
    monkeypatch.setattr(manual_data_audit.session_config, "_year", 2025)

    error = log_manual_data_changes(
        [{"runner": "Alice", "field": "Club", "old_value": "Old Club", "new_value": "New Club"}],
        source="UI",
        action="Update",
    )

    assert error == "Output directory is not configured for this session."


def test_log_manual_data_changes_writes_workbook_and_appends_rows(tmp_path, monkeypatch):
    data_root = tmp_path / "data"
    monkeypatch.setattr(manual_data_audit.session_config, "_data_root", data_root)
    monkeypatch.setattr(manual_data_audit.session_config, "_year", 2026)
    monkeypatch.setattr(manual_data_audit, "getpass", type("G", (), {"getuser": staticmethod(lambda: "tester")})())
    monkeypatch.setattr(manual_data_audit.os, "getenv", lambda key, default="": "TESTHOST" if key == "COMPUTERNAME" else default)

    changes = [
        {
            "runner": "Alice",
            "field": "Club",
            "old_value": "Old Club",
            "new_value": "New Club",
            "file_path": tmp_path / "data" / "2026" / "inputs" / "raw_data" / "Race 1.xls",
            "row_idx": 5,
            "issue_identity": "row|aud-row-001|1|5|alice",
            "issue_code": "AUD-ROW-001",
            "issue_race": "1",
            "issue_source_row": "5",
            "resolution_status": "Resolved",
            "verification_note": "Checked",
        }
    ]

    error = log_manual_data_changes(changes, source="UI", action="Update")
    assert error is None

    output_dir = data_root / "2026" / "outputs"
    workbook_path = output_dir / "audit" / "manual-changes" / "Manual_Data_Audit.xlsx"
    assert workbook_path.exists()

    df = pd.read_excel(workbook_path, sheet_name="Manual Changes", dtype=str).fillna("")
    assert len(df) == 1
    assert df.iloc[0]["User"] == "tester"
    assert df.iloc[0]["Action"] == "Update"
    assert df.iloc[0]["Issue Code"] == "AUD-ROW-001"
    assert df.iloc[0]["Issue Identity"] == "row|aud-row-001|1|5|alice"

    # Append a second row and verify the workbook grows.
    error = log_manual_data_changes(
        [
            {
                "runner": "Bob",
                "field": "Category",
                "old_value": "Sen",
                "new_value": "V40",
                "file_path": tmp_path / "data" / "2026" / "inputs" / "raw_data" / "Race 1.xls",
                "row_idx": 7,
                "issue_identity": "row|aud-row-002|1|7|bob",
                "issue_code": "AUD-ROW-002",
                "issue_race": "1",
                "issue_source_row": "7",
                "resolution_status": "Resolved",
                "verification_note": "Checked",
            }
        ],
        source="UI",
        action="Update",
    )
    assert error is None

    df = pd.read_excel(workbook_path, sheet_name="Manual Changes", dtype=str).fillna("")
    assert len(df) == 2
    assert set(df["Runner"]) == {"Alice", "Bob"}


def test_safe_rel_path_returns_relative_path_for_output_parent(tmp_path):
    output_dir = tmp_path / "outputs"
    path = output_dir.parent / "inputs" / "raw_data" / "Race 1.xls"
    assert _safe_rel_path(path, output_dir) == str(path.relative_to(output_dir.parent))


def test_safe_rel_path_returns_absolute_path_when_relative_fails(tmp_path):
    output_dir = tmp_path / "outputs"
    path = Path("C:/other/location/Race 1.xls")
    assert _safe_rel_path(path, output_dir) == str(path)


def test_machine_id_uses_computername_env(monkeypatch):
    monkeypatch.setattr(manual_data_audit.os, "getenv", lambda key, default="": "MYHOST" if key == "COMPUTERNAME" else default)
    assert _machine_id() == "MYHOST"


def test_machine_id_falls_back_to_hostname(monkeypatch):
    monkeypatch.setattr(manual_data_audit.os, "getenv", lambda key, default="": "")
    monkeypatch.setattr(manual_data_audit.socket, "gethostname", lambda: "fallback-host")
    assert _machine_id() == "fallback-host"
