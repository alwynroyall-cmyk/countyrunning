from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from league_scorer import audit_data_service
from league_scorer.audit_data_service import (
    ACTIONABLE_COLUMNS,
    ACTIONABLE_SHEET,
    SEASON_AUDIT_WORKBOOK,
    find_latest_audit_workbook,
    find_latest_manual_review_workbook,
    list_audit_workbooks,
    load_actionable_issues,
)


def _write_excel(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name)


def test_list_audit_workbooks_finds_output_and_audited_files(tmp_path: Path, monkeypatch) -> None:
    data_root = tmp_path / "data"
    year_dir = data_root / "1999"
    output_dir = year_dir / "outputs"
    input_dir = year_dir / "inputs"
    (output_dir / "audit" / "workbooks").mkdir(parents=True)
    (input_dir / "audited").mkdir(parents=True)

    audit_file = output_dir / "audit" / "workbooks" / SEASON_AUDIT_WORKBOOK
    audit_file.write_text("audit", encoding="utf-8")
    audited_file = input_dir / "audited" / "Race 1 (audited).xlsx"
    audited_file.write_text("audited", encoding="utf-8")

    monkeypatch.setattr(audit_data_service.session_config, "_data_root", data_root)
    monkeypatch.setattr(audit_data_service.session_config, "_year", 1999)

    workbooks = list_audit_workbooks()

    assert f"Audit / {audit_file.name}" in workbooks
    assert f"Inputs / {audited_file.name}" in workbooks
    assert workbooks[f"Audit / {audit_file.name}"] == audit_file
    assert workbooks[f"Inputs / {audited_file.name}"] == audited_file


def test_find_latest_audit_workbook_prefers_season_audit(tmp_path: Path, monkeypatch) -> None:
    data_root = tmp_path / "data"
    year_dir = data_root / "1999"
    output_dir = year_dir / "outputs"
    (output_dir / "audit" / "workbooks").mkdir(parents=True)

    season_audit = output_dir / "audit" / "workbooks" / SEASON_AUDIT_WORKBOOK
    other_audit = output_dir / "audit" / "workbooks" / "Race 1 - Audit.xlsx"
    season_audit.write_text("season", encoding="utf-8")
    other_audit.write_text("other", encoding="utf-8")

    monkeypatch.setattr(audit_data_service.session_config, "_data_root", data_root)
    monkeypatch.setattr(audit_data_service.session_config, "_year", 1999)

    assert find_latest_audit_workbook() == season_audit


def test_load_actionable_issues_filters_resolved_rows(tmp_path: Path, monkeypatch) -> None:
    data_root = tmp_path / "data"
    year_dir = data_root / "1999"
    output_dir = year_dir / "outputs"
    manual_changes_dir = output_dir / "audit" / "manual-changes"
    manual_changes_dir.mkdir(parents=True)

    workbook = output_dir / "audit" / "workbooks" / "Season Audit.xlsx"
    _write_excel(
        workbook,
        {
            ACTIONABLE_SHEET: pd.DataFrame(
                [
                    {
                        "Type": "Row",
                        "Severity": "warning",
                        "Issue Code": "AUD-ROW-001",
                        "Race": "1",
                        "Source Row": "2",
                        "Key": "",
                        "Name": "Alice",
                        "Club": "Club A",
                        "Message": "Missing gender",
                        "Next Step": "Fix source",
                    },
                    {
                        "Type": "Club",
                        "Severity": "warning",
                        "Issue Code": "AUD-CLUB-002",
                        "Race": "",
                        "Source Row": "",
                        "Key": "Club A",
                        "Name": "",
                        "Club": "Club A",
                        "Message": "Inconsistent divisions",
                        "Next Step": "Review club lookup",
                    },
                ]
            )
        },
    )

    manual_audit = manual_changes_dir / "Manual_Data_Audit.xlsx"
    _write_excel(
        manual_audit,
        {
            "Manual Changes": pd.DataFrame(
                [
                    {
                        "Timestamp": datetime.now().isoformat(),
                        "Resolution Status": "Resolved",
                        "Issue Identity": "row|aud-row-001|1|2|alice",
                    }
                ]
            )
        },
    )

    filtered = load_actionable_issues(workbook)

    assert len(filtered) == 1
    assert filtered.iloc[0]["Issue Code"] == "AUD-CLUB-002"


def test_load_actionable_issues_returns_empty_dataframe_when_sheet_missing(tmp_path: Path) -> None:
    workbook = tmp_path / "Season Audit.xlsx"
    with pd.ExcelWriter(workbook, engine="openpyxl") as writer:
        pd.DataFrame([{"A": 1}]).to_excel(writer, index=False, sheet_name="Other Sheet")

    result = load_actionable_issues(workbook)

    assert list(result.columns) == ACTIONABLE_COLUMNS
    assert result.empty


def test_find_latest_manual_review_workbook_prefers_non_empty_review_sheet(tmp_path: Path, monkeypatch) -> None:
    data_root = tmp_path / "data"
    year_dir = data_root / "1999"
    output_dir = year_dir / "outputs"
    workbooks_dir = output_dir / "audit" / "workbooks"
    workbooks_dir.mkdir(parents=True)

    empty_club = workbooks_dir / "Empty Club Review.xlsx"
    with pd.ExcelWriter(empty_club, engine="openpyxl") as writer:
        pd.DataFrame([], columns=["Club Review"]).to_excel(writer, index=False, sheet_name="Club Review")

    filled_name = workbooks_dir / "Name Review.xlsx"
    with pd.ExcelWriter(filled_name, engine="openpyxl") as writer:
        pd.DataFrame([{"Raw Name": "Alice"}]).to_excel(writer, index=False, sheet_name="Name Review")

    monkeypatch.setattr(audit_data_service.session_config, "_data_root", data_root)
    monkeypatch.setattr(audit_data_service.session_config, "_year", 1999)

    assert find_latest_manual_review_workbook() == filled_name
