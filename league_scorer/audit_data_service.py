"""Shared helpers for reading audit workbooks and actionable issue data."""

from __future__ import annotations

from pathlib import Path
from datetime import datetime

import pandas as pd

from .input_layout import build_input_paths
from .issue_tracking import build_issue_identity
from .output_layout import build_output_paths
from .session_config import config as session_config

AUDIT_DIR = "audit/workbooks"
SEASON_AUDIT_WORKBOOK = "Season Audit.xlsx"
ACTIONABLE_SHEET = "Actionable Issues"

ACTIONABLE_COLUMNS = [
    "Type",
    "Severity",
    "Issue Code",
    "Race",
    "Source Row",
    "Key",
    "Name",
    "Club",
    "Message",
    "Next Step",
]


def list_audit_workbooks() -> dict[str, Path]:
    """Return selectable audit workbooks for UI panels, newest first."""
    workbooks: dict[str, Path] = {}

    output_dir = session_config.output_dir
    if output_dir:
        audit_dir = build_output_paths(output_dir).audit_workbooks_dir
        if audit_dir.exists():
            for path in sorted(audit_dir.glob("*.xlsx"), key=lambda item: item.stat().st_mtime, reverse=True):
                workbooks[f"Audit / {path.name}"] = path

    input_dir = session_config.input_dir
    audited_dir = build_input_paths(input_dir).audited_dir if input_dir else None
    if audited_dir and audited_dir.exists():
        for path in sorted(audited_dir.glob("* (audited).xlsx"), key=lambda item: item.stat().st_mtime, reverse=True):
            workbooks[f"Inputs / {path.name}"] = path

    return workbooks


def find_latest_audit_workbook() -> Path | None:
    """Find the best available audit workbook, preferring Season Audit workbook names."""
    candidates = list(list_audit_workbooks().values())

    if not candidates:
        return None

    for path in candidates:
        if path.name == SEASON_AUDIT_WORKBOOK:
            return path
    return candidates[0]


def load_actionable_issues(workbook: Path) -> pd.DataFrame:
    """Load the Actionable Issues sheet as a normalized DataFrame."""
    xl = pd.ExcelFile(workbook)
    if ACTIONABLE_SHEET not in xl.sheet_names:
        return pd.DataFrame(columns=ACTIONABLE_COLUMNS)

    df = xl.parse(ACTIONABLE_SHEET, dtype=str).fillna("")
    for col in ACTIONABLE_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[ACTIONABLE_COLUMNS].copy()

    resolved_identities = _load_recently_resolved_issue_keys(workbook)
    if resolved_identities:
        identities = df.apply(lambda row: build_issue_identity(row.to_dict()), axis=1)
        df = df.loc[~identities.isin(resolved_identities)].reset_index(drop=True)
    return df


def _load_recently_resolved_issue_keys(workbook: Path) -> set[str]:
    manual_audit_path = workbook.parent.parent / "manual-changes" / "Manual_Data_Audit.xlsx"
    if not manual_audit_path.exists():
        return set()

    workbook_mtime = datetime.fromtimestamp(workbook.stat().st_mtime)
    try:
        xls = pd.ExcelFile(manual_audit_path)
        sheet_name = "Manual Changes" if "Manual Changes" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(manual_audit_path, sheet_name=sheet_name, dtype=str).fillna("")
    except Exception:
        return set()

    if "Resolution Status" not in df.columns or "Issue Identity" not in df.columns:
        return set()

    resolved: set[str] = set()
    for _, row in df.iterrows():
        if str(row.get("Resolution Status", "")).strip().lower() != "resolved":
            continue
        identity = str(row.get("Issue Identity", "")).strip()
        if not identity:
            continue
        timestamp_text = str(row.get("Timestamp", "")).strip()
        try:
            logged_at = datetime.fromisoformat(timestamp_text)
        except Exception:
            continue
        if logged_at >= workbook_mtime:
            resolved.add(identity.lower())
    return resolved
