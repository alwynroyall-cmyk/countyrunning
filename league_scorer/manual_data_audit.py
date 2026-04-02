"""Utilities for logging manual data edits to a season audit workbook."""

from __future__ import annotations

import getpass
import os
import socket
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

from .audit_writer import write_audit_workbook
from .session_config import config as session_config

_SHEET_NAME = "Manual Changes"
_FILENAME = "Manual_Data_Audit.xlsx"


def log_manual_data_changes(
    changes: Iterable[dict],
    *,
    source: str,
    action: str,
) -> str | None:
    """Append manual-change rows to outputs/audit/Manual_Data_Audit.xlsx.

    Returns an error string when logging fails, otherwise None.
    """
    rows = list(changes)
    if not rows:
        return None

    output_dir = session_config.output_dir
    if output_dir is None:
        return "Output directory is not configured for this session."

    audit_path = output_dir / "audit" / _FILENAME
    timestamp = datetime.now().isoformat(timespec="seconds")
    change_date = datetime.now().date().isoformat()
    user_name = getpass.getuser()
    machine_id = _machine_id()

    normalized_rows = []
    for row in rows:
        normalized_rows.append(
            {
                "Timestamp": timestamp,
                "Date of Change": change_date,
                "User": user_name,
                "Author/User": user_name,
                "Machine ID": machine_id,
                "Source": source,
                "Action": action,
                "Runner": str(row.get("runner", "") or ""),
                "Field": str(row.get("field", "") or ""),
                "Old Value": str(row.get("old_value", "") or ""),
                "New Value": str(row.get("new_value", "") or ""),
                "File": _safe_rel_path(row.get("file_path"), output_dir),
                "Row": row.get("row_idx", ""),
                "Issue Identity": str(row.get("issue_identity", "") or ""),
                "Issue Code": str(row.get("issue_code", "") or ""),
                "Issue Race": str(row.get("issue_race", "") or ""),
                "Issue Source Row": str(row.get("issue_source_row", "") or ""),
                "Resolution Status": str(row.get("resolution_status", "") or ""),
                "Verification Note": str(row.get("verification_note", "") or ""),
                "Season": session_config.year,
            }
        )

    new_df = pd.DataFrame(normalized_rows)
    try:
        existing_df = _read_existing(audit_path)
        combined = pd.concat([existing_df, new_df], ignore_index=True)
        write_audit_workbook({_SHEET_NAME: combined}, audit_path)
        return None
    except Exception as exc:  # pragma: no cover - surfaced to UI callers
        return str(exc)


def _read_existing(audit_path: Path) -> pd.DataFrame:
    if not audit_path.exists():
        return pd.DataFrame()

    try:
        xls = pd.ExcelFile(audit_path)
        if _SHEET_NAME in xls.sheet_names:
            return pd.read_excel(audit_path, sheet_name=_SHEET_NAME, dtype=str).fillna("")
        if xls.sheet_names:
            return pd.read_excel(audit_path, sheet_name=xls.sheet_names[0], dtype=str).fillna("")
    except Exception:
        pass
    return pd.DataFrame()


def _safe_rel_path(file_path: object, output_dir: Path) -> str:
    if not file_path:
        return ""
    path = Path(str(file_path))
    try:
        return str(path.relative_to(output_dir.parent))
    except Exception:
        return str(path)


def _machine_id() -> str:
    computername = os.getenv("COMPUTERNAME", "").strip()
    if computername:
        return computername
    return socket.gethostname().strip()
