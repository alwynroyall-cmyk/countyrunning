"""RAES helpers — service layer for the RAES manual correction UI.

This module reuses existing anomaly-detection helpers to build a simple
runner list with a processed flag. It also persists per-runner processed
state to the outputs folder.
"""
from __future__ import annotations

from typing import List, Dict
from pathlib import Path
import json

from ..session_config import config as session_config
from ..graphical.manual_review_helpers import (
    scan_workbook_for_runner_state,
    detect_runner_anomalies,
)
from ..graphical.results_workbook import find_latest_results_workbook
import pandas as pd


def build_raes_runner_rows(show_all: bool = False) -> List[Dict[str, object]]:
    """Return list of anomaly rows for the RAES left-hand list.

    Each row is a dict: {"runner": str, "anomalies": str, "details": str, "processed": bool}.
    Only runners that show anomalies are returned. `processed` is False
    by default and should be set by the UI when the runner is reviewed.
    """
    runner_state = scan_workbook_for_runner_state()
    if not runner_state:
        return []

    anomalies = detect_runner_anomalies(runner_state)

    # Load persisted processed flags and apply where available
    processed_map = load_processed_state()

    # If a latest results workbook exists, build a map of runner -> total points
    points_map: dict | None = None
    try:
        wb = find_latest_results_workbook(session_config.output_dir) if session_config.output_dir else None
        if wb is not None:
            points_map = {}
            xl = pd.ExcelFile(wb)
            for sheet in xl.sheet_names:
                if str(sheet).strip().lower() not in ("male", "female"):
                    continue
                try:
                    df = xl.parse(sheet).fillna("")
                except Exception:
                    continue
                if "Name" not in df.columns:
                    continue
                for _, row in df.iterrows():
                    name = str(row.get("Name", "")).strip()
                    if not name:
                        continue
                    norm = name.lower()
                    pts = None
                    for cname in row.index:
                        if cname and str(cname).lower() in ("total points", "total_points", "points", "pts"):
                            pts = row.get(cname)
                            break
                    try:
                        pval = float(pts) if pts is not None and str(pts).strip() != "" else 0.0
                    except Exception:
                        pval = 0.0
                    points_map[norm] = max(points_map.get(norm, 0.0), pval)
    except Exception:
        points_map = None

    rows: List[Dict[str, object]] = []
    for item in anomalies:
        name = item.get("runner", "")
        # If points_map is available and user has not requested show_all,
        # skip runners with zero or missing points.
        if points_map is not None and not show_all:
            if points_map.get(name.lower(), 0.0) <= 0.0:
                continue
        rows.append({
            "runner": name,
            "anomalies": item.get("anomalies", ""),
            "details": item.get("details", ""),
            "processed": bool(processed_map.get(name)),
        })

    rows.sort(key=lambda r: r["runner"].lower())
    return rows


def _processed_state_path() -> Path | None:
    out = session_config.output_dir
    if out is None:
        return None
    # Persist processed-state under outputs/<year>/raes/processed_state.json
    # so the RAES UI can remember which runners a reviewer has already
    # marked reviewed. This is intentionally lightweight JSON rather
    # than embedding state in workbooks.
    d = Path(out) / "raes"
    d.mkdir(parents=True, exist_ok=True)
    return d / "processed_state.json"


def load_processed_state() -> Dict[str, bool]:
    """Load processed-state mapping from outputs/raes/processed_state.json.

    Returns an empty dict if no persisted state exists or output dir not set.
    """
    p = _processed_state_path()
    if p is None or not p.exists():
        return {}
    try:
        with open(p, "r", encoding="utf-8") as fh:
            data = json.load(fh)
            if isinstance(data, dict):
                return {str(k): bool(v) for k, v in data.items()}
    except Exception:
        return {}
    return {}


def save_processed_state(state: Dict[str, bool]) -> None:
    p = _processed_state_path()
    if p is None:
        return
    try:
        # Write atomically in future if needed; for now a simple write
        # is sufficient since processed-state is small and non-critical.
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(state, fh, ensure_ascii=False, indent=2)
    except Exception:
        return


def set_runner_processed(runner: str, processed: bool) -> None:
    st = load_processed_state()
    st[runner] = bool(processed)
    save_processed_state(st)


def clear_processed_state() -> None:
    p = _processed_state_path()
    if p is None or not p.exists():
        return
    try:
        p.unlink()
    except Exception:
        return
