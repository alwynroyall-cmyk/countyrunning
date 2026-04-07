import json
from pathlib import Path

import pytest

from league_scorer.raes import raes_service
from league_scorer.raes.raes_service import build_raes_runner_rows, set_runner_processed, load_processed_state


def test_processed_state_roundtrip(tmp_path, monkeypatch):
    # Point session_config.output_dir to tmp_path
    monkeypatch.setattr(raes_service, "session_config", type("C", (), {"output_dir": tmp_path}))
    # Ensure starting clean
    spath = tmp_path / "raes" / "processed_state.json"
    if spath.exists():
        spath.unlink()

    set_runner_processed("Alice Runner", True)
    st = load_processed_state()
    assert st.get("Alice Runner") is True


def test_build_rows_respects_anomalies_and_processed(tmp_path, monkeypatch):
    # Stub workbook scan and anomaly detection
    fake_runner_state = {
        "alice runner": {"raw_names": ["Alice Runner"], "club": ["Club A"], "gender": ["F"], "category": ["S"]}
    }
    monkeypatch.setattr(raes_service, "scan_workbook_for_runner_state", lambda: fake_runner_state)
    monkeypatch.setattr(raes_service, "detect_runner_anomalies", lambda s: [{"runner": "Alice Runner", "anomalies": "Club mismatch", "details": "Details"}])
    monkeypatch.setattr(raes_service, "session_config", type("C", (), {"output_dir": tmp_path}))

    # Mark processed and verify build_raes_runner_rows reflects it
    set_runner_processed("Alice Runner", True)
    rows = build_raes_runner_rows()
    assert isinstance(rows, list)
    assert any(r.get("runner") == "Alice Runner" and r.get("processed") for r in rows)
