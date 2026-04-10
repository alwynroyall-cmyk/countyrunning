from __future__ import annotations

import json
from argparse import Namespace
from pathlib import Path

from scripts import run_full_autopilot as autopilot
from scripts import run_provisional_fast_track as provisional


def test_run_provisional_fast_track_missing_data_root_writes_report(monkeypatch, tmp_path: Path):
    args = Namespace(year=1999, data_root=None, report_dir=tmp_path / "reports")
    monkeypatch.setattr(provisional, "parse_args", lambda: args)
    monkeypatch.setattr(provisional, "_resolve_data_root", lambda explicit: None)

    result = provisional.main()

    assert result == 1
    report_root = tmp_path / "reports" / "year-1999"
    payload = json.loads((report_root / "provisional_fast_track.json").read_text(encoding="utf-8"))
    assert payload["success"] is False
    assert "No data root configured" in payload["error"]["message"]


def test_run_full_autopilot_no_raw_files_writes_report(monkeypatch, tmp_path: Path):
    data_root = tmp_path / "data"
    input_dir = data_root / "1999" / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)

    args = Namespace(
        year=1999,
        data_root=data_root,
        report_dir=tmp_path / "reports",
        staged_report_dir=tmp_path / "staged",
        quality_gate_threshold=80.0,
        baseline_file=tmp_path / "baseline.json",
        write_baseline=False,
        allow_missing_data=False,
        mode="dry-run",
        max_fix_attempts=0,
    )
    monkeypatch.setattr(autopilot, "parse_args", lambda: args)
    monkeypatch.setattr(autopilot, "_generate_audited_race_files", lambda input_dir, overwrite_existing: {})
    monkeypatch.setattr(autopilot, "_run_audit_snapshot", lambda *args, **kwargs: None)
    monkeypatch.setattr(autopilot, "_print_preflight_summary", lambda input_dir: None)
    monkeypatch.setattr(autopilot, "discover_race_files", lambda *_args, **_kwargs: {})

    result = autopilot.main()

    assert result == 1
    report_root = tmp_path / "reports" / "year-1999"
    payload = json.loads((report_root / "autopilot_report.json").read_text(encoding="utf-8"))
    assert payload["success"] is False
    assert "No raw race files found" in payload["error"]["message"]
