from __future__ import annotations

import tkinter as tk
from pathlib import Path

from league_scorer.graphical import gui
import league_scorer.config.session_config as session_config_module


class DummyConfig:
    def __init__(self, output_dir: Path | None) -> None:
        self.output_dir = output_dir


def test_dirty_indicator_shows_red_when_dirty(monkeypatch, tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    (output_dir / "autopilot").mkdir(parents=True, exist_ok=True)
    (output_dir / "autopilot" / "dirty").write_text("dirty", encoding="utf-8")

    monkeypatch.setattr(session_config_module, "config", DummyConfig(output_dir=output_dir))

    root = tk.Tk()
    try:
        app = gui.LeagueScorerApp(root, input_dir=tmp_path / "inputs", output_dir=output_dir, year=2025)
        app._update_dirty_indicator()
        assert app._dirty_indicator.cget("fg") == "red"
        assert app._run_btn.cget("bg") == gui.AMBER
    finally:
        root.destroy()


def test_dirty_indicator_shows_green_when_clean(monkeypatch, tmp_path: Path) -> None:
    output_dir = tmp_path / "outputs"
    output_dir.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(session_config_module, "config", DummyConfig(output_dir=output_dir))

    root = tk.Tk()
    try:
        app = gui.LeagueScorerApp(root, input_dir=tmp_path / "inputs", output_dir=output_dir, year=2025)
        app._update_dirty_indicator()
        assert app._dirty_indicator.cget("fg") == "green"
        assert app._run_btn.cget("bg") == gui.GREEN
    finally:
        root.destroy()
