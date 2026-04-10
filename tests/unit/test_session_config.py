from pathlib import Path

from league_scorer.session_config import SessionConfig


def test_session_config_derived_paths_and_year(tmp_path: Path) -> None:
    config = SessionConfig()
    config._data_root = tmp_path
    config._year = 2026

    assert config.input_dir == tmp_path / "2026" / "inputs"
    assert config.output_dir == tmp_path / "2026" / "outputs"
    assert config.raw_data_dir == tmp_path / "2026" / "inputs" / "raw_data"
    assert config.control_dir == tmp_path / "2026" / "inputs" / "control"
    assert config.audited_dir == tmp_path / "2026" / "inputs" / "audited"


def test_session_config_available_years_returns_range():
    years = SessionConfig.available_years()
    assert years[0] == 2020
    assert years[-1] >= 2026
