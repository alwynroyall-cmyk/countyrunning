from pathlib import Path

import pytest

from league_scorer.main import LeagueScorer
from league_scorer.exceptions import FatalError


def test_validate_paths_accepts_expected_layout(tmp_path: Path) -> None:
    year_dir = tmp_path / "1999"
    input_dir = year_dir / "inputs"
    output_dir = year_dir / "outputs"
    input_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)

    scorer = LeagueScorer(input_dir, output_dir, 1999)
    scorer._validate_paths()


def test_validate_paths_rejects_wrong_input_folder(tmp_path: Path) -> None:
    year_dir = tmp_path / "1999"
    input_dir = year_dir / "data"
    output_dir = year_dir / "outputs"
    input_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)

    scorer = LeagueScorer(input_dir, output_dir, 1999)
    with pytest.raises(FatalError):
        scorer._validate_paths()


def test_validate_paths_rejects_wrong_year_folder(tmp_path: Path) -> None:
    input_dir = tmp_path / "1998" / "inputs"
    output_dir = tmp_path / "1999" / "outputs"
    input_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True)

    scorer = LeagueScorer(input_dir, output_dir, 1999)
    with pytest.raises(FatalError):
        scorer._validate_paths()
