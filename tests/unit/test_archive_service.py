from pathlib import Path

from league_scorer.archive_service import ensure_archived, ensure_archived_in_inputs


def test_ensure_archived_copies_new_file(tmp_path: Path) -> None:
    source = tmp_path / "source.txt"
    source.write_text("data", encoding="utf-8")
    archive_dir = tmp_path / "archive"

    target = ensure_archived(source, archive_dir)

    assert target == archive_dir / source.name
    assert target.exists()
    assert target.read_text(encoding="utf-8") == "data"


def test_ensure_archived_returns_none_when_existing(tmp_path: Path) -> None:
    source = tmp_path / "source.txt"
    source.write_text("data", encoding="utf-8")
    archive_dir = tmp_path / "archive"
    existing = archive_dir / source.name
    archive_dir.mkdir(parents=True)
    existing.write_text("existing", encoding="utf-8")

    result = ensure_archived(source, archive_dir)

    assert result is None
    assert existing.read_text(encoding="utf-8") == "existing"


def test_ensure_archived_in_inputs_archives_to_raw_data_archive(tmp_path: Path) -> None:
    input_dir = tmp_path / "inputs"
    raw_data_archive_dir = input_dir / "raw_data_archive"
    source = tmp_path / "source.txt"
    source.write_text("data", encoding="utf-8")

    target = ensure_archived_in_inputs(source, input_dir)

    assert target == raw_data_archive_dir / source.name
    assert target.exists()
