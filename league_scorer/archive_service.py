"""Write-once archival helpers for raw race input files."""

from __future__ import annotations

from pathlib import Path
import shutil

from .input_layout import build_input_paths


def ensure_archived(filepath: Path, archive_dir: Path) -> Path | None:
    """Copy filepath into archive_dir if a same-name file is not already archived."""
    source = Path(filepath)
    archive_dir = Path(archive_dir)
    archive_dir.mkdir(parents=True, exist_ok=True)

    target = archive_dir / source.name
    if target.exists():
        return None

    shutil.copy2(source, target)
    return target


def ensure_archived_in_inputs(filepath: Path, input_dir: Path) -> Path | None:
    """Archive filepath to inputs/raw_data_archive using filename-only keying."""
    paths = build_input_paths(input_dir)
    return ensure_archived(filepath, paths.raw_data_archive_dir)
