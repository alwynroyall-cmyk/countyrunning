"""Package publish artifacts into a single deployable folder.

This wrapper copies the structured publish outputs into
`outputs/publish/package/`, flattening them into a single top-level package
by default. It can optionally copy the package to another destination or
create a zip archive for deployment.
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path
import sys

_repo_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_repo_root))

from league_scorer.output.output_layout import package_publish_artifacts  # noqa: E402
from scripts.autopilot.run_full_autopilot import _resolve_data_root, _season_paths  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Package publish artifacts for deployment.")
    parser.add_argument("--year", type=int, default=1999)
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument(
        "--dest",
        type=Path,
        default=None,
        help="Optional folder to copy the packaged publish files into.",
    )
    parser.add_argument(
        "--zip",
        type=Path,
        default=None,
        help="Optional zip file path to create from the packaged publish files.",
    )
    parser.add_argument(
        "--no-flatten",
        action="store_true",
        help="Preserve nested publish directories under the package folder instead of flattening files into a single top-level package.",
    )
    return parser.parse_args()


def _copy_tree(source: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    for item in sorted(source.rglob("*")):
        if item.is_dir():
            continue
        target = destination / item.relative_to(source)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(item), str(target))


def main() -> int:
    args = parse_args()
    data_root_resolved = _resolve_data_root(args.data_root)
    if data_root_resolved is None:
        print("No data root configured. Set Data Root before packaging publish artifacts.")
        return 1

    input_dir, output_dir = _season_paths(data_root_resolved, args.year)
    package_dir = package_publish_artifacts(output_dir, flatten=not args.no_flatten)
    print(f"Packaged publish artifacts into: {package_dir}")

    if args.dest:
        destination = args.dest
        _copy_tree(package_dir, destination)
        print(f"Copied package files to: {destination}")

    if args.zip:
        zip_path = args.zip.with_suffix("")
        archive_path = shutil.make_archive(str(zip_path), "zip", root_dir=str(package_dir))
        print(f"Created zip archive: {archive_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
