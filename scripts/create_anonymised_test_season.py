"""Create a deterministic anonymised test season from a real season input set.

Default usage copies:
  data/2025/inputs -> data/1999/inputs

The same real runner name is mapped to the same anonymised name across all files.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from io import StringIO
import pandas as pd
from pathlib import Path

FIRST_NAMES = [
    "Alex", "Avery", "Bailey", "Blake", "Casey", "Charlie", "Cory", "Dale", "Devon", "Drew",
    "Eden", "Elliot", "Emery", "Finley", "Frankie", "Georgie", "Harley", "Hayden", "Indy", "Jaden",
    "Jamie", "Jesse", "Jordan", "Jules", "Kai", "Kendall", "Lane", "Logan", "Marley", "Morgan",
    "Nico", "Noel", "Oakley", "Parker", "Quinn", "Reese", "Riley", "River", "Robin", "Rowan",
    "Sage", "Sam", "Sawyer", "Sky", "Sydney", "Taylor", "Toby", "Winter", "Zion", "Zephyr",
]

LAST_NAMES = [
    "Adair", "Alden", "Arden", "Ashby", "Aster", "Barden", "Blythe", "Briar", "Caelan", "Caldwell",
    "Carey", "Carlisle", "Corwin", "Darby", "Dawson", "Ellery", "Ellis", "Emerson", "Farrow", "Fletcher",
    "Hadley", "Harlow", "Harper", "Hollis", "Ives", "Jensen", "Keaton", "Kendrick", "Kingsley", "Lennox",
    "Marlow", "Merrick", "Monroe", "Nolan", "Oakley", "Palmer", "Peyton", "Quincy", "Ramsey", "Reagan",
    "Ridley", "Sawyer", "Shaw", "Sutton", "Teagan", "Vale", "Vaughn", "Winslow", "Yardley", "Zeller",
]

FULL_NAME_HEADERS = {
    "name", "runner", "runner name", "raw name", "athlete", "athlete name",
}
FIRST_NAME_HEADERS = {"first", "first name", "forename", "given name"}
LAST_NAME_HEADERS = {"last", "last name", "surname", "family name"}


class NameMapper:
    def __init__(self, seed: str) -> None:
        self.seed = seed

    def _digest(self, key: str) -> bytes:
        return hashlib.sha256(f"{self.seed}|{key}".encode("utf-8")).digest()

    def map_name(self, raw_name: str) -> tuple[str, str, str]:
        key = normalize_name_key(raw_name)
        if not key:
            return "", "", ""

        digest = self._digest(key)
        first = FIRST_NAMES[digest[0] % len(FIRST_NAMES)]
        last = LAST_NAMES[digest[1] % len(LAST_NAMES)]
        suffix = f"{int.from_bytes(digest[2:4], 'big') % 1000:03d}"

        anon_last = f"{last}{suffix}"
        full = f"{first} {anon_last}"
        return first, anon_last, full


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().split())


def normalize_name_key(value: object) -> str:
    if value is None:
        return ""
    text = " ".join(str(value).strip().split())
    if not text:
        return ""
    return text.lower()


def _normalise_cell_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    return text


def _anonymise_sheet(df: pd.DataFrame, name_mapper: NameMapper) -> tuple[pd.DataFrame, int]:
    renamed_names = 0

    out = df.copy()
    headers = [normalize_header(col) for col in out.columns]
    full_name_cols = [col for col, h in zip(out.columns, headers) if h in FULL_NAME_HEADERS]
    first_col = next((col for col, h in zip(out.columns, headers) if h in FIRST_NAME_HEADERS), None)
    last_col = next((col for col, h in zip(out.columns, headers) if h in LAST_NAME_HEADERS), None)

    if first_col is not None and last_col is not None:
        for idx in out.index:
            first_val = _normalise_cell_text(out.at[idx, first_col])
            last_val = _normalise_cell_text(out.at[idx, last_col])
            combined = f"{first_val} {last_val}".strip()
            if not combined:
                continue
            anon_first, anon_last, _ = name_mapper.map_name(combined)
            if anon_first and anon_last:
                out.at[idx, first_col] = anon_first
                out.at[idx, last_col] = anon_last
                renamed_names += 1

    for col in full_name_cols:
        for idx in out.index:
            original_text = _normalise_cell_text(out.at[idx, col])
            if not original_text:
                continue
            _, _, anon_full = name_mapper.map_name(original_text)
            if anon_full:
                out.at[idx, col] = anon_full
                renamed_names += 1

    return out, renamed_names


def _read_excel_like(source_path: Path) -> dict[str, pd.DataFrame]:
    suffix = source_path.suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(source_path, sheet_name=None, dtype=object, engine="openpyxl")

    if suffix == ".xls":
        try:
            return pd.read_excel(source_path, sheet_name=None, dtype=object, engine="xlrd")
        except Exception:
            for encoding in ("utf-16", "utf-8-sig", "utf-8"):
                try:
                    df = pd.read_csv(source_path, sep="\t", dtype=object, encoding=encoding)
                    return {"Sheet1": df}
                except Exception:
                    continue

            for encoding in ("utf-16", "utf-8-sig", "utf-8"):
                try:
                    text = source_path.read_text(encoding=encoding)
                except Exception:
                    continue

                try:
                    tables = pd.read_html(StringIO(text))
                    if tables:
                        return {"Sheet1": tables[0]}
                except Exception:
                    pass

                try:
                    df = pd.read_csv(StringIO(text), sep=None, engine="python", dtype=object)
                    return {"Sheet1": df}
                except Exception:
                    pass

            raise

    raise ValueError(f"Unsupported workbook suffix: {source_path.suffix}")


def anonymise_workbook(
    source_path: Path,
    target_path: Path,
    name_mapper: NameMapper,
 ) -> int:
    all_sheets = _read_excel_like(source_path)
    renamed_names = 0

    target_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
        for sheet_name, df in all_sheets.items():
            out_df, n_names = _anonymise_sheet(df, name_mapper)
            out_df.to_excel(writer, sheet_name=sheet_name, index=False)
            renamed_names += n_names

    return renamed_names


def copy_and_anonymise_inputs(
    source_input_dir: Path,
    target_input_dir: Path,
    name_mapper: NameMapper,
    clear_target: bool,
) -> None:
    if not source_input_dir.exists():
        raise FileNotFoundError(f"Source input folder not found: {source_input_dir}")

    if clear_target and target_input_dir.exists():
        shutil.rmtree(target_input_dir)

    target_input_dir.mkdir(parents=True, exist_ok=True)

    total_files = 0
    workbook_files = 0
    total_names = 0

    for source_file in sorted(source_input_dir.rglob("*")):
        if source_file.is_dir():
            continue

        relative = source_file.relative_to(source_input_dir)
        total_files += 1
        target_file = target_input_dir / relative
        target_file.parent.mkdir(parents=True, exist_ok=True)

        if source_file.name.lower() == "clubs.xlsx":
            shutil.copy2(source_file, target_file)
            continue

        if source_file.suffix.lower() in {".xlsx", ".xls"}:
            if source_file.suffix.lower() == ".xls":
                target_file = target_file.with_suffix(".xlsx")
            name_count = anonymise_workbook(source_file, target_file, name_mapper)
            workbook_files += 1
            total_names += name_count
        else:
            shutil.copy2(source_file, target_file)

    print(f"Copied {total_files} files ({workbook_files} workbook files anonymised).")
    print(f"Anonymised name-cell updates: {total_names}")
    print(f"Target input folder: {target_input_dir}")


def ensure_target_output_dir(target_output_dir: Path) -> None:
    target_output_dir.mkdir(parents=True, exist_ok=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create deterministic anonymised season test data.")
    parser.add_argument("--source-year", type=int, default=2025, help="Source season year (default: 2025)")
    parser.add_argument("--target-year", type=int, default=1999, help="Target test season year (default: 1999)")
    parser.add_argument(
        "--data-root",
        type=Path,
        default=None,
        help="Data root containing year folders (default: auto-detect from prefs, then ./data)",
    )
    parser.add_argument(
        "--seed",
        type=str,
        default="wrrl-back-to-the-future",
        help="Deterministic anonymisation seed",
    )
    parser.add_argument(
        "--keep-target",
        action="store_true",
        help="Keep existing target inputs and overwrite files in place",
    )
    return parser.parse_args()


def resolve_data_root(explicit: Path | None) -> Path:
    if explicit is not None:
        return explicit

    prefs_path = Path.home() / ".wrrl_prefs.json"
    if prefs_path.exists():
        try:
            data = json.loads(prefs_path.read_text(encoding="utf-8"))
            data_root = data.get("data_root")
            if data_root:
                return Path(data_root)
        except Exception:
            pass

    return Path(__file__).resolve().parents[1] / "data"


def main() -> None:
    args = parse_args()
    data_root = resolve_data_root(args.data_root)

    source_input_dir = data_root / str(args.source_year) / "inputs"
    target_root = data_root / str(args.target_year)
    target_input_dir = target_root / "inputs"
    target_output_dir = target_root / "outputs"

    name_mapper = NameMapper(seed=args.seed)
    copy_and_anonymise_inputs(
        source_input_dir=source_input_dir,
        target_input_dir=target_input_dir,
        name_mapper=name_mapper,
        clear_target=not args.keep_target,
    )
    ensure_target_output_dir(target_output_dir)

    print(f"Target output folder ensured: {target_output_dir}")
    print(f"Data root used: {data_root}")
    print("Done.")


if __name__ == "__main__":
    main()
