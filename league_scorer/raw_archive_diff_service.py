from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from difflib import SequenceMatcher
import csv
from html.parser import HTMLParser

import pandas as pd

from .input_layout import RACE_FILE_SUFFIXES, build_input_paths


@dataclass(frozen=True)
class ComparableFilePair:
    filename: str
    raw_path: Path
    archive_path: Path


@dataclass(frozen=True)
class DiffRow:
    left_line_no: int | None
    right_line_no: int | None
    left_text: str
    right_text: str
    status: str


def list_comparable_file_pairs(input_dir: Path) -> list[ComparableFilePair]:
    paths = build_input_paths(input_dir)
    if not paths.raw_data_dir.exists() or not paths.raw_data_archive_dir.exists():
        return []

    raw_files = {
        path.name: path
        for path in sorted(paths.raw_data_dir.iterdir())
        if path.is_file() and path.suffix.lower() in RACE_FILE_SUFFIXES
    }
    archive_files = {
        path.name: path
        for path in sorted(paths.raw_data_archive_dir.iterdir())
        if path.is_file() and path.suffix.lower() in RACE_FILE_SUFFIXES
    }

    pairs: list[ComparableFilePair] = []
    for filename in sorted(set(raw_files) & set(archive_files), key=str.lower):
        pairs.append(
            ComparableFilePair(
                filename=filename,
                raw_path=raw_files[filename],
                archive_path=archive_files[filename],
            )
        )
    return pairs


def load_comparable_lines(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _load_csv_lines(path)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return _load_excel_lines(path)
    return path.read_text(encoding="utf-8", errors="replace").splitlines()


def build_side_by_side_diff(left_lines: list[str], right_lines: list[str]) -> list[DiffRow]:
    matcher = SequenceMatcher(a=left_lines, b=right_lines)
    rows: list[DiffRow] = []
    left_no = 1
    right_no = 1

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for left_text, right_text in zip(left_lines[i1:i2], right_lines[j1:j2]):
                rows.append(DiffRow(left_no, right_no, left_text, right_text, "same"))
                left_no += 1
                right_no += 1
            continue

        if tag == "replace":
            block_len = max(i2 - i1, j2 - j1)
            left_block = left_lines[i1:i2]
            right_block = right_lines[j1:j2]
            for index in range(block_len):
                left_text = left_block[index] if index < len(left_block) else ""
                right_text = right_block[index] if index < len(right_block) else ""
                current_left_no = left_no if index < len(left_block) else None
                current_right_no = right_no if index < len(right_block) else None
                rows.append(DiffRow(current_left_no, current_right_no, left_text, right_text, "replace"))
                if index < len(left_block):
                    left_no += 1
                if index < len(right_block):
                    right_no += 1
            continue

        if tag == "delete":
            for left_text in left_lines[i1:i2]:
                rows.append(DiffRow(left_no, None, left_text, "", "delete"))
                left_no += 1
            continue

        if tag == "insert":
            for right_text in right_lines[j1:j2]:
                rows.append(DiffRow(None, right_no, "", right_text, "insert"))
                right_no += 1

    return rows


def _load_csv_lines(path: Path) -> list[str]:
    lines: list[str] = []
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
        reader = csv.reader(handle)
        for row in reader:
            lines.append(" | ".join(_normalise_cell(value) for value in row))
    return lines


def _load_excel_lines(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        engine = "xlrd"
    else:
        engine = "openpyxl"

    try:
        excel_file = pd.ExcelFile(path, engine=engine)
    except Exception:
        if suffix == ".xls":
            return _load_text_spreadsheet_lines(path)
        raise

    lines: list[str] = []
    for sheet_name in excel_file.sheet_names:
        df = excel_file.parse(sheet_name, dtype=str).fillna("")
        lines.append(f"=== Sheet: {sheet_name} ===")
        lines.append(" | ".join(_normalise_cell(column) for column in df.columns.tolist()))
        for _, row in df.iterrows():
            lines.append(" | ".join(_normalise_cell(value) for value in row.tolist()))
    return lines


def _load_text_spreadsheet_lines(path: Path) -> list[str]:
    raw_bytes = path.read_bytes()
    text = _decode_text_spreadsheet_bytes(raw_bytes)

    parsed_lines = _parse_html_table_lines(text)
    if parsed_lines:
        return parsed_lines

    return [line.strip() for line in text.splitlines() if line.strip()]


def _decode_text_spreadsheet_bytes(raw_bytes: bytes) -> str:
    if raw_bytes.startswith(b"\xff\xfe"):
        return raw_bytes.decode("utf-16", errors="replace")
    if raw_bytes.startswith(b"\xfe\xff"):
        return raw_bytes.decode("utf-16", errors="replace")
    return raw_bytes.decode("utf-8-sig", errors="replace")


def _parse_html_table_lines(text: str) -> list[str]:
    parser = _HtmlTableParser()
    try:
        parser.feed(text)
        parser.close()
    except Exception:
        return []

    if not parser.tables:
        return []

    lines: list[str] = []
    for table_index, table in enumerate(parser.tables, start=1):
        if not table:
            continue
        lines.append(f"=== Table: {table_index} ===")
        header = table[0]
        lines.append(" | ".join(_normalise_cell(value) for value in header))
        for row in table[1:]:
            lines.append(" | ".join(_normalise_cell(value) for value in row))
    return lines


class _HtmlTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._in_table = False
        self._in_row = False
        self._in_cell = False
        self._current_table: list[list[str]] | None = None
        self._current_row: list[str] | None = None
        self._cell_chunks: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        if tag == "table":
            self._in_table = True
            self._current_table = []
        elif self._in_table and tag == "tr":
            self._in_row = True
            self._current_row = []
        elif self._in_row and tag in {"td", "th"}:
            self._in_cell = True
            self._cell_chunks = []

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._cell_chunks.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"} and self._in_cell:
            value = " ".join(chunk.strip() for chunk in self._cell_chunks if chunk.strip())
            if self._current_row is not None:
                self._current_row.append(value)
            self._in_cell = False
            self._cell_chunks = []
        elif tag == "tr" and self._in_row:
            if self._current_table is not None and self._current_row is not None:
                if any(cell for cell in self._current_row):
                    self._current_table.append(self._current_row)
            self._current_row = None
            self._in_row = False
        elif tag == "table" and self._in_table:
            if self._current_table:
                self.tables.append(self._current_table)
            self._current_table = None
            self._in_table = False


def _normalise_cell(value: object) -> str:
    text = str(value).replace("\r\n", " ").replace("\n", " ").strip()
    return text