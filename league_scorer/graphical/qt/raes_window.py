from __future__ import annotations

import datetime
import json
import os
import queue
import threading
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
import pandas as pd
from PySide6.QtCore import QTimer, Qt, Slot
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QTextEdit,
)

from ...manual_edit_service import _find_columns, _row_name_value
from ...raes.raes_write_service import apply_field_to_files, find_candidate_source_files
from ..results_workbook import find_latest_results_workbook, sorted_race_sheet_names
from ...session_config import config as session_config
from ...structured_logging import log_event

RAES_LIGHT = "#f5f5f5"
RAES_NAVY = "#3a4658"
RAES_GREEN = "#2d7a4a"


def proper_case(s: str) -> str:
    import re

    return re.sub(
        r"([A-Za-z])([A-Za-z']*)",
        lambda m: m.group(1).upper() + m.group(2).lower(),
        s,
    )


def _processed_state_path() -> Path | None:
    out = session_config.output_dir
    if out is None:
        return None
    d = Path(out) / "raes"
    d.mkdir(parents=True, exist_ok=True)
    return d / "processed_state.json"


def load_processed_state() -> Dict[str, bool]:
    p = _processed_state_path()
    if p is None or not p.exists():
        return {}
    try:
        with open(p, "r", encoding="utf-8") as fh:
            data = json.load(fh)
            if isinstance(data, dict):
                return {str(k): bool(v) for k, v in data.items()}
    except Exception:
        return {}
    return {}


def save_processed_state(state: Dict[str, bool]) -> None:
    p = _processed_state_path()
    if p is None:
        return
    try:
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(state, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


def set_runner_processed(runner: str, processed: bool) -> None:
    st = load_processed_state()
    st[runner] = bool(processed)
    save_processed_state(st)


def _scan_workbook_for_runner_state() -> Dict[str, Dict[str, set]]:
    if session_config.output_dir is None:
        return {}

    workbook = find_latest_results_workbook(session_config.output_dir)
    if workbook is None:
        return {}

    runner_state: Dict[str, Dict[str, set]] = {}
    try:
        xl = pd.ExcelFile(workbook)
        for sheet in sorted_race_sheet_names(xl):
            df = xl.parse(sheet).fillna("")
            if "Name" not in df.columns:
                continue
            for _, row in df.iterrows():
                name = str(row.get("Name", "")).strip()
                if not name:
                    continue
                norm = name.lower()
                state = runner_state.setdefault(
                    norm,
                    {
                        "club": set(),
                        "gender": set(),
                        "category": set(),
                        "raw_names": set(),
                    },
                )
                state["raw_names"].add(name)

                club = str(row.get("Club", "")).strip()
                if club:
                    state["club"].add(club)

                gender = str(row.get("Gender", "")).strip().upper()
                if gender:
                    state["gender"].add(gender)

                category = str(row.get("Category", "")).strip()
                if category:
                    state["category"].add(category)
    except Exception:
        return {}

    return runner_state


def _detect_runner_anomalies(runner_state: Dict[str, Dict[str, set]]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for norm, state in runner_state.items():
        flags: List[str] = []
        detail_parts: List[str] = []

        if len(state["club"]) > 1:
            flags.append("Club")
            detail_parts.append("clubs=" + ", ".join(sorted(state["club"])))

        if len(state["gender"]) > 1:
            flags.append("Gender")
            detail_parts.append("gender=" + ", ".join(sorted(state["gender"])))

        if len(state["category"]) > 1:
            flags.append("Category")
            preview = sorted(state["category"])
            detail_parts.append(
                "categories="
                + ", ".join(preview[:6])
                + ("..." if len(preview) > 6 else "")
            )

        if not flags:
            continue

        raw_name = next(iter(state["raw_names"]))
        rows.append(
            {
                "runner": proper_case(raw_name),
                "anomalies": "/".join(flags),
                "details": " | ".join(detail_parts),
                "processed": bool(load_processed_state().get(raw_name)),
            }
        )

    rows.sort(key=lambda item: item["runner"].lower())
    return rows


def _build_points_map() -> Optional[Dict[str, float]]:
    if session_config.output_dir is None:
        return None
    try:
        wb = find_latest_results_workbook(session_config.output_dir)
        if wb is None:
            return None
        points_map: Dict[str, float] = {}
        xl = pd.ExcelFile(wb)
        for sheet in xl.sheet_names:
            if str(sheet).strip().lower() not in ("male", "female"):
                continue
            try:
                df = xl.parse(sheet).fillna("")
            except Exception:
                continue
            if "Name" not in df.columns:
                continue
            for _, row in df.iterrows():
                name = str(row.get("Name", "")).strip()
                if not name:
                    continue
                norm = name.lower()
                pts = None
                for cname in row.index:
                    if cname and str(cname).lower() in ("total points", "total_points", "points", "pts"):
                        pts = row.get(cname)
                        break
                try:
                    pval = float(pts) if pts is not None and str(pts).strip() != "" else 0.0
                except Exception:
                    pval = 0.0
                points_map[norm] = max(points_map.get(norm, 0.0), pval)
        return points_map
    except Exception:
        return None


def build_raes_runner_rows(show_all: bool = False) -> List[Dict[str, object]]:
    runner_state = _scan_workbook_for_runner_state()
    if not runner_state:
        return []

    anomalies = _detect_runner_anomalies(runner_state)
    processed_map = load_processed_state()
    points_map = _build_points_map()

    rows: List[Dict[str, object]] = []
    for item in anomalies:
        name = item.get("runner", "")
        if points_map is not None and not show_all:
            if points_map.get(name.lower(), 0.0) <= 0.0:
                continue
        rows.append(
            {
                "runner": name,
                "anomalies": item.get("anomalies", ""),
                "details": item.get("details", ""),
                "processed": bool(processed_map.get(name)),
            }
        )

    rows.sort(key=lambda r: r["runner"].lower())
    return rows


class RAESWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("RAES Manual Corrections")
        self.setMinimumSize(980, 720)
        self._rows: List[Dict[str, object]] = []
        self._selected_runner: Optional[str] = None
        self._candidate_checkboxes: Dict[str, QCheckBox] = {}
        self._file_value_labels: Dict[str, QLabel] = {}
        self._candidate_paths: list[Path] = []
        self._source_value_cache: Dict[tuple[str, str, str], str] = {}
        self._scan_queue: Optional[queue.Queue] = None
        self._source_value_queue: Optional[queue.Queue] = None

        self._build_ui()
        self._refresh_list()
        self._start_dirty_timer()

    def _build_ui(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        title_row = QWidget(self)
        title_layout = QHBoxLayout(title_row)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(8)

        title = QLabel("RAES Manual Corrections", self)
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title.setStyleSheet(f"color: {RAES_NAVY};")
        title_layout.addWidget(title)
        title_layout.addStretch(1)

        refresh_btn = QPushButton("Refresh", self)
        refresh_btn.clicked.connect(self._refresh_list)
        title_layout.addWidget(refresh_btn)

        self._show_all_cb = QCheckBox("Show all runners (include non-league / zero-point runners)", self)
        self._show_all_cb.stateChanged.connect(self._refresh_list)
        title_layout.addWidget(self._show_all_cb)

        self._count_label = QLabel("Anomalies: 0", self)
        self._count_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        title_layout.addWidget(self._count_label)

        layout.addWidget(title_row)

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.setHandleWidth(4)
        splitter.setContentsMargins(0, 0, 0, 0)

        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        self._runner_tree = QTreeWidget(self)
        self._runner_tree.setHeaderLabels(["Runner", "Processed"])
        self._runner_tree.setRootIsDecorated(False)
        self._runner_tree.setSelectionBehavior(QTreeWidget.SelectRows)
        self._runner_tree.setSelectionMode(QTreeWidget.SingleSelection)
        self._runner_tree.itemSelectionChanged.connect(self._on_runner_selected)
        left_layout.addWidget(self._runner_tree)

        splitter.addWidget(left)

        right = QWidget(self)
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        self._summary_label = QLabel("Select a runner from the left to review anomalies.", self)
        self._summary_label.setWordWrap(True)
        self._summary_label.setStyleSheet(f"color: {RAES_NAVY};")
        self._summary_label.setFont(QFont("Segoe UI", 11))
        right_layout.addWidget(self._summary_label)

        self._source_group = QGroupBox("Source Files", self)
        self._source_group.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self._source_group.setMinimumHeight(240)
        source_layout = QVBoxLayout(self._source_group)
        source_layout.setContentsMargins(8, 8, 8, 8)
        source_layout.setSpacing(8)

        self._source_scroll = QScrollArea(self)
        self._source_scroll.setWidgetResizable(True)
        self._source_scroll.setMinimumHeight(220)
        self._source_widget = QWidget(self)
        self._source_widget.setLayout(QVBoxLayout())
        self._source_widget.layout().setContentsMargins(0, 0, 0, 0)
        self._source_widget.layout().setSpacing(8)
        self._source_scroll.setWidget(self._source_widget)
        source_layout.addWidget(self._source_scroll)

        right_layout.addWidget(self._source_group)

        action_row = QWidget(self)
        action_layout = QHBoxLayout(action_row)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(10)

        action_layout.addWidget(QLabel("Field:", self))
        self._field_combo = QComboBox(self)
        self._field_combo.addItems(["category", "club", "gender"])
        self._field_combo.currentTextChanged.connect(self._on_field_changed)
        action_layout.addWidget(self._field_combo)

        action_layout.addWidget(QLabel("Value:", self))
        self._value_combo = QComboBox(self)
        action_layout.addWidget(self._value_combo)

        apply_btn = QPushButton("Apply to selected files", self)
        apply_btn.clicked.connect(self._on_apply)
        action_layout.addWidget(apply_btn)

        mark_btn = QPushButton("Mark Reviewed", self)
        mark_btn.clicked.connect(self._on_mark_reviewed)
        action_layout.addWidget(mark_btn)

        action_layout.addStretch(1)
        right_layout.addWidget(action_row)

        diag_group = QGroupBox("Diagnostics", self)
        diag_layout = QVBoxLayout(diag_group)
        diag_layout.setContentsMargins(8, 8, 8, 8)
        self._diag_text = QTextEdit(self)
        self._diag_text.setReadOnly(True)
        self._diag_text.setStyleSheet("background: #f7f7f9; color: #333333;")
        diag_layout.addWidget(self._diag_text)
        right_layout.addWidget(diag_group, 1)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter)

        status_row = QWidget(self)
        status_layout = QHBoxLayout(status_row)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(12)

        self._last_updated_label = QLabel("Last updated: -", self)
        status_layout.addWidget(self._last_updated_label)
        self._workbook_label = QLabel("Workbook: -", self)
        status_layout.addWidget(self._workbook_label)
        status_layout.addStretch(1)
        self._dirty_label = QLabel("", self)
        self._dirty_label.setStyleSheet("color: #b22222;")
        status_layout.addWidget(self._dirty_label)

        layout.addWidget(status_row)

        self._field_combo.setCurrentText("category")
        self._populate_value_options("category")

    def _refresh_list(self) -> None:
        self._runner_tree.clear()
        self._count_label.setText("Scanning…")
        self._selected_runner = None
        self._summary_label.setText("Select a runner from the left to review anomalies.")
        self._clear_source_checkboxes()
        self._diag_text.clear()

        q: queue.Queue = queue.Queue()
        self._scan_queue = q

        def worker() -> None:
            try:
                rows = build_raes_runner_rows(self._show_all_cb.isChecked())
                workbook = find_latest_results_workbook(session_config.output_dir)
                q.put((rows, workbook))
            except Exception as exc:
                q.put((None, exc))

        threading.Thread(target=worker, daemon=True).start()
        QTimer.singleShot(100, self._poll_scan)

    def _poll_scan(self) -> None:
        if self._scan_queue is None:
            return
        try:
            payload = self._scan_queue.get_nowait()
        except queue.Empty:
            QTimer.singleShot(100, self._poll_scan)
            return

        self._scan_queue = None
        if payload is None or payload[0] is None and isinstance(payload[1], Exception):
            QMessageBox.critical(self, "RAES Error", str(payload[1] if payload else "Unknown scan error."))
            self._count_label.setText("Anomalies: 0")
            return

        rows, workbook = payload
        self._rows = rows or []
        for row in self._rows:
            item = QTreeWidgetItem([str(row.get("runner", "")), "✓" if row.get("processed") else ""])
            self._runner_tree.addTopLevelItem(item)

        self._count_label.setText(f"Anomalies: {len(self._rows)}")
        if workbook is not None:
            age = int((datetime.datetime.now() - datetime.datetime.fromtimestamp(Path(workbook).stat().st_mtime)).total_seconds() // 60)
            self._last_updated_label.setText(f"Last updated: {age} minute(s) ago")
            self._workbook_label.setText(f"Workbook: {Path(workbook).name}")
        else:
            self._last_updated_label.setText("Last updated: -")
            self._workbook_label.setText("Workbook: -")

    def _on_runner_selected(self) -> None:
        items = self._runner_tree.selectedItems()
        if not items:
            return
        runner = items[0].text(0)
        self._selected_runner = runner
        self._populate_runner_detail(runner)

    def _populate_runner_detail(self, runner: str) -> None:
        row = next((r for r in self._rows if r.get("runner") == runner), None)
        if row is None:
            return

        self._summary_label.setText(
            f"<b>{runner}</b><br><br>"
            f"Anomalies: {row.get('anomalies', '')}<br>"
            f"Details: {row.get('details', '')}<br>"
            f"Processed: {'Yes' if row.get('processed') else 'No'}"
        )
        self._summary_label.setTextFormat(Qt.RichText)

        self._populate_source_files(runner)
        self._on_field_changed(self._field_combo.currentText())
        self._refresh_diagnostics(runner)

    def _populate_source_files(self, runner: str) -> None:
        self._clear_source_checkboxes()
        try:
            candidates = find_candidate_source_files(runner)
            if not candidates.get("series") and not candidates.get("raw"):
                label = QLabel("No candidate source files found for this runner.", self._source_widget)
                self._source_widget.layout().addWidget(label)
                return

            def add_file_row(prefix: str, p: Path) -> None:
                row = QWidget(self._source_widget)
                row.setMinimumHeight(34)
                row_layout = QHBoxLayout(row)
                row_layout.setContentsMargins(0, 0, 0, 0)
                row_layout.setSpacing(10)

                checkbox = QCheckBox(f"{prefix}: {p.name}", row)
                checkbox.setChecked(False)
                checkbox.setFont(QFont("Segoe UI", 10))
                self._candidate_checkboxes[str(p)] = checkbox
                row_layout.addWidget(checkbox)

                value_lbl = QLabel("", row)
                value_lbl.setStyleSheet("color: #444444;")
                value_lbl.setFont(QFont("Segoe UI", 10, QFont.DemiBold))
                row_layout.addWidget(value_lbl)
                row_layout.addStretch(1)
                self._file_value_labels[str(p)] = value_lbl

                self._source_widget.layout().addWidget(row)

            self._candidate_paths = list(candidates.get("series", [])) + list(candidates.get("raw", []))
            for p in candidates.get("series", []):
                add_file_row("Series", p)
            for p in candidates.get("raw", []):
                add_file_row("Raw", p)

            self._refresh_source_file_values_async()
        except Exception:
            label = QLabel("Failed to load candidate source files.", self._source_widget)
            self._source_widget.layout().addWidget(label)

    def _clear_source_checkboxes(self) -> None:
        self._candidate_checkboxes.clear()
        self._file_value_labels.clear()
        layout = self._source_widget.layout()
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _populate_value_options(self, field: str) -> None:
        self._value_combo.clear()
        if self._selected_runner is None:
            return

        if field == "category":
            opts = ["Jun", "Sen", "V40", "V50", "V60", "V70"]
            self._value_combo.addItems(opts)
        elif field == "gender":
            self._value_combo.addItems(["Male", "Female"])
        else:
            opts_set: set[str] = set()
            for p in self._candidate_paths:
                try:
                    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                except Exception:
                    continue
                try:
                    for sname in wb.sheetnames:
                        try:
                            ws = wb[sname]
                            name_col, club_col = _find_columns(ws, "club")
                            if name_col is None or club_col is None:
                                continue
                            for ri in range(2, ws.max_row + 1):
                                nm = _row_name_value(ws, ri, name_col)
                                if nm is None or nm.strip().lower() != self._selected_runner.lower():
                                    continue
                                val = "" if ws.cell(row=ri, column=club_col).value is None else str(ws.cell(row=ri, column=club_col).value).strip()
                                if val:
                                    opts_set.add(val)
                        except Exception:
                            continue
                finally:
                    try:
                        wb.close()
                    except Exception:
                        pass
            opts = sorted(opts_set)
            if not opts:
                opts = [""]
            self._value_combo.addItems(opts)
        if self._value_combo.count() > 0:
            self._value_combo.setCurrentIndex(0)
        if self._selected_runner:
            self._refresh_source_file_values_async()

    def _on_field_changed(self, field: str) -> None:
        self._populate_value_options(field)
        if self._selected_runner:
            self._refresh_source_file_values_async()

    def _refresh_source_file_values_async(self) -> None:
        if not self._selected_runner:
            return
        runner = self._selected_runner
        field = self._field_combo.currentText()
        for label in self._file_value_labels.values():
            label.setText("[loading]")

        q = queue.Queue()
        self._source_value_queue = q
        paths = list(self._file_value_labels.keys())

        def worker() -> None:
            results: Dict[str, str] = {}
            for path_str in paths:
                try:
                    value = self._read_runner_value_for_file(Path(path_str), runner, field)
                except Exception:
                    value = "[error]"
                results[path_str] = value
            q.put((runner, field, results))

        threading.Thread(target=worker, daemon=True).start()
        QTimer.singleShot(100, self._poll_source_value_results)

    def _poll_source_value_results(self) -> None:
        if self._source_value_queue is None:
            return
        try:
            runner, field, results = self._source_value_queue.get_nowait()
        except queue.Empty:
            QTimer.singleShot(100, self._poll_source_value_results)
            return

        if runner != self._selected_runner or field != self._field_combo.currentText():
            return

        for path_str, label in self._file_value_labels.items():
            value = results.get(path_str, "")
            if value == "":
                label.setText("[missing]")
            elif value == "[error]":
                label.setText("[error]")
            else:
                label.setText(f"[{value}]")

    def _read_runner_value_for_file(self, path: Path, runner: str, field: str) -> str:
        cache_key = (runner, field, str(path))
        if cache_key in self._source_value_cache:
            return self._source_value_cache[cache_key]

        runner_key = runner.lower()
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception:
            self._source_value_cache[cache_key] = ""
            return ""

        try:
            for sname in wb.sheetnames:
                ws = wb[sname]
                name_col, field_col = _find_columns(ws, field)
                if name_col is None or field_col is None:
                    continue
                for ri in range(2, ws.max_row + 1):
                    nm = _row_name_value(ws, ri, name_col)
                    if nm is None or nm.strip().lower() != runner_key:
                        continue
                    cell = ws.cell(row=ri, column=field_col)
                    result = "" if cell.value is None else str(cell.value).strip()
                    self._source_value_cache[cache_key] = result
                    return result
        finally:
            try:
                wb.close()
            except Exception:
                pass

        self._source_value_cache[cache_key] = ""
        return ""

    def _on_apply(self) -> None:
        if self._selected_runner is None:
            QMessageBox.warning(self, "Apply", "Select a runner before applying changes.")
            return
        files = [Path(path) for path, cb in self._candidate_checkboxes.items() if cb.isChecked()]
        if not files:
            QMessageBox.warning(self, "Apply", "No source files selected. Please check at least one source file.")
            return
        field = self._field_combo.currentText()
        value = self._value_combo.currentText()
        if not field or not value:
            QMessageBox.warning(self, "Apply", "Field and value must be selected before applying.")
            return

        if not QMessageBox.question(
            self,
            "Apply Changes",
            f"Apply change: set {field} → {value} for {self._selected_runner} in {len(files)} file(s)?",
            QMessageBox.Yes | QMessageBox.No,
        ) == QMessageBox.Yes:
            return

        try:
            audit = apply_field_to_files(files, self._selected_runner, field, value)
        except Exception as exc:
            QMessageBox.critical(self, "Apply Error", f"Failed to apply changes: {exc}")
            return
        self._source_value_cache.clear()

        changed = [a for a in audit if a.get("runner")]
        errors = [a for a in audit if a.get("error")]
        msg = f"Applied edits: {len(changed)} change(s)."
        if errors:
            msg += f" {len(errors)} error(s) occurred. See RAES changes file for details."
        QMessageBox.information(self, "Apply Results", msg)
        set_runner_processed(self._selected_runner, True)
        self._mark_runner_processed(self._selected_runner)
        self._refresh_diagnostics(self._selected_runner)

    def _set_runner_item_processed(self, runner: str, processed: bool) -> None:
        for row in self._rows:
            if row.get("runner") == runner:
                row["processed"] = processed
                break
        for i in range(self._runner_tree.topLevelItemCount()):
            item = self._runner_tree.topLevelItem(i)
            if item.text(0) == runner:
                item.setText(1, "✓" if processed else "")
                break

    def _mark_runner_processed(self, runner: str) -> None:
        self._set_runner_item_processed(runner, True)
        if self._selected_runner == runner:
            self._summary_label.setText(
                self._summary_label.text().replace("Processed: No", "Processed: Yes")
            )

    def _on_mark_reviewed(self) -> None:
        if self._selected_runner is None:
            QMessageBox.warning(self, "Mark Reviewed", "Select a runner before marking reviewed.")
            return
        set_runner_processed(self._selected_runner, True)
        self._mark_runner_processed(self._selected_runner)
        log_event("raes_mark_reviewed", year=session_config.year, runner=self._selected_runner)

    def _refresh_diagnostics(self, runner: str) -> None:
        lines: List[str] = []
        changes_file = None
        if session_config.output_dir is not None:
            changes_file = Path(session_config.output_dir) / "raes" / "changes.json"
        recent_changes = []
        if changes_file is not None and changes_file.exists():
            try:
                recent_changes = json.loads(changes_file.read_text(encoding="utf-8"))
            except Exception:
                recent_changes = []
        rc_for_runner = [c for c in recent_changes if str(c.get("runner", "")).lower() == runner.lower()]
        if rc_for_runner:
            lines.append("Recent Changes:")
            for ch in rc_for_runner[-6:]:
                ts = ch.get("timestamp", "")
                file_name = Path(str(ch.get("file", ch.get("file_path", "")))).name
                fld = ch.get("field", "")
                old = ch.get("old_value", ch.get("old", ""))
                nv = ch.get("new_value", ch.get("new", ""))
                lines.append(f"{ts} — {file_name} — {fld}: {old} → {nv}")
        else:
            lines.append("Recent Changes: none")

        try:
            runner_state = _scan_workbook_for_runner_state()
            state = runner_state.get(runner.lower())
            if state:
                if len(state.get("club", set())) == 1:
                    lines.append(f"Club: consistent — {next(iter(state['club']))}")
                if len(state.get("gender", set())) == 1:
                    lines.append(f"Gender: consistent — {next(iter(state['gender']))}")
                if len(state.get("category", set())) == 1:
                    lines.append(f"Category: consistent — {next(iter(state['category']))}")
                if any(not v for v in [state.get("club"), state.get("gender"), state.get("category")]):
                    lines.append("Warning: Some fields may be missing in the latest results workbook.")
        except Exception:
            pass

        self._diag_text.setPlainText("\n".join(lines))

    def _start_dirty_timer(self) -> None:
        QTimer.singleShot(1000, self._update_dirty_indicator)

    def _update_dirty_indicator(self) -> None:
        is_dirty = False
        if session_config.output_dir is not None:
            ap_flag = Path(session_config.output_dir) / "autopilot" / "dirty"
            raes_flag = Path(session_config.output_dir) / "raes" / "dirty"
            is_dirty = ap_flag.exists() or raes_flag.exists()
        self._dirty_label.setText("DATA DIRTY — run Autopilot" if is_dirty else "")
        QTimer.singleShot(1000, self._update_dirty_indicator)
