"""Qt window for viewing and copying runner history from the latest results workbook."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QGuiApplication
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from league_scorer.graphical.results_workbook import find_latest_results_workbook, sorted_race_sheet_names
from league_scorer.graphical.runner_history_helpers import (
    detect_conflicts,
    extract_runner_history,
    extract_runner_names,
    load_workbook_cache,
    proper_case,
    sanitise_df_for_export,
)
from league_scorer.race_processor import extract_race_number

WRRL_NAVY = "#3a4658"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"


class RunnerHistoryWindow(QMainWindow):
    def __init__(self, output_dir: Path | None = None) -> None:
        super().__init__()
        self._output_dir = output_dir
        self._cache = None
        self._current_df = None
        self.setWindowTitle("Runner / Club Enquiry")
        self.resize(1160, 760)
        self._build_ui()
        self._refresh_runner_list()
        self._refresh_club_list()

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setStyleSheet(f"background: {WRRL_LIGHT};")
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header = QLabel("Runner / Club Enquiry", central)
        header.setFont(QFont("Segoe UI", 17, QFont.Bold))
        header.setStyleSheet(f"color: {WRRL_NAVY};")
        layout.addWidget(header)

        self._tabs = QTabWidget(central)
        self._tabs.setStyleSheet("QTabBar::tab { height: 32px; padding: 6px 14px; }\n")
        layout.addWidget(self._tabs, 1)

        self._build_runner_tab()
        self._build_club_tab()
        self._build_club_races_tab()

    def _build_runner_tab(self) -> None:
        runner_tab = QWidget()
        runner_layout = QVBoxLayout(runner_tab)
        runner_layout.setContentsMargins(0, 0, 0, 0)
        runner_layout.setSpacing(10)

        controls = QWidget(runner_tab)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        label = QLabel("Runner:", controls)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(label)

        self._runner_combo = QComboBox(controls)
        self._runner_combo.setEditable(True)
        self._runner_combo.setInsertPolicy(QComboBox.NoInsert)
        self._runner_combo.setMinimumWidth(420)
        self._runner_combo.activated.connect(self._load_runner_history)
        controls_layout.addWidget(self._runner_combo)

        load_btn = QPushButton("Load", controls)
        load_btn.clicked.connect(self._load_runner_history)
        controls_layout.addWidget(load_btn)

        refresh_btn = QPushButton("Refresh", controls)
        refresh_btn.clicked.connect(self._refresh_runner_list)
        controls_layout.addWidget(refresh_btn)

        copy_btn = QPushButton("Copy Results", controls)
        copy_btn.clicked.connect(self._copy_results)
        controls_layout.addWidget(copy_btn)

        controls_layout.addStretch(1)
        runner_layout.addWidget(controls)

        self._summary_label = QLabel("", runner_tab)
        self._summary_label.setFont(QFont("Segoe UI", 10))
        self._summary_label.setStyleSheet("color: #4d5d6d;")
        runner_layout.addWidget(self._summary_label)

        self._table = QTableWidget(runner_tab)
        self._table.setSortingEnabled(True)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        runner_layout.addWidget(self._table, 1)

        self._tabs.addTab(runner_tab, "Runner")

    def _build_club_tab(self) -> None:
        club_tab = QWidget()
        club_layout = QVBoxLayout(club_tab)
        club_layout.setContentsMargins(0, 0, 0, 0)
        club_layout.setSpacing(10)

        controls = QWidget(club_tab)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        label = QLabel("Club:", controls)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(label)

        self._club_combo = QComboBox(controls)
        self._club_combo.setEditable(True)
        self._club_combo.setInsertPolicy(QComboBox.NoInsert)
        self._club_combo.setMinimumWidth(420)
        self._club_combo.activated.connect(self._load_club_history)
        controls_layout.addWidget(self._club_combo)

        load_btn = QPushButton("Load", controls)
        load_btn.clicked.connect(self._load_club_history)
        controls_layout.addWidget(load_btn)

        refresh_btn = QPushButton("Refresh", controls)
        refresh_btn.clicked.connect(self._refresh_club_list)
        controls_layout.addWidget(refresh_btn)

        club_layout.addWidget(controls)

        self._club_summary_label = QLabel("", club_tab)
        self._club_summary_label.setFont(QFont("Segoe UI", 10))
        self._club_summary_label.setStyleSheet("color: #4d5d6d;")
        club_layout.addWidget(self._club_summary_label)

        self._club_table = QTableWidget(club_tab)
        self._club_table.setSortingEnabled(True)
        self._club_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._club_table.setSelectionMode(QTableWidget.SingleSelection)
        self._club_table.setAlternatingRowColors(True)
        self._club_table.setEditTriggers(QTableWidget.NoEditTriggers)
        club_layout.addWidget(self._club_table, 1)

        self._tabs.addTab(club_tab, "Club - Runners")
        self._build_club_races_tab()

    def _build_club_races_tab(self) -> None:
        club_races_tab = QWidget()
        club_races_layout = QVBoxLayout(club_races_tab)
        club_races_layout.setContentsMargins(0, 0, 0, 0)
        club_races_layout.setSpacing(10)

        controls = QWidget(club_races_tab)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        label = QLabel("Club:", controls)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(label)

        self._club_races_combo = QComboBox(controls)
        self._club_races_combo.setEditable(True)
        self._club_races_combo.setInsertPolicy(QComboBox.NoInsert)
        self._club_races_combo.setMinimumWidth(420)
        self._club_races_combo.activated.connect(self._load_club_races)
        controls_layout.addWidget(self._club_races_combo)

        load_btn = QPushButton("Load", controls)
        load_btn.clicked.connect(self._load_club_races)
        controls_layout.addWidget(load_btn)

        refresh_btn = QPushButton("Refresh", controls)
        refresh_btn.clicked.connect(self._refresh_club_list)
        controls_layout.addWidget(refresh_btn)

        club_races_layout.addWidget(controls)

        self._club_races_summary_label = QLabel("", club_races_tab)
        self._club_races_summary_label.setFont(QFont("Segoe UI", 10))
        self._club_races_summary_label.setStyleSheet("color: #4d5d6d;")
        club_races_layout.addWidget(self._club_races_summary_label)

        self._club_races_table = QTableWidget(club_races_tab)
        self._club_races_table.setSortingEnabled(True)
        self._club_races_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._club_races_table.setSelectionMode(QTableWidget.SingleSelection)
        self._club_races_table.setAlternatingRowColors(True)
        self._club_races_table.setEditTriggers(QTableWidget.NoEditTriggers)
        club_races_layout.addWidget(self._club_races_table, 1)

        self._tabs.addTab(club_races_tab, "Club - Races")

    def _results_workbook_path(self) -> Path | None:
        if self._output_dir is None:
            return None
        return find_latest_results_workbook(self._output_dir)

    def _refresh_club_list(self) -> None:
        self._cache = load_workbook_cache()
        clubs: list[str] = []
        if self._cache is not None:
            clubs = self._extract_club_names(self._cache)
        self._club_combo.clear()
        self._club_combo.addItems(clubs)
        if hasattr(self, '_club_races_combo'):
            self._club_races_combo.clear()
            self._club_races_combo.addItems(clubs)
        self._club_summary_label.setText("")
        self._club_table.clear()
        self._club_table.setRowCount(0)
        self._club_table.setColumnCount(0)
        if hasattr(self, '_club_races_table'):
            self._club_races_table.clear()
            self._club_races_table.setRowCount(0)
            self._club_races_table.setColumnCount(0)

        if not clubs:
            self._club_summary_label.setText("No club history workbook found or it could not be loaded.")
            if hasattr(self, '_club_races_summary_label'):
                self._club_races_summary_label.setText("No club history workbook found or it could not be loaded.")

    def _extract_club_names(self, cache) -> list[str]:
        names = set()
        for sheet in cache["race_sheets"]:
            df = cache["sheets"][sheet]
            for col in ("Club", "Team", "Affiliation"):
                if col in df.columns:
                    names.update(str(v).strip() for v in df[col].dropna() if str(v).strip())
        return sorted(names, key=lambda n: n.lower())

    def _load_club_history(self) -> None:
        club_name = self._club_combo.currentText().strip()
        if not club_name:
            self._club_summary_label.setText("Enter or select a club.")
            return

        if self._cache is None:
            self._refresh_club_list()
            if self._cache is None:
                return

        club = club_name
        rows = []
        for sheet in self._cache["race_sheets"]:
            df = self._cache["sheets"][sheet].fillna("")
            club_col = None
            name_col = None
            points_col = None
            time_col = None
            for c in df.columns:
                c_lower = str(c).strip().lower()
                if c_lower in ("club", "team", "affiliation") and club_col is None:
                    club_col = c
                if c_lower == "name" and name_col is None:
                    name_col = c
                if c_lower == "points" and points_col is None:
                    points_col = c
                if c_lower == "time" and time_col is None:
                    time_col = c
            if club_col is None or name_col is None:
                continue

            race_num = extract_race_number(sheet) or ""
            for _, row in df.iterrows():
                rc = str(row.get(club_col, "")).strip()
                if rc.lower() != club.lower():
                    continue
                rows.append({
                    "Race": race_num,
                    "Name": str(row.get(name_col, "")).strip(),
                    "Category": str(row.get("Category", "")).strip(),
                    "Time": str(row.get(time_col, "")).strip() if time_col is not None else "",
                    "Points": str(row.get(points_col, "")).strip() if points_col is not None else "",
                })

        if not rows:
            self._club_summary_label.setText(f"Club '{club}' not found in history workbook.")
            self._club_table.clear()
            self._club_table.setRowCount(0)
            self._club_table.setColumnCount(0)
            return

        self._club_table.clear()
        cols = ["Race", "Name", "Category", "Time", "Points"]
        self._club_table.setColumnCount(len(cols))
        self._club_table.setRowCount(len(rows))
        self._club_table.setHorizontalHeaderLabels(cols)
        for row_index, entry in enumerate(rows):
            for col_index, key in enumerate(cols):
                item = QTableWidgetItem(entry[key])
                if key == "Points":
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self._club_table.setItem(row_index, col_index, item)
        self._club_table.resizeColumnsToContents()
        total_points = 0
        for entry in rows:
            try:
                total_points += int(float(entry["Points"]))
            except Exception:
                pass
        self._club_summary_label.setText(f"{club}: {len(rows)} row(s), total points {total_points}.")

    def _load_club_races(self) -> None:
        club_name = self._club_races_combo.currentText().strip()
        if not club_name:
            self._club_races_summary_label.setText("Enter or select a club.")
            return

        if self._cache is None:
            self._refresh_club_list()
            if self._cache is None:
                return

        club = club_name
        rows = []
        race_points = {}
        for sheet in self._cache["race_sheets"]:
            df = self._cache["sheets"][sheet].fillna("")
            club_col = None
            name_col = None
            points_col = None
            time_col = None
            for c in df.columns:
                c_lower = str(c).strip().lower()
                if c_lower in ("club", "team", "affiliation") and club_col is None:
                    club_col = c
                if c_lower == "name" and name_col is None:
                    name_col = c
                if c_lower == "points" and points_col is None:
                    points_col = c
                if c_lower == "time" and time_col is None:
                    time_col = c
            if club_col is None or name_col is None:
                continue

            race_num = extract_race_number(sheet) or ""
            score_sum = 0
            recs = []
            for _, row in df.iterrows():
                rc = str(row.get(club_col, "")).strip()
                if rc.lower() != club.lower():
                    continue
                points = str(row.get(points_col, "")).strip() if points_col is not None else ""
                try:
                    score_sum += int(float(points))
                except Exception:
                    pass
                recs.append({
                    "Race": race_num,
                    "Name": str(row.get(name_col, "")).strip(),
                    "Time": str(row.get(time_col, "")).strip() if time_col is not None else "",
                    "Points": points,
                })
            if recs:
                rows.extend(recs)
                race_points[race_num] = race_points.get(race_num, 0) + score_sum

        if not rows:
            self._club_races_summary_label.setText(f"Club '{club}' not found in history workbook.")
            self._club_races_table.clear()
            self._club_races_table.setRowCount(0)
            self._club_races_table.setColumnCount(0)
            return

        self._club_races_table.clear()
        cols = ["Race", "Name", "Time", "Points"]
        self._club_races_table.setColumnCount(len(cols))
        self._club_races_table.setRowCount(len(rows))
        self._club_races_table.setHorizontalHeaderLabels(cols)
        for row_index, entry in enumerate(rows):
            for col_index, key in enumerate(cols):
                item = QTableWidgetItem(entry[key])
                if key == "Points":
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self._club_races_table.setItem(row_index, col_index, item)
        self._club_races_table.resizeColumnsToContents()
        total_points = sum(int(float(entry["Points"])) for entry in rows if str(entry["Points"]).strip().replace(" ", "") and self._is_number(entry["Points"]))
        self._club_races_summary_label.setText(
            f"{club}: {len(rows)} row(s), total points {total_points}."
        )

    def _is_number(self, value: str) -> bool:
        try:
            float(value)
            return True
        except Exception:
            return False

    def _refresh_runner_list(self) -> None:
        self._cache = load_workbook_cache()
        names = []
        if self._cache is not None:
            names = extract_runner_names(self._cache)
        self._runner_combo.clear()
        self._runner_combo.addItems(names)
        self._current_df = None
        self._summary_label.setText("")
        self._table.clear()
        self._table.setRowCount(0)
        self._table.setColumnCount(0)

        if not names:
            self._summary_label.setText("No runner history workbook found or it could not be loaded.")

    def _load_runner_history(self) -> None:
        runner_name = self._runner_combo.currentText().strip()
        if not runner_name:
            self._summary_label.setText("Enter or select a runner name.")
            return

        if self._cache is None:
            self._refresh_runner_list()
            if self._cache is None:
                return

        pc = proper_case(runner_name)
        df = extract_runner_history(self._cache, pc)
        if df.empty:
            self._summary_label.setText(f"Runner '{pc}' not found in history workbook.")
            self._current_df = None
            self._table.clear()
            self._table.setRowCount(0)
            self._table.setColumnCount(0)
            return

        self._current_df = df
        total = 0
        try:
            total = int(df["Points"].astype(str).replace("", "0").astype(int).sum())
        except Exception:
            total = 0

        self._summary_label.setText(f"{pc}: {len(df)} race row(s), total points {total}.")
        self._populate_table(df)

        conflicts = detect_conflicts(df)
        if conflicts:
            self._summary_label.setText(self._summary_label.text() + "  Conflicts detected: " + ", ".join(conflicts.keys()))

    def _populate_table(self, df) -> None:
        self._table.clear()
        self._table.setColumnCount(len(df.columns))
        self._table.setRowCount(len(df.index))
        self._table.setHorizontalHeaderLabels([str(col) for col in df.columns])

        numeric_cols = set()
        for col in df.columns:
            try:
                pd_col = df[col]
                numeric_vals = pd.to_numeric(pd_col, errors="coerce")
                if numeric_vals.notna().all():
                    numeric_cols.add(col)
            except Exception:
                continue

        for row_index, row in enumerate(df.itertuples(index=False, name=None)):
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value) if value is not None else "")
                if df.columns[col_index] in numeric_cols:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self._table.setItem(row_index, col_index, item)

        self._table.resizeColumnsToContents()
        self._table.resizeRowsToContents()

    def _copy_results(self) -> None:
        if self._current_df is None or self._current_df.empty:
            QMessageBox.information(self, "Copy Results", "No runner history loaded.", parent=self)
            return

        df = sanitise_df_for_export(self._current_df)
        text = df.to_string(index=False)
        QGuiApplication.clipboard().setText(text)
        QMessageBox.information(self, "Copy Results", "Runner history copied to clipboard.", parent=self)
