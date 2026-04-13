"""Qt window for viewing generated league results from the latest standings workbook."""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from league_scorer.graphical.results_workbook import (
    display_race_column_header,
    display_race_sheet_name,
    find_latest_results_workbook,
    sorted_race_sheet_names,
)
from league_scorer.session_config import config as session_config

WRRL_NAVY = "#3a4658"
WRRL_LIGHT = "#f5f5f5"

_VIEW_OPTIONS = [
    ("Division 1 Teams", "div1"),
    ("Division 2 Teams", "div2"),
    ("Top 20 Male Individuals", "male"),
    ("Top 20 Female Individuals", "female"),
    ("Race Results", "race"),
]


class ResultsViewerWindow(QMainWindow):
    def __init__(self, output_dir: Path | None = None) -> None:
        super().__init__()
        self._output_dir = output_dir
        self._current_df: pd.DataFrame | None = None
        self._current_results_path: Path | None = None
        self._current_option = _VIEW_OPTIONS[0][1]
        self.setWindowTitle("View League Results")
        self.resize(1200, 780)
        self._build_ui()
        self._refresh_results()

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setObjectName("resultsViewer")
        central.setStyleSheet(f"background: {WRRL_LIGHT};")
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header = QLabel("View League Results", central)
        header.setFont(QFont("Segoe UI", 18, QFont.Bold))
        header.setStyleSheet(f"color: {WRRL_NAVY};")
        layout.addWidget(header)

        controls = QWidget(central)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(10)

        refresh_btn = QPushButton("Refresh Workbook", controls)
        refresh_btn.clicked.connect(self._refresh_results)
        controls_layout.addWidget(refresh_btn)

        open_btn = QPushButton("Open Workbook", controls)
        open_btn.clicked.connect(self._open_results_workbook)
        controls_layout.addWidget(open_btn)

        controls_layout.addStretch(1)

        view_label = QLabel("View:", controls)
        controls_layout.addWidget(view_label)

        self._view_selector = QComboBox(controls)
        self._view_selector.addItems([name for name, _ in _VIEW_OPTIONS])
        self._view_selector.currentIndexChanged.connect(self._on_view_changed)
        controls_layout.addWidget(self._view_selector)

        race_label = QLabel("Race:", controls)
        controls_layout.addWidget(race_label)

        self._race_selector = QComboBox(controls)
        self._race_selector.setEnabled(False)
        self._race_selector.currentIndexChanged.connect(self._load_results)
        controls_layout.addWidget(self._race_selector)

        layout.addWidget(controls)

        self._status_label = QLabel("", central)
        self._status_label.setFont(QFont("Segoe UI", 9))
        self._status_label.setStyleSheet("color: #4d5d6d;")
        layout.addWidget(self._status_label)

        self._table = QTableWidget(central)
        self._table.setSortingEnabled(True)
        self._table.setAlternatingRowColors(True)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self._table, 1)

    def _results_workbook_path(self) -> Path | None:
        return find_latest_results_workbook(self._output_dir)

    def _refresh_results(self) -> None:
        results_path = self._results_workbook_path()
        self._current_results_path = results_path
        if results_path is None:
            self._show_message("No standings workbook found in outputs/publish/standings.")
            return

        self._status_label.setText(f"Workbook: {results_path.name}")
        self._load_race_options()
        self._load_results()

    def _open_results_workbook(self) -> None:
        if self._current_results_path is None or not self._current_results_path.exists():
            QMessageBox.warning(self, "Workbook Missing", "No standings workbook is available to open.", parent=self)
            return

        try:
            if os.name == "nt":
                os.startfile(str(self._current_results_path))
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._current_results_path)))
        except Exception as exc:
            QMessageBox.critical(self, "Open Failed", f"Could not open workbook: {exc}", parent=self)

    def _on_view_changed(self, index: int) -> None:
        self._current_option = _VIEW_OPTIONS[index][1]
        self._race_selector.setEnabled(self._current_option == "race")
        self._load_results()

    def _load_race_options(self) -> None:
        self._race_selector.clear()
        if self._current_results_path is None:
            return

        try:
            xl = pd.ExcelFile(self._current_results_path)
        except Exception:
            return

        races = sorted_race_sheet_names(xl)
        for race in races:
            self._race_selector.addItem(display_race_sheet_name(race), race)
        if races:
            self._race_selector.setCurrentIndex(0)

    def _load_results(self) -> None:
        if self._current_results_path is None or not self._current_results_path.exists():
            self._show_message("No standings workbook found.")
            return

        try:
            xl = pd.ExcelFile(self._current_results_path)
        except Exception as exc:
            self._show_message(f"Failed to open workbook: {exc}")
            return

        try:
            option = self._current_option
            if option == "Summary":
                df = xl.parse("Summary")
            elif option == "div1":
                df = xl.parse("Div 1")
            elif option == "div2":
                df = xl.parse("Div 2")
            elif option == "male":
                df = xl.parse("Male").head(20)
            elif option == "female":
                df = xl.parse("Female").head(20)
            elif option == "race":
                race_name = self._race_selector.currentData() or self._race_selector.currentText()
                if not race_name:
                    self._show_message("Select a race sheet to display.")
                    return
                df = xl.parse(race_name)
            else:
                self._show_message("Unknown results view selected.")
                return

            self._current_df = df
            self._populate_table(df)
            self._status_label.setText(f"Workbook: {self._current_results_path.name} | View: {self._view_selector.currentText()}")
        except Exception as exc:
            self._show_message(f"Failed to load results: {exc}")

    def _populate_table(self, df: pd.DataFrame) -> None:
        self._table.clear()
        self._table.setColumnCount(len(df.columns))
        self._table.setRowCount(len(df.index))
        self._display_headers = [display_race_column_header(str(col)) for col in df.columns]
        self._table.setHorizontalHeaderLabels(self._display_headers)

        numeric_cols = set()
        for col in df.columns:
            numeric = pd.to_numeric(df[col], errors="coerce")
            if numeric.notna().all():
                numeric_cols.add(col)

        for row_idx, row in enumerate(df.itertuples(index=False, name=None)):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value) if value is not None else "")
                if df.columns[col_idx] in numeric_cols:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self._table.setItem(row_idx, col_idx, item)

        self._table.resizeColumnsToContents()
        self._table.resizeRowsToContents()

    def _show_message(self, message: str) -> None:
        self._table.clear()
        self._table.setColumnCount(1)
        self._table.setRowCount(1)
        self._table.setHorizontalHeaderLabels(["Message"])
        self._table.setItem(0, 0, QTableWidgetItem(message))
        self._table.resizeColumnsToContents()
        self._status_label.setText(message)
