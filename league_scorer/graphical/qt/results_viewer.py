"""Qt window for viewing generated league results from the latest standings workbook."""

from __future__ import annotations

import math
import os
import re
from pathlib import Path

import pandas as pd
from pandas.api import types as pdtypes
from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QStyle,
    QTreeWidget,
    QTreeWidgetItem,
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
from league_scorer.normalisation import parse_time_to_seconds
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
        self._current_race_sheet: str | None = None
        self._current_option = _VIEW_OPTIONS[0][1]
        self._wiltshire_only = False
        self._gender_filter = "all"
        self.setWindowTitle("League Results")
        self.resize(1200, 780)
        self._build_ui()
        self._refresh_results()

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setObjectName("resultsViewer")
        central.setStyleSheet(f"background: {WRRL_LIGHT};")
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        top_bar = QWidget(central)
        top_bar.setStyleSheet("background: #ffffff; border-radius: 12px;")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(16, 16, 16, 16)
        top_layout.setSpacing(12)

        header = QLabel("League Results", top_bar)
        header.setFont(QFont("Segoe UI", 16, QFont.Bold))
        header.setStyleSheet("color: #2d7a4a;")
        top_layout.addWidget(header)
        top_layout.addStretch(1)

        root_layout.addWidget(top_bar)

        button_bar = QWidget(central)
        button_bar.setStyleSheet("background: #ffffff; border-radius: 12px;")
        button_layout = QHBoxLayout(button_bar)
        button_layout.setContentsMargins(16, 12, 16, 12)
        button_layout.setSpacing(12)

        refresh_btn = QPushButton("Refresh", button_bar)
        refresh_btn.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        refresh_btn.clicked.connect(self._refresh_results)
        refresh_btn.setCursor(Qt.PointingHandCursor)
        button_layout.addWidget(refresh_btn)

        open_btn = QPushButton("Open Workbook", button_bar)
        open_btn.clicked.connect(self._open_results_workbook)
        open_btn.setCursor(Qt.PointingHandCursor)
        button_layout.addWidget(open_btn)

        button_layout.addStretch(1)
        root_layout.addWidget(button_bar)

        content = QWidget(central)
        content_layout = QHBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(12)

        left_panel = QWidget(content)
        left_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)

        action_label = QLabel("Reports", left_panel)
        action_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        action_label.setStyleSheet("color: #2d7a4a;")
        left_layout.addWidget(action_label)

        self._view_tree = QTreeWidget(left_panel)
        self._view_tree.setHeaderHidden(True)
        self._view_tree.itemSelectionChanged.connect(self._on_tree_item_selected)
        self._view_tree.setIndentation(16)
        left_layout.addWidget(self._view_tree, 1)

        self._wiltshire_only_btn = QPushButton("Wiltshire only", left_panel)
        self._wiltshire_only_btn.setCheckable(True)
        self._wiltshire_only_btn.setEnabled(False)
        self._wiltshire_only_btn.clicked.connect(self._on_wiltshire_only_toggled)
        left_layout.addWidget(self._wiltshire_only_btn)

        self._gender_filter_btn = QPushButton("Gender: All", left_panel)
        self._gender_filter_btn.setCheckable(False)
        self._gender_filter_btn.setEnabled(False)
        self._gender_filter_btn.clicked.connect(self._on_gender_filter_clicked)
        left_layout.addWidget(self._gender_filter_btn)

        content_layout.addWidget(left_panel, 1)

        right_panel = QWidget(content)
        right_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)

        self._table = QTableWidget(right_panel)
        self._table.setSortingEnabled(True)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setShowGrid(False)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        right_layout.addWidget(self._table, 1)

        content_layout.addWidget(right_panel, 3)
        root_layout.addWidget(content, 1)

        self._message_bar = QWidget(central)
        self._message_bar.setStyleSheet("background: #1e1e1e; border-radius: 8px;")
        message_layout = QHBoxLayout(self._message_bar)
        message_layout.setContentsMargins(14, 10, 14, 10)

        self._message_label = QLabel("Ready", self._message_bar)
        self._message_label.setFont(QFont("Segoe UI", 10))
        self._message_label.setStyleSheet("color: #d4d4d4;")
        message_layout.addWidget(self._message_label)

        root_layout.addWidget(self._message_bar)

    def _results_workbook_path(self) -> Path | None:
        return find_latest_results_workbook(self._output_dir)

    def _refresh_results(self) -> None:
        results_path = self._results_workbook_path()
        self._current_results_path = results_path
        if results_path is None:
            self._show_message("No standings workbook found in outputs/publish/standings.")
            return

        self._message_label.setText(f"Workbook: {results_path.name}")
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

    def _on_tree_item_selected(self) -> None:
        selected_items = self._view_tree.selectedItems()
        if not selected_items:
            return
        item = selected_items[0]
        data = item.data(0, Qt.UserRole)
        if not isinstance(data, dict):
            return

        if data["type"] == "option":
            self._current_option = data["option"]
            self._current_race_sheet = None
        elif data["type"] == "race":
            self._current_option = "race"
            self._current_race_sheet = data["sheet"]
        else:
            return

        self._update_race_only_button()
        self._load_results()

    def _on_wiltshire_only_toggled(self, checked: bool) -> None:
        self._wiltshire_only = checked
        self._wiltshire_only_btn.setText("Wiltshire only" if not checked else "Wiltshire only ✓")
        self._load_results()

    def _on_gender_filter_clicked(self) -> None:
        self._gender_filter = {
            "all": "male",
            "male": "female",
            "female": "all",
        }[self._gender_filter]
        labels = {"all": "Gender: All", "male": "Gender: Male", "female": "Gender: Female"}
        self._gender_filter_btn.setText(labels[self._gender_filter])
        self._load_results()

    def _update_race_only_button(self) -> None:
        enabled = self._current_option == "race"
        self._wiltshire_only_btn.setEnabled(enabled)
        self._gender_filter_btn.setEnabled(enabled)
        if not enabled:
            self._wiltshire_only_btn.setChecked(False)
            self._wiltshire_only = False
            self._wiltshire_only_btn.setText("Wiltshire only")
            self._gender_filter = "all"
            self._gender_filter_btn.setText("Gender: All")

    def _load_race_options(self) -> None:
        self._view_tree.clear()
        if self._current_results_path is None:
            return

        try:
            xl = pd.ExcelFile(self._current_results_path)
        except Exception:
            return

        teams_root = QTreeWidgetItem(self._view_tree, ["Teams"])
        teams_root.setFlags(teams_root.flags() & ~Qt.ItemIsSelectable)
        for label, option in [("Division 1", "div1"), ("Division 2", "div2")]:
            child = QTreeWidgetItem(teams_root, [label])
            child.setData(0, Qt.UserRole, {"type": "option", "option": option})

        runners_root = QTreeWidgetItem(self._view_tree, ["Runners"])
        runners_root.setFlags(runners_root.flags() & ~Qt.ItemIsSelectable)
        for label, option in [("Male Top 20", "male"), ("Female Top 20", "female")]:
            child = QTreeWidgetItem(runners_root, [label])
            child.setData(0, Qt.UserRole, {"type": "option", "option": option})

        races_root = QTreeWidgetItem(self._view_tree, ["Races"])
        races_root.setFlags(races_root.flags() & ~Qt.ItemIsSelectable)
        races = sorted_race_sheet_names(xl)
        for race in races:
            child = QTreeWidgetItem(races_root, [display_race_sheet_name(race)])
            child.setData(0, Qt.UserRole, {"type": "race", "sheet": race})

        self._view_tree.expandItem(teams_root)
        self._view_tree.expandItem(runners_root)
        self._view_tree.expandItem(races_root)

        if self._current_option == "race" and self._current_race_sheet:
            self._select_tree_item("race", self._current_race_sheet)
        else:
            self._select_tree_item("option", self._current_option)

    def _select_tree_item(self, item_type: str, value: str) -> None:
        root_count = self._view_tree.topLevelItemCount()
        for i in range(root_count):
            root = self._view_tree.topLevelItem(i)
            for j in range(root.childCount()):
                child = root.child(j)
                data = child.data(0, Qt.UserRole)
                if isinstance(data, dict) and data.get("type") == item_type:
                    if item_type == "option" and data.get("option") == value:
                        self._view_tree.blockSignals(True)
                        self._view_tree.setCurrentItem(child)
                        self._view_tree.blockSignals(False)
                        return
                    if item_type == "race" and data.get("sheet") == value:
                        self._view_tree.blockSignals(True)
                        self._view_tree.setCurrentItem(child)
                        self._view_tree.blockSignals(False)
                        return

    def _load_results(self) -> None:
        self._update_race_only_button()
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
            status_info = ""
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
                if not self._current_race_sheet:
                    self._show_message("Select a race sheet to display.")
                    return
                df = xl.parse(self._current_race_sheet)
                if self._wiltshire_only:
                    eligible_col = None
                    for col in df.columns:
                        if str(col).strip().lower() == "wiltshire eligible":
                            eligible_col = col
                            break
                    if eligible_col is not None:
                        df = df[df[eligible_col].astype(str).str.strip().str.lower() != "no"]
                status_info = ""
                if self._gender_filter != "all":
                    gender_col = None
                    normalized_columns = [str(col).strip().lower() for col in df.columns]
                    if "gender" in normalized_columns:
                        gender_col = df.columns[normalized_columns.index("gender")]
                    elif "sex" in normalized_columns:
                        gender_col = df.columns[normalized_columns.index("sex")]
                    else:
                        candidates = [
                            col
                            for col, name in zip(df.columns, normalized_columns)
                            if ("gender" in name or "sex" in name)
                            and "pos" not in name
                            and "position" not in name
                        ]
                        if candidates:
                            gender_col = candidates[0]
                        else:
                            for col, name in zip(df.columns, normalized_columns):
                                if "gender" in name or "sex" in name:
                                    gender_col = col
                                    break

                    if gender_col is None:
                        available = ", ".join(str(col) for col in df.columns)
                        self._show_message(
                            "Gender filter unavailable: no Gender/Sex column found. "
                            f"Available columns: {available}"
                        )
                        return

                    expected = self._gender_filter.lower()

                    unique_values = (
                        df[gender_col]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .replace("", pd.NA)
                        .dropna()
                        .astype(str)
                        .str.upper()
                        .unique()
                    )
                    unique_values = sorted(unique_values, key=str)
                    status_info = f" | Gender col: {gender_col} values: {', '.join(unique_values[:10])}"

                    def gender_matches(value: object) -> bool:
                        if pd.isna(value):
                            return False
                        text = str(value).strip().lower()
                        if not text:
                            return False
                        if expected == "male":
                            return bool(re.match(r"^(m|male|1)$", text))
                        return bool(re.match(r"^(f|female|2)$", text))

                    df = df[df[gender_col].apply(gender_matches)]
                drop_cols = {"gender pos", "club", "category"}
                df = df[[col for col in df.columns if str(col).strip().lower() not in drop_cols]]
            else:
                self._show_message("Unknown results view selected.")
                return

            self._current_df = df
            self._populate_table(df)
            self._message_label.setText(
                f"Workbook: {self._current_results_path.name} | View: {self._current_option}{status_info}"
            )
        except Exception as exc:
            self._show_message(f"Failed to load results: {exc}")

    def _populate_table(self, df: pd.DataFrame) -> None:
        self._table.clear()
        self._table.setColumnCount(len(df.columns))
        self._table.setRowCount(len(df.index))
        self._display_headers = [display_race_column_header(str(col)) for col in df.columns]
        self._display_headers = [
            re.sub(r"\bscore\b", "", header, flags=re.IGNORECASE).strip()
            for header in self._display_headers
        ]
        self._display_headers = [
            header.replace("Aggregate", "(M+W)").replace("aggregate", "(M+W)")
            for header in self._display_headers
        ]
        self._display_headers = [
            header.replace("Team Points", "Pts").replace("team points", "Pts")
            for header in self._display_headers
        ]
        self._display_headers = [
            header.replace("R1 Men", "R1 M").replace("R1 Women", "R1 W")
            for header in self._display_headers
        ]
        self._table.setHorizontalHeaderLabels(self._display_headers)

        numeric_cols = set()
        center_cols = set()
        time_like_cols = {
            col
            for col, header in zip(df.columns, self._display_headers)
            if "time" in header.strip().lower()
        }
        for col, header in zip(df.columns, self._display_headers):
            if col in time_like_cols:
                continue
            if pdtypes.is_numeric_dtype(df[col]) or pd.to_numeric(df[col], errors="coerce").notna().any():
                numeric_cols.add(col)
            if header.strip().lower() in {"pos", "position"}:
                center_cols.add(col)

        for row_idx, row in enumerate(df.itertuples(index=False, name=None)):
            for col_idx, value in enumerate(row):
                col_name = df.columns[col_idx]

                if isinstance(value, str) and value.strip().lower() == "nan":
                    text = ""
                elif pd.isna(value):
                    text = ""
                elif col_name in time_like_cols:
                    seconds = parse_time_to_seconds(value)
                    if seconds is None:
                        text = str(value).strip() if value is not None else ""
                    else:
                        tenths = int((seconds - int(seconds)) * 10)
                        h = int(seconds // 3600)
                        m = int((seconds % 3600) // 60)
                        s = int(seconds % 60)
                        text = f"{h}:{m:02d}:{s:02d}.{tenths}"
                elif col_name in numeric_cols:
                    try:
                        numeric_value = float(value)
                    except (TypeError, ValueError):
                        text = ""
                    else:
                        if not math.isfinite(numeric_value):
                            text = ""
                        else:
                            text = str(int(round(numeric_value)))
                else:
                    text = str(value) if value is not None else ""

                item = QTableWidgetItem(text)
                if self._current_option == "race":
                    item.setTextAlignment(Qt.AlignCenter)
                elif col_name in center_cols:
                    item.setTextAlignment(Qt.AlignCenter)
                elif col_name in numeric_cols:
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
        self._message_label.setText(message)
