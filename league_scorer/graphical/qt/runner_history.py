"""Qt window for viewing and copying runner history from the latest results workbook."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QGuiApplication, QIcon, QPixmap, QPainter
from league_scorer.events_loader import load_events
from league_scorer.normalisation import parse_time_to_seconds
from league_scorer.session_config import config as session_config
from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
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
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"


class RunnerHistoryWindow(QMainWindow):
    def __init__(self, output_dir: Path | None = None) -> None:
        super().__init__()
        self._output_dir = output_dir
        self._cache = None
        self._current_df = None
        self._events_schedule = None
        self._event_distance_by_race: dict[int, str] = {}
        self._event_name_distance: dict[str, str] = {}
        self._event_title_by_race: dict[int, str] = {}
        self.setWindowTitle("Runner / Club Enquiry")
        self.resize(1160, 760)
        self._build_ui()
        self._refresh_runner_list()
        self._refresh_club_list()

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setStyleSheet(f"background: {WRRL_LIGHT};")
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        header_panel = QWidget(central)
        header_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        header_layout = QHBoxLayout(header_panel)
        header_layout.setContentsMargins(16, 16, 16, 16)
        header_layout.setSpacing(12)

        header = QLabel("Runner / Club Enquiry", header_panel)
        header.setFont(QFont("Segoe UI", 17, QFont.Bold))
        header.setStyleSheet(f"color: {WRRL_GREEN};")
        header_layout.addWidget(header)
        header_layout.addStretch(1)

        root_layout.addWidget(header_panel)

        button_panel = QWidget(central)
        button_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        button_layout = QHBoxLayout(button_panel)
        button_layout.setContentsMargins(16, 12, 16, 12)
        button_layout.setSpacing(12)

        button_style = (
            "QPushButton { background: #ffffff; color: #3a4658; border: 1px solid #ccd7e3; border-radius: 8px; padding: 8px 14px; }"
            "QPushButton:hover { background: #eef2f7; }"
        )

        self._refresh_btn = QPushButton("Refresh", button_panel)
        self._refresh_btn.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self._refresh_btn.setCursor(Qt.PointingHandCursor)
        self._refresh_btn.setStyleSheet(button_style)
        self._refresh_btn.clicked.connect(self._on_top_refresh)
        button_layout.addWidget(self._refresh_btn)

        self._load_btn = QPushButton("Load", button_panel)
        self._load_btn.setCursor(Qt.PointingHandCursor)
        self._load_btn.setStyleSheet(button_style)
        self._load_btn.clicked.connect(self._on_top_load)
        button_layout.addWidget(self._load_btn)

        self._copy_btn = QPushButton("Copy", button_panel)
        self._copy_btn.setCursor(Qt.PointingHandCursor)
        self._copy_btn.setStyleSheet(button_style)
        self._copy_btn.clicked.connect(self._on_top_copy)
        button_layout.addWidget(self._copy_btn)

        self._close_btn = QPushButton("🏠 Close", button_panel)
        self._close_btn.setCursor(Qt.PointingHandCursor)
        self._close_btn.setStyleSheet(button_style)
        self._close_btn.clicked.connect(self.close)
        button_layout.addWidget(self._close_btn)

        button_layout.addStretch(1)
        root_layout.addWidget(button_panel)

        self._nav_group = QButtonGroup(self)
        self._nav_group.setExclusive(True)

        content_panel = QWidget(central)
        content_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        content_layout = QHBoxLayout(content_panel)
        content_layout.setContentsMargins(16, 16, 16, 16)
        content_layout.setSpacing(12)

        nav_panel = QWidget(content_panel)
        nav_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        nav_layout = QVBoxLayout(nav_panel)
        nav_layout.setContentsMargins(16, 16, 16, 16)
        nav_layout.setSpacing(12)

        self._content_stack = QStackedWidget(content_panel)
        self._content_stack.currentChanged.connect(self._update_top_buttons)

        self._build_runner_tab(nav_layout)
        self._build_club_tab(nav_layout)
        self._build_club_races_tab(nav_layout)

        content_layout.addWidget(nav_panel, 0)
        content_layout.addWidget(self._content_stack, 1)
        root_layout.addWidget(content_panel, 1)

        self._update_top_buttons()
        root_layout.addWidget(self._build_status_bar())

    def _build_runner_tab(self, nav_layout: QVBoxLayout) -> None:
        runner_tab = QWidget()
        runner_layout = QHBoxLayout(runner_tab)
        runner_layout.setContentsMargins(0, 0, 0, 0)
        runner_layout.setSpacing(12)

        left_panel = QWidget(runner_tab)
        left_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)

        controls = QWidget(left_panel)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        label = QLabel("Runner:", controls)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(label)

        self._runner_combo = QComboBox(controls)
        self._runner_combo.setEditable(True)
        self._runner_combo.setInsertPolicy(QComboBox.NoInsert)
        self._runner_combo.setMinimumWidth(280)
        self._runner_combo.activated.connect(self._load_runner_history)
        controls_layout.addWidget(self._runner_combo)

        left_layout.addWidget(controls)

        info_card = QFrame(left_panel)
        info_card.setStyleSheet(
            "background: #f8fbf8; border: 1px solid #dbeadf; border-radius: 12px;"
        )
        info_layout = QVBoxLayout(info_card)
        info_layout.setContentsMargins(14, 14, 14, 14)
        info_layout.setSpacing(12)

        self._runner_info_title = QLabel("Athlete summary", info_card)
        self._runner_info_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self._runner_info_title.setStyleSheet(f"color: {WRRL_GREEN};")
        info_layout.addWidget(self._runner_info_title)

        self._runner_info_label = QLabel("Club, gender, category and team details will appear here.", info_card)
        self._runner_info_label.setFont(QFont("Segoe UI", 10))
        self._runner_info_label.setStyleSheet("color: #4d5d6d; line-height: 1.9; margin-bottom: 6px;")
        self._runner_info_label.setWordWrap(True)
        self._runner_info_label.setTextFormat(Qt.RichText)
        info_layout.addWidget(self._runner_info_label)

        self._runner_best_label = QLabel("Best times will appear here.", info_card)
        self._runner_best_label.setFont(QFont("Segoe UI", 10))
        self._runner_best_label.setStyleSheet("color: #37523d; line-height: 1.9;")
        self._runner_best_label.setWordWrap(True)
        self._runner_best_label.setTextFormat(Qt.RichText)
        info_layout.addWidget(self._runner_best_label)

        left_layout.addWidget(info_card)
        left_layout.addStretch(1)
        runner_layout.addWidget(left_panel, 1)

        right_panel = QWidget(runner_tab)
        right_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)

        self._summary_label = QLabel("", right_panel)
        self._summary_label.setFont(QFont("Segoe UI", 10))
        self._summary_label.setStyleSheet("color: #4d5d6d;")
        right_layout.addWidget(self._summary_label)

        self._table = QTableWidget(right_panel)
        self._table.setSortingEnabled(True)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setShowGrid(False)
        self._table.verticalHeader().setVisible(False)
        right_layout.addWidget(self._table, 1)

        runner_layout.addWidget(right_panel, 3)
        runner_tab.setObjectName("runner_tab")
        self._add_nav_page(runner_tab, nav_layout, "👤", "Runner")

    def _build_club_tab(self, nav_layout: QVBoxLayout) -> None:
        club_tab = QWidget()
        club_layout = QHBoxLayout(club_tab)
        club_layout.setContentsMargins(0, 0, 0, 0)
        club_layout.setSpacing(12)

        left_panel = QWidget(club_tab)
        left_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)

        controls = QWidget(left_panel)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        label = QLabel("Club:", controls)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(label)

        self._club_combo = QComboBox(controls)
        self._club_combo.setEditable(True)
        self._club_combo.setInsertPolicy(QComboBox.NoInsert)
        self._club_combo.setMinimumWidth(280)
        self._club_combo.activated.connect(self._load_club_history)
        controls_layout.addWidget(self._club_combo)

        left_layout.addWidget(controls)

        club_card = QFrame(left_panel)
        club_card.setStyleSheet(
            "background: #f8fbf8; border: 1px solid #dbeadf; border-radius: 12px;"
        )
        club_card_layout = QVBoxLayout(club_card)
        club_card_layout.setContentsMargins(14, 14, 14, 14)
        club_card_layout.setSpacing(8)

        self._club_info_title = QLabel("Club summary", club_card)
        self._club_info_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self._club_info_title.setStyleSheet(f"color: {WRRL_GREEN};")
        club_card_layout.addWidget(self._club_info_title)

        self._club_info_label = QLabel("Club info will appear here.", club_card)
        self._club_info_label.setFont(QFont("Segoe UI", 10))
        self._club_info_label.setStyleSheet("color: #4d5d6d; line-height: 1.9;")
        self._club_info_label.setWordWrap(True)
        self._club_info_label.setTextFormat(Qt.RichText)
        club_card_layout.addWidget(self._club_info_label)

        self._club_best_label = QLabel("Best runners by category will appear here.", club_card)
        self._club_best_label.setFont(QFont("Segoe UI", 10))
        self._club_best_label.setStyleSheet("color: #37523d; line-height: 1.9;")
        self._club_best_label.setWordWrap(True)
        self._club_best_label.setTextFormat(Qt.RichText)
        club_card_layout.addWidget(self._club_best_label)
        left_layout.addWidget(club_card)

        left_layout.addStretch(1)
        club_layout.addWidget(left_panel, 1)

        right_panel = QWidget(club_tab)
        right_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)

        self._club_summary_label = QLabel("", right_panel)
        self._club_summary_label.setFont(QFont("Segoe UI", 10))
        self._club_summary_label.setStyleSheet("color: #4d5d6d;")
        right_layout.addWidget(self._club_summary_label)

        self._club_table = QTableWidget(right_panel)
        self._club_table.setSortingEnabled(True)
        self._club_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._club_table.setSelectionMode(QTableWidget.SingleSelection)
        self._club_table.setAlternatingRowColors(True)
        self._club_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._club_table.setShowGrid(False)
        self._club_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self._club_table, 1)

        club_layout.addWidget(right_panel, 3)
        club_tab.setObjectName("club_tab")
        self._add_nav_page(club_tab, nav_layout, "🏅", "Club - Runners")

    def _build_club_races_tab(self, nav_layout: QVBoxLayout) -> None:
        club_races_tab = QWidget()
        club_races_layout = QHBoxLayout(club_races_tab)
        club_races_layout.setContentsMargins(0, 0, 0, 0)
        club_races_layout.setSpacing(12)

        left_panel = QWidget(club_races_tab)
        left_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)

        controls = QWidget(left_panel)
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        label = QLabel("Club:", controls)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(label)

        self._club_races_combo = QComboBox(controls)
        self._club_races_combo.setEditable(True)
        self._club_races_combo.setInsertPolicy(QComboBox.NoInsert)
        self._club_races_combo.setMinimumWidth(280)
        self._club_races_combo.activated.connect(self._load_club_races)
        controls_layout.addWidget(self._club_races_combo)

        left_layout.addWidget(controls)

        club_races_card = QFrame(left_panel)
        club_races_card.setStyleSheet(
            "background: #f8fbf8; border: 1px solid #dbeadf; border-radius: 12px;"
        )
        club_races_card_layout = QVBoxLayout(club_races_card)
        club_races_card_layout.setContentsMargins(14, 14, 14, 14)
        club_races_card_layout.setSpacing(8)

        self._club_races_info_title = QLabel("Club races summary", club_races_card)
        self._club_races_info_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self._club_races_info_title.setStyleSheet(f"color: {WRRL_GREEN};")
        club_races_card_layout.addWidget(self._club_races_info_title)

        self._club_races_info_label = QLabel("Club summary info will appear here.", club_races_card)
        self._club_races_info_label.setFont(QFont("Segoe UI", 10))
        self._club_races_info_label.setStyleSheet("color: #4d5d6d; line-height: 1.9;")
        self._club_races_info_label.setWordWrap(True)
        self._club_races_info_label.setTextFormat(Qt.RichText)
        club_races_card_layout.addWidget(self._club_races_info_label)

        races_title = QLabel("Races", club_races_card)
        races_title.setFont(QFont("Segoe UI", 10, QFont.Bold))
        races_title.setStyleSheet("color: #37523d;")
        club_races_card_layout.addWidget(races_title)

        self._club_races_list = QListWidget(club_races_card)
        self._club_races_list.setSelectionMode(QListWidget.SingleSelection)
        self._club_races_list.itemClicked.connect(self._on_club_race_selected)
        self._club_races_list.setStyleSheet("background: #ffffff; border: 1px solid #dbeadf; border-radius: 8px;")
        club_races_card_layout.addWidget(self._club_races_list)

        left_layout.addWidget(club_races_card)

        left_layout.addStretch(1)
        club_races_layout.addWidget(left_panel, 1)

        right_panel = QWidget(club_races_tab)
        right_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)

        self._club_races_summary_label = QLabel("", right_panel)
        self._club_races_summary_label.setFont(QFont("Segoe UI", 10))
        self._club_races_summary_label.setStyleSheet("color: #4d5d6d;")
        right_layout.addWidget(self._club_races_summary_label)

        self._club_races_table = QTableWidget(right_panel)
        self._club_races_table.setSortingEnabled(True)
        self._club_races_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._club_races_table.setSelectionMode(QTableWidget.SingleSelection)
        self._club_races_table.setAlternatingRowColors(True)
        self._club_races_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._club_races_table.setShowGrid(False)
        self._club_races_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self._club_races_table, 1)

        club_races_layout.addWidget(right_panel, 3)
        club_races_tab.setObjectName("club_races_tab")
        self._add_nav_page(club_races_tab, nav_layout, "📋", "Club - Races")

    def _build_status_bar(self) -> QWidget:
        bar = QWidget(self)
        bar.setStyleSheet("background: #1e1e1e; border-radius: 8px;")
        bar.setFixedHeight(34)

        layout = QHBoxLayout(bar)
        layout.setContentsMargins(16, 8, 16, 8)

        self._message_label = QLabel("Type a name or use the drop-down list to search.", bar)
        self._message_label.setFont(QFont("Segoe UI", 9))
        self._message_label.setStyleSheet("color: #f5f5f5;")
        layout.addWidget(self._message_label)
        layout.addStretch(1)

        return bar

    def _on_top_load(self) -> None:
        current = self._content_stack.currentWidget()
        if current is None:
            return
        name = current.objectName()
        if name == "runner_tab":
            self._load_runner_history()
        elif name == "club_tab":
            self._load_club_history()
        elif name == "club_races_tab":
            self._load_club_races()

    def _on_top_refresh(self) -> None:
        current = self._content_stack.currentWidget()
        if current is None:
            return
        name = current.objectName()
        if name == "runner_tab":
            self._refresh_runner_list()
        else:
            self._refresh_club_list()

    def _on_top_copy(self) -> None:
        current = self._content_stack.currentWidget()
        if current is None:
            return
        name = current.objectName()
        if name == "runner_tab":
            self._copy_results()
        elif name == "club_tab":
            self._copy_table_widget(self._club_table, "Club runners")
        elif name == "club_races_tab":
            self._copy_table_widget(self._club_races_table, "Club race details")

    def _update_top_buttons(self, _index: int = 0) -> None:
        current = self._content_stack.currentWidget()
        if current is None:
            return
        name = current.objectName()
        self._load_btn.setEnabled(True)
        self._refresh_btn.setEnabled(True)
        self._copy_btn.setEnabled(name in {"runner_tab", "club_tab", "club_races_tab"})

    def _update_runner_info(self, df) -> None:
        club_vals = sorted({str(v).strip() for v in df.get("Club", []) if str(v).strip()}, key=str.lower)
        gender_vals = sorted({str(v).strip() for v in df.get("Gender", []) if str(v).strip()}, key=str.lower)
        category_vals = sorted({str(v).strip() for v in df.get("Category", []) if str(v).strip()}, key=str.lower)
        team_vals = sorted({str(v).strip() for v in df.get("Team", []) if str(v).strip()}, key=str.lower)

        club_text = ", ".join(club_vals) if club_vals else "Unknown"
        gender_text = ", ".join(gender_vals) if gender_vals else "Unknown"
        category_text = ", ".join(category_vals) if category_vals else "Unknown"
        team_text = ", ".join(team_vals) if team_vals else "Unknown"

        self._load_events_schedule()

        points = []
        for value in df.get("Points", []):
            try:
                points.append(int(float(value)))
            except Exception:
                continue
        total_points = sum(points)
        best_six_points = sum(sorted(points, reverse=True)[:6])

        self._runner_info_label.setText(
            f"<div style='margin-bottom: 6px;'><b>Club:</b> {club_text}</div>"
            f"<div style='margin-bottom: 6px;'><b>Gender:</b> {gender_text}</div>"
            f"<div style='margin-bottom: 6px;'><b>Category:</b> {category_text}</div>"
            f"<div style='margin-bottom: 6px;'><b>Race rows:</b> {len(df)}</div>"
            f"<div style='margin-bottom: 6px;'><b>Total points:</b> {total_points}</div>"
            f"<div><b>Best six points:</b> {best_six_points}</div>"
        )

        best_times = {}
        for distance, source, race, time in zip(
            df.get("Distance", []),
            df.get("Source", []),
            df.get("Race", []),
            df.get("Time", []),
        ):
            dist_value = str(distance).strip()
            if not dist_value:
                dist_value = self._distance_from_source(source) or self._distance_from_source(race) or ""
            category = self._distance_category(dist_value)
            if category is None:
                race_num = extract_race_number(str(race))
                if race_num is not None:
                    event_distance = self._event_distance_by_race.get(race_num, "")
                    category = self._distance_category(event_distance)
                    dist_value = event_distance
            if category is None:
                continue
            seconds = parse_time_to_seconds(time)
            if seconds is None:
                continue
            if category not in best_times or seconds < best_times[category]:
                best_times[category] = seconds

        best_text = []
        for label in ["Best 5k", "Best 5 Mile", "Best 10k", "Best Half"]:
            value = best_times.get(label)
            display = self._format_time(value) if value is not None else "N/A"
            best_text.append(
                f"<div style='margin-bottom: 6px;'><b>{label}:</b> {display}</div>"
            )
        self._runner_best_label.setText("".join(best_text))

    def _update_club_info(self, club: str, rows: list[dict]) -> None:
        runners = sorted({entry.get("Name", "") for entry in rows if entry.get("Name")})
        races = sorted({entry.get("Race", "") for entry in rows if entry.get("Race")})
        total_points = sum(int(float(entry["Points"])) for entry in rows if str(entry["Points"]).strip().replace(" ", "") and self._is_number(entry["Points"]))
        male_score = 0
        female_score = 0
        team_a_score = 0
        team_b_score = 0
        best_by_category: dict[str, tuple[str, int]] = {}

        for entry in rows:
            points = 0
            try:
                points = int(float(entry.get("Points", 0)))
            except Exception:
                continue
            gender = str(entry.get("Gender", "")).strip().lower()
            category = str(entry.get("Category", "")).strip().lower()
            if category in {"jun", "sen", "v40", "v50", "v60", "v70"}:
                if gender.startswith("m") or gender.startswith("f"):
                    key = (category, "m" if gender.startswith("m") else "f")
                    existing = best_by_category.get(key)
                    if existing is None or points > existing[1]:
                        best_by_category[key] = (str(entry.get("Name", "")), points)

        division_stats = self._division_scores_for_club(club)
        best_six_points = sum(sorted((int(entry['Points']) for entry in rows if str(entry.get('Points', '')).strip().replace(' ', '') and self._is_number(entry['Points'])), reverse=True)[:6])
        team_a_place = ""
        if division_stats.get("team_a_division") is not None and division_stats.get("team_a_place") is not None:
            team_a_place = f", Div {division_stats['team_a_division']} - {self._ordinal(division_stats['team_a_place'])} Place"
        team_b_place = ""
        if division_stats.get("team_b_division") is not None and division_stats.get("team_b_place") is not None:
            team_b_place = f", Div {division_stats['team_b_division']} - {self._ordinal(division_stats['team_b_place'])} Place"

        summary_lines = [
            f"<div style='margin-bottom: 6px;'><b>Club:</b> {club}</div>",
            f"<div style='margin-bottom: 6px;'><b>Runners:</b> {len(runners)}</div>",
            f"<div style='margin-bottom: 6px;'><b>Races:</b> {division_stats['race_count']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Male score:</b> {division_stats['male_score']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Female score:</b> {division_stats['female_score']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Combined score:</b> {division_stats['combined_score']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Team A points:</b> {division_stats['team_a_best6']} ({division_stats['team_a_points_total']}){team_a_place}</div>",
            f"<div><b>Team B points:</b> {division_stats['team_b_best6']} ({division_stats['team_b_points_total']}){team_b_place}</div>",
        ]
        self._club_info_label.setText("".join(summary_lines))

        valid_categories = {"jun", "sen", "v40", "v50", "v60", "v70"}
        best_text = []
        for (category, gender), (name, points) in sorted(best_by_category.items(), key=lambda x: (x[0][0].lower(), x[0][1])):
            gender_label = "Male" if gender == "m" else "Female"
            best_text.append(
                f"<div style='margin-bottom: 4px;'><b>{category.upper()} {gender_label}:</b> {name} ({points})</div>"
            )
        self._club_best_label.setText("".join(best_text) if best_text else "No category best runners found.")

    def _division_scores_for_club(self, club: str) -> dict[str, int]:
        if self._cache is None:
            return {
                "race_count": 0,
                "male_score": 0,
                "female_score": 0,
                "combined_score": 0,
                "team_a_points_total": 0,
                "team_b_points_total": 0,
                "team_a_best6": 0,
                "team_b_best6": 0,
                "team_a_division": None,
                "team_a_place": None,
                "team_b_division": None,
                "team_b_place": None,
            }

        male_score = 0
        female_score = 0
        team_a_points = []
        team_b_points = []
        team_a_place = None
        team_b_place = None
        team_a_division = None
        team_b_division = None

        division_sheets = self._cache.get("division_sheets", [])
        race_ids = set()
        for sheet in division_sheets:
            df = self._cache["sheets"].get(sheet)
            if df is None:
                continue
            club_col = None
            for c in df.columns:
                if str(c).strip().lower() == "club":
                    club_col = c
                    break
            if club_col is None:
                continue

            score_cols = [c for c in df.columns if "men score" in str(c).lower() or "women score" in str(c).lower()]
            team_points_cols = [c for c in df.columns if "team points" in str(c).lower()]
            if not score_cols and not team_points_cols:
                continue

            position_col = None
            for c in df.columns:
                if str(c).strip().lower() == "position":
                    position_col = c
                    break

            division_match = re.search(r"div\D*([12])", str(sheet).strip().lower())
            division_number = int(division_match.group(1)) if division_match else None

            for index, row in df.iterrows():
                club_value = str(row.get(club_col, "")).strip()
                if not club_value:
                    continue
                lower_club_value = club_value.strip().lower()
                club_key = re.escape(club.lower())
                is_team_a = bool(re.fullmatch(rf"{club_key}\s*(?:--\s*)?a", lower_club_value)) or bool(re.fullmatch(rf"{club_key}a", lower_club_value))
                is_team_b = bool(re.fullmatch(rf"{club_key}\s*(?:--\s*)?b", lower_club_value)) or bool(re.fullmatch(rf"{club_key}b", lower_club_value))
                if not is_team_a and not is_team_b:
                    continue

                team_id = "A" if is_team_a else "B"
                place = None
                if position_col is not None:
                    try:
                        place = int(float(row.get(position_col, 0)))
                    except Exception:
                        place = None
                if place is None:
                    place = index + 1

                if team_id == "A":
                    team_a_place = place
                    team_a_division = division_number
                elif team_id == "B":
                    team_b_place = place
                    team_b_division = division_number

                for col in score_cols:
                    val = row.get(col, "")
                    try:
                        ic = int(float(val))
                    except Exception:
                        continue
                    if ic:
                        col_lower = str(col).lower()
                        if "women score" in col_lower:
                            female_score += ic
                        elif "men score" in col_lower:
                            male_score += ic
                        race_part = col_lower.split()[1] if len(col_lower.split()) > 1 else col_lower
                        race_ids.add(race_part)
                for col in team_points_cols:
                    val = row.get(col, "")
                    try:
                        ic = int(float(val))
                    except Exception:
                        continue
                    if ic:
                        if team_id == "A":
                            team_a_points.append(ic)
                        elif team_id == "B":
                            team_b_points.append(ic)
                        race_parts = [part for part in str(col).strip().lower().split() if part.isdigit() or part.startswith("race")]
                        if race_parts:
                            race_ids.add(race_parts[-1])

        team_a_points_total = sum(team_a_points)
        team_b_points_total = sum(team_b_points)
        team_a_best6 = sum(sorted(team_a_points, reverse=True)[:6])
        team_b_best6 = sum(sorted(team_b_points, reverse=True)[:6])

        return {
            "race_count": len(race_ids),
            "male_score": male_score,
            "female_score": female_score,
            "combined_score": male_score + female_score,
            "team_a_points_total": team_a_points_total,
            "team_b_points_total": team_b_points_total,
            "team_a_best6": team_a_best6,
            "team_b_best6": team_b_best6,
            "team_a_division": team_a_division,
            "team_a_place": team_a_place,
            "team_b_division": team_b_division,
            "team_b_place": team_b_place,
        }

    def _update_club_races_info(self, club: str, rows: list[dict]) -> None:
        races = sorted({entry.get("Race", "") for entry in rows if entry.get("Race")})
        total_points = sum(int(float(entry["Points"])) for entry in rows if str(entry["Points"]).strip().replace(" ", "") and self._is_number(entry["Points"]))
        self._club_races_info_label.setText(
            f"Club: {club}\n"
            f"Race rows: {len(rows)}\n"
            f"Races: {len(races)}\n"
            f"Total points: {total_points}"
        )

    def _distance_category(self, value: object) -> str | None:
        if value is None:
            return None
        s = str(value).strip().lower()
        if not s:
            return None
        s = re.sub(r"[\s\.,]+", "", s)
        if re.search(r"^(5k|5km|5000m)$", s) or ("5" in s and "k" in s and "mile" not in s):
            return "Best 5k"
        if "5mile" in s or ("5" in s and "mile" in s):
            return "Best 5 Mile"
        if re.search(r"^(10k|10km|10000m)$", s) or ("10" in s and "k" in s and "mile" not in s):
            return "Best 10k"
        if "half" in s or "21k" in s or "21km" in s or "13.1" in s:
            return "Best Half"
        return None

    def _normalize_event_label(self, label: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", label.lower().strip())

    def _ordinal(self, value: int) -> str:
        if value % 100 in (11, 12, 13):
            suffix = "th"
        elif value % 10 == 1:
            suffix = "st"
        elif value % 10 == 2:
            suffix = "nd"
        elif value % 10 == 3:
            suffix = "rd"
        else:
            suffix = "th"
        return f"{value}{suffix}"

    def _distance_from_source(self, source: object) -> str | None:
        if self._events_schedule is None:
            return None
        source_text = str(source).strip()
        if not source_text:
            return None

        race_num = extract_race_number(source_text)
        if race_num is not None:
            for event in self._events_schedule.events:
                try:
                    event_race_num = extract_race_number(str(getattr(event, "race_ref", "")))
                    if event_race_num == race_num:
                        return event.distance
                except Exception:
                    continue

        source_key = self._normalize_event_label(source_text)
        if source_key:
            distance = self._event_name_distance.get(source_key)
            if distance:
                return distance
            for event_name, distance in self._event_name_distance.items():
                if source_key == event_name or source_key in event_name or event_name in source_key:
                    return distance

        return None

    def _load_events_from_excel(self, path: Path) -> bool:
        try:
            if path.suffix.lower() == ".xlsx":
                self._events_schedule = load_events(path)
                return True

            xl = pd.ExcelFile(path)
            if "Championship Events" not in xl.sheet_names:
                return False

            df = xl.parse("Championship Events")
            xl.close()
            distances = {}
            event_names = {}
            if "RaceRef" not in df.columns or "Distance" not in df.columns:
                return False

            for _, row in df.iterrows():
                race_ref = str(row.get("RaceRef", "")).strip()
                distance = str(row.get("Distance", "")).strip()
                event_name = str(row.get("EventName", "")).strip() if "EventName" in df.columns else ""
                if not race_ref or not distance:
                    continue
                rid = extract_race_number(race_ref)
                if rid is None:
                    continue
                distances[rid] = distance
                event_names[rid] = event_name

            if not distances:
                return False

            class _Schedule:
                def __init__(self, events):
                    self.events = events

                def __iter__(self):
                    return iter(self.events)

            class _Event:
                def __init__(self, race_ref, distance, event_name=""):
                    self.race_ref = race_ref
                    self.distance = distance
                    self.event_name = event_name

            self._events_schedule = _Schedule([
                _Event(race_ref=str(r), distance=d, event_name=event_names.get(r, ""))
                for r, d in distances.items()
            ])
            return True
        except Exception:
            self._events_schedule = None
            return False

    def _add_nav_page(self, page: QWidget, nav_layout: QVBoxLayout, emoji: str, tooltip: str) -> None:
        self._content_stack.addWidget(page)
        index = self._content_stack.indexOf(page)
        button = QToolButton(self)
        button.setCheckable(True)
        button.setIcon(self._emoji_icon(emoji))
        button.setIconSize(QSize(28, 28))
        button.setToolButtonStyle(Qt.ToolButtonIconOnly)
        button.setToolTip(tooltip)
        button.setStyleSheet(
            "QToolButton { background: transparent; border: none; padding: 8px; margin: 4px; }"
            "QToolButton:hover { background: #eef5f0; }"
            "QToolButton:checked { background: #dbeadf; border-radius: 12px; }"
        )
        button.clicked.connect(lambda checked, idx=index: self._content_stack.setCurrentIndex(idx))
        self._nav_group.addButton(button)
        nav_layout.addWidget(button)
        if index == 0:
            button.setChecked(True)

    def _emoji_icon(self, emoji: str) -> QIcon:
        pix = QPixmap(40, 40)
        pix.fill(Qt.transparent)
        painter = QPainter(pix)
        painter.setFont(QFont("Segoe UI Emoji", 22))
        painter.setPen(Qt.black)
        painter.drawText(pix.rect(), Qt.AlignCenter, emoji)
        painter.end()
        return QIcon(pix)

    def _format_time(self, value: object) -> str:
        if isinstance(value, (int, float)):
            seconds = float(value)
        else:
            seconds = parse_time_to_seconds(value)
        if seconds is None:
            return str(value).strip() if value is not None else ""
        integer_seconds = int(seconds)
        fraction = abs(seconds - integer_seconds)
        dec = int(fraction * 10 + 0.000001)
        if dec == 10:
            integer_seconds += 1
            dec = 0
        h = integer_seconds // 3600
        m = (integer_seconds % 3600) // 60
        s = integer_seconds % 60
        return f"{h}:{m:02d}:{s:02d}.{dec}"

    def _club_row_team_id(self, club: str, club_value: str, team_value: str) -> str | None:
        if not club_value:
            return None
        club_lower = club.strip().lower()
        value_lower = club_value.strip().lower()
        team_value = team_value.strip().upper() if team_value else ""

        # Exact club with explicit team column
        if value_lower == club_lower and team_value.endswith("A"):
            return "A"
        if value_lower == club_lower and team_value.endswith("B"):
            return "B"

        # Club with A/B suffix in club value
        if value_lower.startswith(club_lower):
            suffix = value_lower[len(club_lower):].strip()
            if suffix in {"a", "b", "-- a", "-- b", "- a", "- b"}:
                return "A" if suffix.endswith("a") else "B"
        if value_lower.endswith(" a") or value_lower.endswith(" -- a") or value_lower.endswith(" - a"):
            return "A"
        if value_lower.endswith(" b") or value_lower.endswith(" -- b") or value_lower.endswith(" - b"):
            return "B"

        return None

    def _results_workbook_path(self) -> Path | None:
        if self._output_dir is None:
            return None
        return find_latest_results_workbook(self._output_dir)

    def _find_events_path(self) -> Path | None:
        if session_config.events_path and session_config.events_path.exists():
            return session_config.events_path

        control_dir = session_config.control_dir
        if not control_dir or not control_dir.exists():
            return None

        for candidate in sorted(control_dir.iterdir(), key=lambda p: p.name.lower()):
            if candidate.name.lower() in {"wrrl_events.xlsx", "wrrl_events.xls"} and candidate.exists():
                return candidate

        candidates = [
            child
            for child in control_dir.iterdir()
            if child.suffix.lower() in {".xlsx", ".xls"} and "event" in child.name.lower()
        ]
        if candidates:
            candidates.sort(key=lambda p: p.name.lower())
            return candidates[0]

        return None

    def _load_events_schedule(self) -> None:
        if self._events_schedule is not None:
            return
        events_path = self._find_events_path()
        if events_path is None:
            self._events_schedule = None
            self._event_distance_by_race = {}
            self._event_name_distance = {}
            return
        self._load_events_from_excel(events_path)
        self._event_distance_by_race = {}
        self._event_name_distance = {}
        self._event_title_by_race = {}
        if self._events_schedule is None:
            return
        for event in self._events_schedule.events:
            distance = str(getattr(event, "distance", "")).strip()
            if not distance:
                continue
            race_ref = str(getattr(event, "race_ref", "")).strip()
            race_num = extract_race_number(race_ref)
            if race_num is not None:
                self._event_distance_by_race[race_num] = distance
            event_name = str(getattr(event, "event_name", "")).strip()
            if event_name:
                self._event_name_distance[self._normalize_event_label(event_name)] = distance
            if race_num is not None and event_name:
                self._event_title_by_race[race_num] = event_name

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
        self._club_info_label.setText("Club info will appear here.")
        self._club_table.clear()
        self._club_table.setRowCount(0)
        self._club_table.setColumnCount(0)
        if hasattr(self, '_club_races_list'):
            self._club_races_list.clear()
        if hasattr(self, '_club_races_table'):
            self._club_races_table.clear()
            self._club_races_table.setRowCount(0)
            self._club_races_table.setColumnCount(0)
        if hasattr(self, '_club_races_info_label'):
            self._club_races_info_label.setText("Club summary info will appear here.")

        if not clubs:
            self._club_summary_label.setText("No club history workbook found or it could not be loaded.")
            if hasattr(self, '_club_races_summary_label'):
                self._club_races_summary_label.setText("No club history workbook found or it could not be loaded.")

    def _extract_club_names(self, cache) -> list[str]:
        names = set()
        for sheet, df in cache["sheets"].items():
            for col in ("Club", "Team", "Affiliation"):
                if col in df.columns:
                    for v in df[col].dropna():
                        club_name = self._base_club_name(str(v).strip())
                        if club_name:
                            names.add(club_name)
        names = {n for n in names if n.upper() not in {"A", "B"}}
        return sorted(names, key=lambda n: n.lower())

    def _base_club_name(self, club_value: str) -> str:
        if not club_value:
            return ""
        value = club_value.strip()
        base = re.sub(r"(?:\s*(?:--|-)\s*|\s+)([AB])$", "", value, flags=re.IGNORECASE).strip()
        return base if base else value

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
        summary_sheets = self._cache.get("summary_sheets", []) if self._cache else []
        for sheet in summary_sheets:
            df = self._cache["sheets"][sheet].fillna("")
            club_col = None
            name_col = None
            points_col = None
            gender_col = None
            team_col = None
            for c in df.columns:
                c_lower = str(c).strip().lower()
                if "club" in c_lower and club_col is None:
                    club_col = c
                elif c_lower == "name" and name_col is None:
                    name_col = c
                elif "total" in c_lower and "point" in c_lower and points_col is None:
                    points_col = c
                elif c_lower == "points" and points_col is None:
                    points_col = c
                elif c_lower == "gender" and gender_col is None:
                    gender_col = c
                elif c_lower == "team" and team_col is None:
                    team_col = c
            if club_col is None or name_col is None or points_col is None:
                continue

            inferred_gender = ""
            if str(sheet).strip().lower() == "male":
                inferred_gender = "M"
            elif str(sheet).strip().lower() == "female":
                inferred_gender = "F"

            for _, row in df.iterrows():
                rc = str(row.get(club_col, "")).strip()
                if rc.lower() != club.lower():
                    continue
                points_value = str(row.get(points_col, "")).strip()
                points_int = 0
                try:
                    points_int = int(float(points_value))
                except Exception:
                    points_int = 0
                rows.append({
                    "Name": str(row.get(name_col, "")).strip(),
                    "Gender": str(row.get(gender_col, "")).strip() if gender_col is not None else inferred_gender,
                    "Category": str(row.get("Category", "")).strip(),
                    "Team": str(row.get(team_col, "")).strip() if team_col is not None else "",
                    "Points": str(points_int),
                })
        rows.sort(
            key=lambda entry: (
                -(int(entry["Points"])) if entry["Points"].isdigit() else 0,
                str(entry["Name"]).lower(),
            )
        )

        if not rows:
            self._club_summary_label.setText(f"Club '{club}' not found in history workbook.")
            self._club_info_label.setText("Club info will appear here.")
            self._club_table.clear()
            self._club_table.setRowCount(0)
            self._club_table.setColumnCount(0)
            return

        self._club_table.clear()
        display_cols = ["Name", "Gender", "Category", "Points"]
        if any(entry.get("Team") for entry in rows):
            display_cols.insert(1, "Team")
        self._club_table.setColumnCount(len(display_cols))
        self._club_table.setRowCount(len(rows))
        self._club_table.setHorizontalHeaderLabels(display_cols)
        for row_index, entry in enumerate(rows):
            for col_index, key in enumerate(display_cols):
                value = entry.get(key, "")
                if key == "Time":
                    value = self._format_time(value)
                item = QTableWidgetItem(value)
                if key == "Race":
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                elif key.lower() in {"gender", "category", "team"}:
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                elif key == "Points":
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
        self._update_club_info(club, rows)

    def _club_summary_rows(self, club: str) -> list[dict]:
        rows = []
        if self._cache is None:
            return rows

        summary_sheets = self._cache.get("summary_sheets", [])
        for sheet in summary_sheets:
            df = self._cache["sheets"][sheet].fillna("")
            club_col = None
            name_col = None
            points_col = None
            gender_col = None
            team_col = None
            for c in df.columns:
                c_lower = str(c).strip().lower()
                if "club" in c_lower and club_col is None:
                    club_col = c
                elif c_lower == "name" and name_col is None:
                    name_col = c
                elif "total" in c_lower and "point" in c_lower and points_col is None:
                    points_col = c
                elif c_lower == "points" and points_col is None:
                    points_col = c
                elif c_lower == "gender" and gender_col is None:
                    gender_col = c
                elif c_lower == "team" and team_col is None:
                    team_col = c
            if club_col is None or name_col is None or points_col is None:
                continue

            inferred_gender = ""
            if str(sheet).strip().lower() == "male":
                inferred_gender = "M"
            elif str(sheet).strip().lower() == "female":
                inferred_gender = "F"

            for _, row in df.iterrows():
                rc = str(row.get(club_col, "")).strip()
                if rc.lower() != club.lower():
                    continue
                points_value = str(row.get(points_col, "")).strip()
                points_int = 0
                try:
                    points_int = int(float(points_value))
                except Exception:
                    points_int = 0
                rows.append({
                    "Name": str(row.get(name_col, "")).strip(),
                    "Gender": str(row.get(gender_col, "")).strip() if gender_col is not None else inferred_gender,
                    "Category": str(row.get("Category", "")).strip(),
                    "Team": str(row.get(team_col, "")).strip() if team_col is not None else "",
                    "Points": str(points_int),
                })
        return rows

    def _club_summary_html(self, club: str, rows: list[dict]) -> str:
        division_stats = self._division_scores_for_club(club)
        best_six_points = sum(
            sorted(
                (int(entry["Points"]) for entry in rows if str(entry.get("Points", "")).strip().replace(" ", "") and self._is_number(entry["Points"])),
                reverse=True,
            )[:6]
        )
        total_points = sum(
            int(float(entry["Points"])) for entry in rows if str(entry.get("Points", "")).strip().replace(" ", "") and self._is_number(entry["Points"])
        )
        team_a_place_text = ""
        if division_stats.get("team_a_division") is not None and division_stats.get("team_a_place") is not None:
            team_a_place_text = f", Div {division_stats['team_a_division']} - {self._ordinal(division_stats['team_a_place'])} Place"
        team_b_place_text = ""
        if division_stats.get("team_b_division") is not None and division_stats.get("team_b_place") is not None:
            team_b_place_text = f", Div {division_stats['team_b_division']} - {self._ordinal(division_stats['team_b_place'])} Place"

        summary_lines = [
            f"<div style='margin-bottom: 6px;'><b>Club:</b> {club}</div>",
            f"<div style='margin-bottom: 6px;'><b>Runners:</b> {len({entry.get('Name', '') for entry in rows if entry.get('Name')})}</div>",
            f"<div style='margin-bottom: 6px;'><b>Races:</b> {division_stats['race_count']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Male score:</b> {division_stats['male_score']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Female score:</b> {division_stats['female_score']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Combined score:</b> {division_stats['combined_score']}</div>",
            f"<div style='margin-bottom: 6px;'><b>Team A points:</b> {division_stats['team_a_best6']} ({division_stats['team_a_points_total']}){team_a_place_text}</div>",
            f"<div><b>Team B points:</b> {division_stats['team_b_best6']} ({division_stats['team_b_points_total']}){team_b_place_text}</div>",
        ]
        return "".join(summary_lines)

    def _club_best_html(self, rows: list[dict]) -> str:
        best_by_category: dict[str, tuple[str, int]] = {}
        valid_categories = {"jun", "sen", "v40", "v50", "v60", "v70"}
        for entry in rows:
            points = 0
            try:
                points = int(float(entry.get("Points", 0)))
            except Exception:
                continue
            gender = str(entry.get("Gender", "")).strip().lower()
            category = str(entry.get("Category", "")).strip().lower()
            if category in valid_categories and (gender.startswith("m") or gender.startswith("f")):
                key = (category, "m" if gender.startswith("m") else "f")
                existing = best_by_category.get(key)
                if existing is None or points > existing[1]:
                    best_by_category[key] = (str(entry.get("Name", "")), points)

        best_text = []
        for (category, gender), (name, points) in sorted(best_by_category.items(), key=lambda x: (x[0][0].lower(), x[0][1])):
            gender_label = "Male" if gender == "m" else "Female"
            best_text.append(
                f"<div style='margin-bottom: 4px;'><b>{category.upper()} {gender_label}:</b> {name} ({points})</div>"
            )
        return "".join(best_text) if best_text else "No category best runners found."

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
        summary_rows = self._club_summary_rows(club)
        if summary_rows:
            self._club_races_info_label.setText(self._club_summary_html(club, summary_rows))
        else:
            self._club_races_info_label.setText("Club info will appear here.")

        self._club_races_list.clear()
        self._club_race_details = {}
        self._load_events_schedule()

        race_team_totals: dict[tuple[str, str], int] = {}
        for sheet in self._cache["race_sheets"]:
            df = self._cache["sheets"][sheet].fillna("")
            club_col = None
            name_col = None
            points_col = None
            time_col = None
            team_col = None
            gender_col = None
            position_col = None
            for c in df.columns:
                c_lower = str(c).strip().lower()
                if c_lower == "club" and club_col is None:
                    club_col = c
                elif c_lower == "affiliation" and club_col is None:
                    club_col = c
                elif c_lower == "team" and team_col is None:
                    team_col = c
                elif c_lower in {"gender", "sex"} and gender_col is None:
                    gender_col = c
                elif c_lower in {"pos", "position", "overall pos", "overall position"} and position_col is None:
                    position_col = c
                if c_lower == "name" and name_col is None:
                    name_col = c
                if c_lower == "points" and points_col is None:
                    points_col = c
                if c_lower == "time" and time_col is None:
                    time_col = c
            if club_col is None or name_col is None:
                continue

            race_num = extract_race_number(sheet) or ""
            display_title = self._race_title_for_sheet(sheet)
            sort_key = int(race_num) if str(race_num).isdigit() else sheet.lower()
            team_rows = {"A": [], "B": []}
            found = False
            for row_index, row in df.iterrows():
                rc = str(row.get(club_col, "")).strip()
                team_value = str(row.get(team_col, "")).strip() if team_col is not None else ""
                team_id = self._club_row_team_id(club, rc, team_value)
                if team_id is None:
                    continue
                found = True
                points = str(row.get(points_col, "")).strip() if points_col is not None else ""
                score = 0
                try:
                    score = int(float(points))
                except Exception:
                    pass
                key = (race_num, team_id)
                race_team_totals[key] = race_team_totals.get(key, 0) + score

                name = str(row.get(name_col, "")).strip()
                time_value = str(row.get(time_col, "")).strip() if time_col is not None else ""
                position = None
                if position_col is not None:
                    try:
                        position = int(float(row.get(position_col, 0)))
                    except Exception:
                        position = None
                gender_value = str(row.get(gender_col, "")).strip() if gender_col is not None else ""
                gender = "M" if gender_value.strip().upper().startswith("M") else "F" if gender_value.strip().upper().startswith("F") else ""
                runner = {
                    "Pos": str(position) if position is not None else "",
                    "Name": name,
                    "Time": time_value,
                    "Points": str(score),
                    "Gender": gender,
                }
                team_rows[team_id].append(runner)
            if not found:
                continue

            self._club_race_details[sheet] = {
                "display_text": display_title,
                "sort_key": sort_key,
                "team_rows": team_rows,
            }

        if not race_team_totals:
            self._club_races_summary_label.setText(f"Club '{club}' not found in history workbook.")
            self._club_races_info_label.setText("Club summary info will appear here.")
            self._club_races_list.clear()
            self._club_races_table.clear()
            self._club_races_table.setRowCount(0)
            self._club_races_table.setColumnCount(0)
            return

        total_points = sum(points for points in race_team_totals.values())
        self._club_races_summary_label.setText(
            f"{club}: {len(self._club_race_details)} races, total points {total_points}."
        )

        if self._club_race_details:
            for sheet, details in sorted(self._club_race_details.items(), key=lambda kv: kv[1]["sort_key"]):
                item = QListWidgetItem(details["display_text"])
                item.setData(Qt.UserRole, sheet)
                self._club_races_list.addItem(item)
            if self._club_races_list.count() > 0:
                self._club_races_list.setCurrentRow(0)
                selected_item = self._club_races_list.currentItem()
                if selected_item is not None:
                    self._on_club_race_selected(selected_item)
        else:
            self._club_races_table.clear()
            self._club_races_table.setRowCount(0)
            self._club_races_table.setColumnCount(0)

    def _race_title_for_sheet(self, sheet: str) -> str:
        race_num = extract_race_number(sheet)
        if race_num is not None and self._event_title_by_race.get(race_num):
            return self._event_title_by_race[race_num]
        return sheet

    def _on_club_race_selected(self, item: QListWidgetItem) -> None:
        sheet = item.data(Qt.UserRole)
        if not sheet:
            return
        self._display_club_race_details(sheet)

    def _display_club_race_details(self, sheet: str) -> None:
        details = self._club_race_details.get(sheet)
        if details is None:
            return

        team_a = sorted(details["team_rows"].get("A", []), key=lambda r: (int(r["Pos"]) if r["Pos"].isdigit() else float("inf"), r["Name"].lower()))
        team_b = sorted(details["team_rows"].get("B", []), key=lambda r: (int(r["Pos"]) if r["Pos"].isdigit() else float("inf"), r["Name"].lower()))

        title = details["display_text"]
        count_a = len(team_a)
        count_b = len(team_b)
        self._club_races_summary_label.setText(
            f"{title}: Team A {count_a} runners, Team B {count_b} runners"
        )

        rows = []
        def add_team_rows(team_label: str, runners: list[dict]) -> None:
            if not runners:
                return
            males = [r for r in runners if r.get("Gender", "").upper() == "M"]
            females = [r for r in runners if r.get("Gender", "").upper() == "F"]
            males = sorted(males, key=lambda r: (int(r["Pos"]) if r["Pos"].isdigit() else float("inf"), r["Name"].lower()))[:5]
            females = sorted(females, key=lambda r: (int(r["Pos"]) if r["Pos"].isdigit() else float("inf"), r["Name"].lower()))[:5]
            rows.append({"Pos": "", "Name": team_label, "Time": "", "Points": ""})
            if males:
                rows.append({"Pos": "", "Name": "Male", "Time": "", "Points": ""})
                for runner in males:
                    rows.append({
                        "Pos": runner["Pos"],
                        "Name": runner["Name"],
                        "Time": self._format_time(runner["Time"]),
                        "Points": runner["Points"],
                    })
                rows.append({"Pos": "", "Name": "", "Time": "", "Points": ""})
            if females:
                rows.append({"Pos": "", "Name": "Female", "Time": "", "Points": ""})
                for runner in females:
                    rows.append({
                        "Pos": runner["Pos"],
                        "Name": runner["Name"],
                        "Time": self._format_time(runner["Time"]),
                        "Points": runner["Points"],
                    })
        add_team_rows("Team A", team_a)
        if team_a and team_b:
            rows.append({"Pos": "", "Name": "", "Time": "", "Points": ""})
        add_team_rows("Team B", team_b)

        self._club_races_table.clear()
        cols = ["Pos", "Name", "Time", "Points"]
        self._club_races_table.setColumnCount(len(cols))
        self._club_races_table.setRowCount(len(rows))
        self._club_races_table.setHorizontalHeaderLabels(cols)

        for row_index, entry in enumerate(rows):
            for col_index, key in enumerate(cols):
                value = entry.get(key, "")
                item = QTableWidgetItem(value)
                if key == "Pos":
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                elif key == "Points":
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                if entry["Name"] in {"Team A", "Team B"}:
                    item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                self._club_races_table.setItem(row_index, col_index, item)
        self._club_races_table.resizeColumnsToContents()

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
        self._runner_info_label.setText("Athlete info will appear here.")
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

        self._update_runner_info(df)

        conflicts = detect_conflicts(df)
        if conflicts:
            self._summary_label.setText(self._summary_label.text() + "  Conflicts detected: " + ", ".join(conflicts.keys()))

    def _populate_table(self, df) -> None:
        self._table.clear()
        display_cols = [col for col in df.columns if str(col).strip().lower() != "source"]
        self._table.setColumnCount(len(display_cols))
        self._table.setRowCount(len(df.index))
        self._table.setHorizontalHeaderLabels([str(col) for col in display_cols])

        numeric_cols = set()
        for col in display_cols:
            try:
                pd_col = df[col]
                numeric_vals = pd.to_numeric(pd_col, errors="coerce")
                if numeric_vals.notna().all():
                    numeric_cols.add(col)
            except Exception:
                continue

        for row_index, row in enumerate(df[display_cols].itertuples(index=False, name=None)):
            for col_index, value in enumerate(row):
                key = display_cols[col_index]
                text = str(value) if value is not None else ""
                if str(key).strip().lower() == "time":
                    text = self._format_time(value)
                item = QTableWidgetItem(text)
                key_lower = str(key).strip().lower()
                if key_lower == "race":
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                elif key_lower in {"gender", "category", "team"}:
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                elif key in numeric_cols:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self._table.setItem(row_index, col_index, item)

        self._table.resizeColumnsToContents()
        self._table.resizeRowsToContents()

    def _copy_results(self) -> None:
        if self._current_df is None or self._current_df.empty:
            QMessageBox.information(self, "Copy Results", "No runner history loaded.")
            return

        df = sanitise_df_for_export(self._current_df)
        text = df.to_string(index=False)
        QGuiApplication.clipboard().setText(text)
        QMessageBox.information(self, "Copy Results", "Runner history copied to clipboard.")

    def _copy_table_widget(self, table: QTableWidget, title: str) -> None:
        row_count = table.rowCount()
        col_count = table.columnCount()
        if row_count == 0 or col_count == 0:
            QMessageBox.information(self, "Copy Results", f"No {title} data to copy.")
            return

        headers = [table.horizontalHeaderItem(c).text() if table.horizontalHeaderItem(c) else "" for c in range(col_count)]
        lines = ["\t".join(headers)]
        for r in range(row_count):
            values = []
            empty_row = True
            for c in range(col_count):
                item = table.item(r, c)
                text = item.text() if item is not None else ""
                values.append(text)
                if text.strip():
                    empty_row = False
            if not empty_row:
                lines.append("\t".join(values))

        QGuiApplication.clipboard().setText("\n".join(lines))
        QMessageBox.information(self, "Copy Results", f"{title} copied to clipboard.")
