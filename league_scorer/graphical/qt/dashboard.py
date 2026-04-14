"""Qt dashboard implementation for WRRL Admin Suite."""

import os
import queue
import re
import subprocess
import threading
import sys
from pathlib import Path
from PySide6.QtCore import QTimer, QSettings, QUrl, Qt, QByteArray, Slot
from PySide6.QtGui import QDesktopServices, QFont, QIcon, QPixmap, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpacerItem,
    QSizePolicy,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from league_scorer.session_config import config as session_config
from league_scorer.events_loader import load_events
from league_scorer.graphical.sporthive_helpers import is_sporthive_event_summary_url
from league_scorer.graphical.qt.events_viewer import EventsViewerWindow
from league_scorer.graphical.qt.raes_window import RAESWindow
from league_scorer.graphical.qt.runner_history import RunnerHistoryWindow
from league_scorer.graphical.qt.raw_archive_diff_viewer import RawArchiveDiffWindow
from league_scorer.graphical.qt.results_viewer import ResultsViewerWindow
from league_scorer.graphical.qt.view_autopilot import AutopilotReportsWindow
from league_scorer.input.common_files import race_discovery_exclusions
from league_scorer.input.source_loader import discover_race_files
from league_scorer.output.output_layout import build_output_paths, ensure_output_subdirs, sort_existing_output_files
from league_scorer.settings import settings, DEFAULT_SETTINGS
from league_scorer.process.main import LeagueScorer
from league_scorer.raceroster_import import (
    SporthiveRaceNotDirectlyImportableError,
    import_raceroster_results,
    import_sporthive_manual_pages,
)


def _find_repository_root() -> Path:
    candidate = Path(__file__).resolve()
    for parent in candidate.parents:
        if (parent / "scripts").is_dir() and (parent / "league_scorer").is_dir():
            return parent
    return candidate.parents[3]

WRRL_NAVY = "#3a4658"
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"
WRRL_AMBER = "#e6a817"
WRRL_AMBER_LIGHT = "#f7e082"


class ActionCard(QFrame):
    def __init__(self, title: str, subtitle: str, callback, tone: str = "secondary", icon: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._callback = callback
        self._tone = tone
        self._icon = icon
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("actionCard")
        self.setFixedHeight(140)
        self._title = title
        self._subtitle = subtitle
        self._build_ui()

    def _build_ui(self) -> None:
        palette = {
            "primary": {
                "bg": WRRL_GREEN,
                "fg_title": WRRL_WHITE,
                "fg_subtitle": "#d9efe2",
                "hover": "#24653d",
                "border": "#1f5632",
            },
            "secondary": {
                "bg": "#f8fbff",
                "fg_title": WRRL_NAVY,
                "fg_subtitle": "#66788d",
                "hover": "#eef5fb",
                "border": "#edf2f7",
            },
        }[self._tone]

        self._bg = palette["bg"]
        hover_bg = palette["hover"]
        self.setStyleSheet(
            f"background: {self._bg}; border: 1px solid transparent; border-radius: 16px;"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(3)

        header = QWidget(self)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(6)

        if self._icon:
            icon_lbl = QLabel(self._icon, self)
            icon_lbl.setFont(QFont("Segoe UI", 14))
            icon_lbl.setStyleSheet(f"color: {palette['fg_title']};")
            header_layout.addWidget(icon_lbl, 0, Qt.AlignTop)

        title_lbl = QLabel(self._title, self)
        title_lbl.setWordWrap(True)
        title_lbl.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
        title_lbl.setStyleSheet(
            f"color: {palette['fg_title']};"
        )
        header_layout.addWidget(title_lbl, 1)
        layout.addWidget(header)

        subtitle_lbl = QLabel(self._subtitle, self)
        subtitle_lbl.setWordWrap(True)
        subtitle_lbl.setFont(QFont("Segoe UI", 9))
        subtitle_lbl.setStyleSheet(
            f"color: {palette['fg_subtitle']}; line-height: 1.3;"
        )
        layout.addWidget(subtitle_lbl)

        self._hover_bg = hover_bg
        self._title_lbl = title_lbl
        self._subtitle_lbl = subtitle_lbl

    def enterEvent(self, event) -> None:
        self.setStyleSheet(f"background: {self._hover_bg}; border: 1px solid transparent; border-radius: 16px;")
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        self.setStyleSheet(f"background: {self._bg}; border: 1px solid transparent; border-radius: 16px;")
        super().leaveEvent(event)

    def mousePressEvent(self, event) -> None:
        if callable(self._callback):
            self._callback()
        super().mousePressEvent(event)


class WorkflowDialog(QDialog):
    def __init__(self, title: str, initial_status: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumSize(640, 420)
        self.setMaximumWidth(900)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        self.status_label = QLabel(initial_status, self)
        self.status_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(self.status_label)

        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        self.output_text.setStyleSheet("background: #f5f5f5; color: #212121;")
        layout.addWidget(self.output_text, 1)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Close)
        self.buttons.button(QDialogButtonBox.Close).setText("Close")
        self.buttons.rejected.connect(self.reject)
        self.buttons.setEnabled(False)
        layout.addWidget(self.buttons)

    def append_output(self, message: str) -> None:
        self.output_text.append(message)
        self.output_text.moveCursor(QTextCursor.End)

    def set_status(self, message: str) -> None:
        self.status_label.setText(message)

    def set_finished(self, success: bool) -> None:
        self.buttons.setEnabled(True)
        if success:
            self.status_label.setText("Completed successfully.")
        else:
            self.status_label.setText("Finished with errors.")


class SettingsDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setModal(True)
        self.resize(720, 520)

        self._field_widgets: dict[str, QSpinBox] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        title = QLabel("Application Settings", self)
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        form.setHorizontalSpacing(20)
        form.setVerticalSpacing(12)
        layout.addLayout(form)

        self._data_root_display = QLineEdit(self)
        self._data_root_display.setReadOnly(True)
        self._data_root_display.setText(str(session_config.data_root) if session_config.data_root else "Not set")

        browse_btn = QPushButton("Browse…", self)
        browse_btn.clicked.connect(self._browse_data_root)
        data_root_row = QWidget(self)
        data_root_layout = QHBoxLayout(data_root_row)
        data_root_layout.setContentsMargins(0, 0, 0, 0)
        data_root_layout.addWidget(self._data_root_display, 1)
        data_root_layout.addWidget(browse_btn)

        form.addRow("Data Root:", data_root_row)

        self._year_picker = QSpinBox(self)
        self._year_picker.setRange(2020, 2100)
        self._year_picker.setValue(session_config.year)
        form.addRow("Season Year:", self._year_picker)

        self._events_path_display = QLineEdit(self)
        self._events_path_display.setReadOnly(True)
        self._events_path_display.setText(str(session_config.events_path) if session_config.events_path else "Not configured")

        events_browse_btn = QPushButton("Browse…", self)
        events_browse_btn.clicked.connect(self._browse_events_file)
        events_row = QWidget(self)
        events_layout = QHBoxLayout(events_row)
        events_layout.setContentsMargins(0, 0, 0, 0)
        events_layout.addWidget(self._events_path_display, 1)
        events_layout.addWidget(events_browse_btn)

        form.addRow("Events File:", events_row)

        settings_group = QWidget(self)
        settings_group_layout = QFormLayout(settings_group)
        settings_group_layout.setLabelAlignment(Qt.AlignLeft)
        settings_group_layout.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        settings_group_layout.setHorizontalSpacing(20)
        settings_group_layout.setVerticalSpacing(10)

        settings_label = QLabel("League settings", self)
        settings_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        layout.addWidget(settings_label)
        layout.addWidget(settings_group)

        for key, default_value in DEFAULT_SETTINGS.items():
            spinner = QSpinBox(self)
            spinner.setRange(1, 200)
            spinner.setValue(int(settings.get(key)))
            spinner.setFixedWidth(100)
            self._field_widgets[key] = spinner
            settings_group_layout.addRow(key.replace("_", " ").title() + ":", spinner)

        button_row = QWidget(self)
        button_layout = QHBoxLayout(button_row)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(10)

        save_btn = QPushButton("Save", self)
        save_btn.clicked.connect(self._save_settings)
        button_layout.addWidget(save_btn)

        close_btn = QPushButton("Close", self)
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)
        button_layout.addStretch(1)

        layout.addWidget(button_row)

    def _browse_data_root(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Data Root", str(Path.home()))
        if not folder:
            return

        session_config.data_root = Path(folder)
        session_config.ensure_dirs()
        self._data_root_display.setText(folder)
        self._events_path_display.setText(str(session_config.events_path) if session_config.events_path else "Not configured")

    def _browse_events_file(self) -> None:
        if not session_config.data_root:
            QMessageBox.warning(
                self,
                "Data Root Required",
                "Set Data Root before selecting an events file.",
            )
            return

        if session_config.input_dir is None:
            session_config.ensure_dirs()

        control_dir = session_config.control_dir
        initial_dir = str(control_dir) if control_dir and control_dir.exists() else str(Path.home())

        events_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Events Spreadsheet",
            initial_dir,
            "Excel Files (*.xlsx *.xls);;All Files (*)",
        )
        if not events_path:
            return

        session_config.events_path = Path(events_path)
        self._events_path_display.setText(str(session_config.events_path))

    def _save_settings(self) -> None:
        year = int(self._year_picker.value())
        if year != session_config.year:
            session_config.year = year
            session_config.ensure_dirs()

        for key, widget in self._field_widgets.items():
            settings.set(key, int(widget.value()))

        QMessageBox.information(self, "Settings Saved", "Settings have been saved successfully.")


class ManualPageDialog(QDialog):
    def __init__(self, page_number: int, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Sporthive Page {page_number}")
        self.setModal(True)
        self.resize(760, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        prompt = QLabel(
            f"Paste results table text for Sporthive page {page_number}. Copy the rows shown on screen and click Use This Page.",
            self,
        )
        prompt.setWordWrap(True)
        prompt.setFont(QFont("Segoe UI", 10))
        layout.addWidget(prompt)

        self.text_box = QTextEdit(self)
        self.text_box.setFont(QFont("Consolas", 10))
        layout.addWidget(self.text_box, 1)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        cancel_btn = QPushButton("Cancel", self)
        cancel_btn.clicked.connect(self.reject)
        use_btn = QPushButton("Use This Page", self)
        use_btn.clicked.connect(self.accept)
        buttons.addStretch(1)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(use_btn)
        layout.addLayout(buttons)

    def page_text(self) -> str:
        return self.text_box.toPlainText().strip()


class RaceRosterDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Import From Race Roster")
        self.setModal(True)
        self.resize(520, 340)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setFormAlignment(Qt.AlignLeft)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(10)

        self.url_edit = QLineEdit(self)
        form.addRow("Race URL:", self.url_edit)

        self.race_number_spin = QSpinBox(self)
        self.race_number_spin.setRange(1, 99)
        self.race_number_spin.setValue(1)
        form.addRow("Race number:", self.race_number_spin)

        self.race_name_edit = QLineEdit(self)
        form.addRow("Race name:", self.race_name_edit)

        self.sporthive_hint_spin = QSpinBox(self)
        self.sporthive_hint_spin.setRange(0, 9999)
        self.sporthive_hint_spin.setToolTip("Optional Sporthive race ID for event-summary links.")
        form.addRow("Sporthive race ID:", self.sporthive_hint_spin)

        layout.addLayout(form)

        self.hint_label = QLabel(
            "For Sporthive event-summary URLs, enter the race ID from the 'View results' URL.",
            self,
        )
        self.hint_label.setWordWrap(True)
        self.hint_label.setStyleSheet("color: #555555;")
        self.hint_label.setVisible(False)
        layout.addWidget(self.hint_label)

        button_row = QHBoxLayout()
        button_row.setSpacing(10)
        button_row.addStretch(1)

        cancel_btn = QPushButton("Cancel", self)
        cancel_btn.clicked.connect(self.reject)
        button_row.addWidget(cancel_btn)

        import_btn = QPushButton("Import", self)
        import_btn.clicked.connect(self.accept)
        button_row.addWidget(import_btn)

        layout.addLayout(button_row)

        self.url_edit.textChanged.connect(self._update_hint_visibility)
        self._update_hint_visibility(self.url_edit.text())

    def _update_hint_visibility(self, text: str) -> None:
        self.hint_label.setVisible(is_sporthive_event_summary_url(text))

    def get_values(self) -> tuple[str, int, str | None, int | None]:
        url = self.url_edit.text().strip()
        race_number = self.race_number_spin.value()
        race_name = self.race_name_edit.text().strip() or None
        sporthive_hint = self.sporthive_hint_spin.value() or None
        return url, race_number, race_name, sporthive_hint


class QtLeagueScorerDashboard(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._window_settings = QSettings("WRRL", "AdminSuite")
        self.setWindowTitle("WRRL Admin Suite")
        self.resize(1080, 720)
        self.setMinimumSize(860, 640)
        self._restore_window_geometry()
        self.shield_pixmap = None
        self.waa_pixmap = None
        self._events_viewer = None
        self._results_viewer = None
        self._runner_history_viewer = None
        self._raes_window = None
        self._load_logos()
        self._build_ui()
        self._refresh_config_panel()

    def _restore_window_geometry(self) -> None:
        geometry = self._window_settings.value("dashboardGeometry")
        if isinstance(geometry, (bytes, QByteArray)) and geometry:
            self.restoreGeometry(geometry)

    def closeEvent(self, event) -> None:
        self._window_settings.setValue("dashboardGeometry", self.saveGeometry())
        super().closeEvent(event)

    def _load_logos(self) -> None:
        images_dir = Path(__file__).resolve().parents[2] / "images"
        shield_path = images_dir / "WRRL shield concept.png"
        waa_path = images_dir / "WRRL_logo-629x400.png"

        if shield_path.exists():
            pix = QPixmap(str(shield_path))
            self.shield_pixmap = pix.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        if waa_path.exists():
            pix = QPixmap(str(waa_path))
            self.waa_pixmap = pix.scaled(100, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setObjectName("dashboard")
        central.setStyleSheet("background: %s;" % WRRL_LIGHT)
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._create_header())
        root_layout.addWidget(self._create_accent_bar())
        root_layout.addWidget(self._create_main_content())
        root_layout.addWidget(self._create_footer())

    def _create_header(self) -> QWidget:
        header = QWidget(self)
        header.setStyleSheet(f"background: {WRRL_NAVY};")
        header.setMinimumHeight(130)
        header.setMaximumHeight(160)
        header.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QHBoxLayout(header)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)

        if self.shield_pixmap is not None:
            shield_label = QLabel(header)
            shield_label.setPixmap(self.shield_pixmap)
            layout.addWidget(shield_label)

        title_section = QWidget(header)
        title_layout = QVBoxLayout(title_section)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(6)

        title = QLabel("WRRL Admin Suite", title_section)
        title.setStyleSheet("color: %s;" % WRRL_WHITE)
        title.setFont(QFont("Segoe UI", 28, QFont.Bold))
        title_layout.addWidget(title)

        subtitle = QLabel(
            "Wiltshire Road and Running League administration, scoring and publish workflows",
            title_section,
        )
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #a0b0c0;")
        subtitle.setFont(QFont("Segoe UI", 10))
        title_layout.addWidget(subtitle)

        self.freshness_label = QLabel("Ready", title_section)
        self.freshness_label.setStyleSheet("color: #d7dfeb;")
        self.freshness_label.setFont(QFont("Segoe UI", 9))
        title_layout.addWidget(self.freshness_label)

        layout.addWidget(title_section, 1)

        right_section = QWidget(header)
        right_layout = QVBoxLayout(right_section)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)
        right_layout.setAlignment(Qt.AlignTop)

        if self.waa_pixmap is not None:
            waa_label = QLabel(right_section)
            waa_label.setPixmap(self.waa_pixmap)
            right_layout.addWidget(waa_label, alignment=Qt.AlignRight)

        help_btn = QPushButton("Help", right_section)
        help_btn.setCursor(Qt.PointingHandCursor)
        help_btn.setStyleSheet(
            "QPushButton { color: #a0b0c0; background: transparent; border: none; font: 9pt 'Segoe UI'; } "
            "QPushButton:hover { color: #ffffff; }"
        )
        help_btn.clicked.connect(self._on_help)
        right_layout.addWidget(help_btn, alignment=Qt.AlignRight)

        layout.addWidget(right_section)

        return header

    def _create_accent_bar(self) -> QWidget:
        accent = QWidget(self)
        accent.setFixedHeight(4)
        accent.setStyleSheet(f"background: {WRRL_GREEN};")
        return accent

    def _create_main_content(self) -> QWidget:
        content = QWidget(self)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(18)

        layout.addWidget(self._create_home_page())
        layout.addWidget(self._create_race_overview())

        return content

    def _create_config_panel(self) -> QWidget:
        panel = QFrame(self)
        panel.setFrameShape(QFrame.StyledPanel)
        panel.setStyleSheet(f"background: {WRRL_NAVY}; border-radius: 14px;")

        layout = QHBoxLayout(panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(40)

        left = QWidget(panel)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        left_layout.addWidget(self._create_season_picker())

        self.freshness_label = QLabel("Data status unknown", left)
        self.freshness_label.setStyleSheet("color: #d7dfeb;")
        self.freshness_label.setFont(QFont("Segoe UI", 9))
        left_layout.addWidget(self.freshness_label)

        layout.addWidget(left, 0)

        right = QWidget(panel)
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        self.input_path_row = self._create_path_row("Input folder:", self._input_path_text())
        self.output_path_row = self._create_path_row("Output folder:", self._output_path_text())
        right_layout.addWidget(self.input_path_row)
        right_layout.addWidget(self.output_path_row)

        layout.addWidget(right, 1)

        return panel

    def _create_season_picker(self) -> QWidget:
        picker = QWidget(self)
        layout = QHBoxLayout(picker)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        prev_btn = QPushButton("←", picker)
        prev_btn.setCursor(Qt.PointingHandCursor)
        prev_btn.setFixedSize(32, 32)
        prev_btn.setStyleSheet(
            "QPushButton { background: transparent; color: %s; border: 1px solid #506275; border-radius: 6px; font: 14pt 'Segoe UI'; } "
            "QPushButton:hover { background: rgba(255,255,255,0.1); }"
            % WRRL_WHITE
        )
        prev_btn.clicked.connect(self._on_year_prev)
        layout.addWidget(prev_btn)

        self.year_label = QLabel(str(session_config.year), picker)
        self.year_label.setAlignment(Qt.AlignCenter)
        self.year_label.setFixedWidth(96)
        self.year_label.setStyleSheet("color: %s;" % WRRL_WHITE)
        self.year_label.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(self.year_label)

        next_btn = QPushButton("→", picker)
        next_btn.setCursor(Qt.PointingHandCursor)
        next_btn.setFixedSize(32, 32)
        next_btn.setStyleSheet(
            "QPushButton { background: transparent; color: %s; border: 1px solid #506275; border-radius: 6px; font: 14pt 'Segoe UI'; } "
            "QPushButton:hover { background: rgba(255,255,255,0.1); }"
            % WRRL_WHITE
        )
        next_btn.clicked.connect(self._on_year_next)
        layout.addWidget(next_btn)

        return picker

    def _create_home_page(self) -> QWidget:
        home = QWidget(self)
        home_layout = QVBoxLayout(home)
        home_layout.setContentsMargins(0, 0, 0, 0)
        home_layout.setSpacing(16)

        title = QLabel("Quick Actions", home)
        title.setStyleSheet("color: %s;" % WRRL_NAVY)
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        home_layout.addWidget(title)

        grid = QWidget(home)
        grid_layout = QGridLayout(grid)
        grid_layout.setContentsMargins(0, 0, 0, 0)
        grid_layout.setSpacing(14)

        cards = [
            ("Run Autopilot", "Run audit, safe auto-fixes, and staged checks", self._on_run_autopilot, "primary", "▶️"),
            ("View Autopilot Report", "Open latest autopilot summary report", self._on_view_autopilot_report, "secondary", "📄"),
            ("Data Corrections (RAES)", "Review and apply runner-level corrections using the RAES editor.", self._on_review_raes, "secondary", "🛠️"),
            ("Publish Results", "Publish final results from audited files, PDFs and club reports.", self._on_publish_results, "primary", "📤"),
            ("View Results", "Open generated standings and published reports.", self._on_view_results, "secondary", "📊"),
            ("Compare Raw vs Archive", "Inspect line-by-line changes against raw archive data.", self._on_compare_raw_archive, "secondary", "🧾"),
            ("Export Published PDFs", "Copy all published PDFs into a single export folder.", self._on_export_published_pdfs, "secondary", "📁"),
            ("View Events", "Browse loaded events schedule.", self._on_view_events, "secondary", "📅"),
            ("Runner/Club Enquiry", "Search published results by runner or club.", self._on_view_runner_history, "secondary", "🔍"),
        ]

        for index, (title, subtitle, callback, tone, icon) in enumerate(cards):
            row = index // 3
            col = index % 3
            grid_layout.addWidget(ActionCard(title, subtitle, callback, tone, icon, grid), row, col)

        home_layout.addWidget(grid)

        bottom_actions = QWidget(home)
        bottom_layout = QHBoxLayout(bottom_actions)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(12)

        for label, handler in [
            ("⚙️ Settings", self._on_settings),
            ("▶ Run Scorer", self._on_run_scorer),
            ("Publish Provisional", self._on_run_provisional_fast_track),
        ]:
            btn = QPushButton(label, bottom_actions)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(
                "QPushButton { color: %s; background: transparent; border: 1px solid #ccd7e3; border-radius: 10px; padding: 10px 16px; font: 10pt 'Segoe UI'; } "
                "QPushButton:hover { background: #eef2f7; }"
                % WRRL_NAVY
            )
            btn.clicked.connect(handler)
            bottom_layout.addWidget(btn)

        bottom_layout.addStretch(1)
        home_layout.addWidget(bottom_actions)

        return home

    def _create_path_row(self, label_text: str, path_text: str) -> QWidget:
        row = QWidget(self)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        label = QLabel(label_text, row)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        label.setStyleSheet("color: %s;" % WRRL_WHITE)
        layout.addWidget(label)

        path = QLabel(path_text, row)
        path.setFont(QFont("Segoe UI", 10))
        path.setWordWrap(True)
        path.setStyleSheet("color: #d7dfeb;")
        layout.addWidget(path)

        row.path_label = path
        return row

    def _input_path_text(self) -> str:
        input_dir = session_config.input_dir
        return str(input_dir) if input_dir else "Not set"

    def _output_path_text(self) -> str:
        output_dir = session_config.output_dir
        return str(output_dir) if output_dir else "Not set"

    def _discover_race_names(self, input_dir: Path | None) -> list[str]:
        if not input_dir or not input_dir.is_dir():
            return []
        discovered = discover_race_files(input_dir, excluded_names=race_discovery_exclusions())
        formatted = []
        for path in discovered.values():
            stem = path.stem
            cleaned = re.sub(r"^race\s*#?\s*", "", stem, flags=re.IGNORECASE).strip()
            if len(cleaned) > 30:
                cleaned = cleaned[:27].rstrip() + "..."
            formatted.append(cleaned)
        return formatted

    def _create_race_overview(self) -> QWidget:
        panel = QFrame(self)
        panel.setFrameShape(QFrame.StyledPanel)
        panel.setStyleSheet("background: #ffffff; border-radius: 14px;")
        panel.setMaximumHeight(180)
        panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        title = QLabel("Race discovery overview", panel)
        title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        title.setStyleSheet("color: %s;" % WRRL_NAVY)
        layout.addWidget(title)

        path = session_config.raw_data_dir
        races = self._discover_race_names(path) if path else []
        overview_text = "No race files found." if not races else f"Discovered {len(races)} race files."

        overview = QLabel(overview_text, panel)
        overview.setFont(QFont("Segoe UI", 10))
        overview.setStyleSheet("color: #334455;")
        layout.addWidget(overview)

        if races:
            names_preview = ", ".join(races[:10])
            if len(races) > 10:
                names_preview += f" and {len(races) - 10} more"

            details = QLabel(names_preview, panel)
            details.setFont(QFont("Segoe UI", 9))
            details.setStyleSheet("color: #5a6878;")
            details.setWordWrap(True)
            layout.addWidget(details)

        return panel

    def _create_footer(self) -> QWidget:
        footer = QWidget(self)
        footer.setStyleSheet(f"background: {WRRL_NAVY};")
        footer.setFixedHeight(48)

        layout = QHBoxLayout(footer)
        layout.setContentsMargins(24, 0, 24, 0)

        from league_scorer import __version__
        text = QLabel(f"© 2026 Wiltshire Athletics Assoc. | WRRL Admin Suite v{__version__}", footer)
        text.setStyleSheet("color: #707080;")
        text.setFont(QFont("Segoe UI", 9))
        layout.addWidget(text)
        layout.addStretch(1)

        return footer

    def _refresh_config_panel(self) -> None:
        if hasattr(self, "year_label"):
            self.year_label.setText(str(session_config.year))

        if not session_config.data_root:
            status = "Data root not configured"
        elif session_config.input_dir is None or session_config.output_dir is None:
            status = "Configuring season folders..."
        else:
            status = "Ready"
        if hasattr(self, "freshness_label"):
            self.freshness_label.setText(status)
        QTimer.singleShot(1500, self._refresh_config_panel)

    def _require_configured(self, action: str) -> bool:
        if not session_config.is_configured:
            QMessageBox.warning(
                self,
                "Not Configured",
                f"Please set a Data Root folder before using {action}.\n\nUse Settings (⚙️) to configure your data paths.",
            )
            return False
        return True

    def _open_path_in_system(self, path: Path) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _run_scorer(self) -> None:
        if not self._require_configured("Run WRRL Admin Suite"):
            return
        if session_config.input_dir is None or session_config.output_dir is None:
            QMessageBox.warning(self, "Configuration Incomplete", "Input or output folders are not configured.")
            return

        session_config.ensure_dirs()
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)

        dialog = WorkflowDialog("Run Scorer", "Starting scorer...", self)
        dialog.show()

        result_queue: queue.Queue = queue.Queue()

        def worker() -> None:
            try:
                scorer = LeagueScorer(session_config.input_dir, session_config.output_dir, session_config.year)
                warnings = scorer.run()
                result_queue.put(("done", warnings))
            except Exception as exc:
                result_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

        def poll() -> None:
            try:
                msg = result_queue.get_nowait()
            except queue.Empty:
                QTimer.singleShot(100, poll)
                return

            if msg[0] == "done":
                warnings = msg[1]
                dialog.set_finished(True)
                if warnings:
                    dialog.append_output("Warnings:\n" + "\n".join(warnings))
                    QMessageBox.information(self, "Run Complete", "Run complete with warnings. See the log for details.")
                else:
                    QMessageBox.information(self, "Run Complete", "Run completed successfully.")
            elif msg[0] == "error":
                dialog.append_output(str(msg[1]))
                dialog.set_finished(False)
                QMessageBox.critical(self, "Run Failed", f"Run failed:\n{msg[1]}")

        QTimer.singleShot(100, poll)

    def _run_workflow(
        self,
        *,
        script_name: str,
        title: str,
        initial_status: str,
        extra_cmd_args: list[str],
    ) -> None:
        if not self._require_configured(title):
            return
        if session_config.data_root is None:
            QMessageBox.critical(self, "Data Root Missing", "Set Data Root before running this workflow.")
            return
        if session_config.output_dir is None:
            QMessageBox.critical(self, "Output Missing", "Output directory is not configured.")
            return

        session_config.ensure_dirs()
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)

        repository_root = _find_repository_root()
        script_path = repository_root / "scripts" / script_name
        if not script_path.exists():
            QMessageBox.critical(
                self,
                "Script Missing",
                f"Workflow script not found: {script_path}\n\nExpected repository root: {repository_root}",
            )
            return

        dialog = WorkflowDialog(title, initial_status, self)
        dialog.show()
        dialog.append_output(f"Running: {script_path}")

        result_queue: queue.Queue = queue.Queue()

        def worker() -> None:
            try:
                output_paths = ensure_output_subdirs(session_config.output_dir)
                report_base = output_paths.autopilot_runs_dir
                cmd = [
                    sys.executable,
                    "-u",
                    str(script_path),
                    "--year",
                    str(session_config.year),
                    "--data-root",
                    str(session_config.data_root),
                    "--report-dir",
                    str(report_base),
                    *extra_cmd_args,
                ]
                repo_root = _find_repository_root()
                env = {k: v for k, v in os.environ.items() if k != "PYTHONSTARTUP"}
                env["PYTHONPATH"] = str(repo_root)
                proc = subprocess.Popen(
                    cmd,
                    cwd=str(repo_root),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    env=env,
                )
                assert proc.stdout is not None

                stderr_lines: list[str] = []

                def drain_stderr() -> None:
                    if proc.stderr is None:
                        return
                    for raw in proc.stderr:
                        stderr_lines.append(raw.rstrip("\n"))

                threading.Thread(target=drain_stderr, daemon=True).start()

                for raw in proc.stdout:
                    line = raw.rstrip("\n")
                    result_queue.put(("line", line))

                proc.wait()
                stderr_text = "\n".join(stderr_lines)
                result_queue.put(("done", proc.returncode, stderr_text))
            except Exception as exc:
                result_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

        def handle_line(line: str) -> None:
            if line.startswith("PROGRESS:"):
                readable = line.replace("PROGRESS:", "", 1).strip()
                dialog.set_status(readable)
            dialog.append_output(line)

        def poll() -> None:
            try:
                item = result_queue.get_nowait()
            except queue.Empty:
                QTimer.singleShot(100, poll)
                return

            if item[0] == "line":
                handle_line(item[1])
                QTimer.singleShot(20, poll)
                return
            if item[0] == "done":
                code, stderr_text = item[1], item[2]
                success = code == 0
                dialog.set_finished(success)
                if success:
                    QMessageBox.information(self, f"{title} Complete", f"{title} completed successfully.")
                else:
                    dialog.append_output(stderr_text)
                    QMessageBox.critical(self, f"{title} Failed", f"{title} failed. See output for details.")
                return
            if item[0] == "error":
                dialog.append_output(item[1])
                dialog.set_finished(False)
                QMessageBox.critical(self, f"{title} Error", item[1])
                return

        QTimer.singleShot(100, poll)

    def _prompt_manual_sporthive_pages(self, race_url: str) -> list[str] | None:
        pages: list[str] = []
        page_number = 1
        while True:
            dialog = ManualPageDialog(page_number, self)
            if dialog.exec() != QDialog.Accepted:
                return None
            page_text = dialog.page_text()
            if not page_text:
                QMessageBox.warning(self, "No Content", "No rows detected. Paste the page content and try again.")
                continue
            pages.append(page_text)
            more = QMessageBox.question(
                self,
                "Add Another Page?",
                "Add another Sporthive results page? Choose No when finished.",
                QMessageBox.Yes | QMessageBox.No,
            )
            if more == QMessageBox.No:
                break
            page_number += 1
        return pages

    @Slot()
    def _on_help(self) -> None:
        QMessageBox.information(
            self,
            "Help",
            "Use the dashboard actions to run workflows, publish results, and open reports.",
        )

    @Slot()
    def _on_run_scorer(self) -> None:
        self._run_scorer()

    @Slot()
    def _on_run_autopilot(self) -> None:
        self._run_workflow(
            script_name="autopilot/run_full_autopilot.py",
            title="Autopilot",
            initial_status="Initialising autopilot...",
            extra_cmd_args=[
                "--mode",
                "apply-safe-fixes",
                "--staged-report-dir",
                str(ensure_output_subdirs(session_config.output_dir).quality_staged_checks_dir),
                "--data-quality-output-dir",
                str(ensure_output_subdirs(session_config.output_dir).quality_data_dir),
            ],
        )

    @Slot()
    def _on_publish_results(self) -> None:
        self._run_workflow(
            script_name="run_publish_results.py",
            title="Publish Results",
            initial_status="Initialising publish...",
            extra_cmd_args=[],
        )

    @Slot()
    def _on_export_published_pdfs(self) -> None:
        if not self._require_configured("Export Published PDFs"):
            return
        if session_config.output_dir is None:
            QMessageBox.critical(self, "Output Missing", "Output directory is not configured.")
            return

        export_dir = QFileDialog.getExistingDirectory(self, "Select folder to export published PDFs", str(Path.home()))
        if not export_dir:
            return

        try:
            from league_scorer.output.output_layout import export_publish_pdfs

            export_path = export_publish_pdfs(session_config.output_dir, Path(export_dir), flatten=True)
            count = sum(1 for _ in export_path.glob("*.pdf"))
            QMessageBox.information(self, "Export Complete", f"Exported {count} published PDF(s) to {export_path}")
            self._open_path_in_system(export_path)
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", str(exc))

    @Slot()
    def _on_import_raceroster(self) -> None:
        if not self._require_configured("Import Race Roster"):
            return
        dialog = RaceRosterDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return
        race_url, race_number, race_name, race_hint = dialog.get_values()
        if not race_url:
            QMessageBox.warning(self, "Missing URL", "Race URL is required.")
            return

        if session_config.raw_data_dir is None:
            QMessageBox.critical(self, "Input Missing", "Raw data folder is not configured.")
            return
        session_config.ensure_dirs()

        progress = WorkflowDialog("Import Race Roster", "Importing race roster results...", self)
        progress.show()

        result_queue: queue.Queue = queue.Queue()

        def worker() -> None:
            try:
                output_path, count, history_path = import_raceroster_results(
                    race_url=race_url,
                    input_dir=session_config.raw_data_dir,
                    league_race_number=race_number,
                    race_name_override=race_name,
                    sporthive_race_id_hint=race_hint,
                )
            except SporthiveRaceNotDirectlyImportableError as exc:
                result_queue.put(("manual", str(exc)))
            except Exception as exc:
                result_queue.put(("error", str(exc)))
            else:
                result_queue.put(("ok", (output_path, count, history_path)))

        threading.Thread(target=worker, daemon=True).start()

        def poll() -> None:
            try:
                msg = result_queue.get_nowait()
            except queue.Empty:
                QTimer.singleShot(100, poll)
                return

            if msg[0] == "ok":
                output_path, count, history_path = msg[1]
                progress.append_output(f"Imported {count} rows to {output_path}")
                progress.set_finished(True)
                QMessageBox.information(
                    self,
                    "Import Complete",
                    f"Imported {count} rows to:\n{output_path}\n\nImport history:\n{history_path}",
                )
            elif msg[0] == "manual":
                progress.append_output(msg[1])
                progress.set_finished(False)
                if QMessageBox.question(
                    self,
                    "Manual Sporthive Import",
                    "Direct Sporthive import failed. Use manual page paste mode?",
                    QMessageBox.Yes | QMessageBox.No,
                ) == QMessageBox.Yes:
                    pages = self._prompt_manual_sporthive_pages(race_url)
                    if pages is None:
                        QMessageBox.information(self, "Cancelled", "Manual import cancelled.")
                        return
                    manual_progress = WorkflowDialog("Manual Sporthive Import", "Importing manual Sporthive pages...", self)
                    manual_progress.show()
                    manual_queue: queue.Queue = queue.Queue()

                    def manual_worker() -> None:
                        try:
                            output_path, count, history_path = import_sporthive_manual_pages(
                                race_url=race_url,
                                pages_text=pages,
                                input_dir=session_config.raw_data_dir,
                                league_race_number=race_number,
                                race_name_override=race_name,
                            )
                        except Exception as exc:
                            manual_queue.put(("error", str(exc)))
                        else:
                            manual_queue.put(("ok", (output_path, count, history_path)))

                    threading.Thread(target=manual_worker, daemon=True).start()

                    def manual_poll() -> None:
                        try:
                            item = manual_queue.get_nowait()
                        except queue.Empty:
                            QTimer.singleShot(100, manual_poll)
                            return
                        if item[0] == "ok":
                            output_path, count, history_path = item[1]
                            manual_progress.append_output(f"Imported {count} rows to {output_path}")
                            manual_progress.set_finished(True)
                            QMessageBox.information(
                                self,
                                "Import Complete",
                                f"Imported {count} rows to:\n{output_path}\n\nImport history:\n{history_path}",
                            )
                        else:
                            manual_progress.append_output(item[1])
                            manual_progress.set_finished(False)
                            QMessageBox.critical(self, "Import Failed", item[1])

                    QTimer.singleShot(100, manual_poll)
            elif msg[0] == "error":
                progress.append_output(msg[1])
                progress.set_finished(False)
                QMessageBox.critical(self, "Import Failed", msg[1])

        QTimer.singleShot(100, poll)

    @Slot()
    def _on_view_autopilot_report(self) -> None:
        if not self._require_configured("View Autopilot Report"):
            return
        if session_config.output_dir is None:
            QMessageBox.warning(self, "No Output", "Output directory is not configured.")
            return

        sort_existing_output_files(session_config.output_dir)
        viewer = AutopilotReportsWindow(session_config.output_dir, session_config.year)
        viewer.show()
        self._autopilot_viewer = viewer

    @Slot()
    def _on_view_results(self) -> None:
        if not self._require_configured("View Results"):
            return
        if session_config.output_dir is None:
            QMessageBox.warning(self, "No Output", "Output directory is not configured.")
            return

        viewer = ResultsViewerWindow(session_config.output_dir)
        viewer.show()
        self._results_viewer = viewer

    @Slot()
    def _find_events_path(self) -> Path | None:
        events_path = session_config.events_path
        if events_path and events_path.exists():
            return events_path

        control_dir = session_config.control_dir
        if control_dir and control_dir.exists():
            # Prefer the canonical events spreadsheet name, case-insensitively.
            for child in sorted(control_dir.iterdir(), key=lambda p: p.name.lower()):
                if child.name.lower() == "wrrl_events.xlsx" and child.exists():
                    session_config.events_path = child
                    return child

            candidates = [
                child
                for child in control_dir.iterdir()
                if child.suffix.lower() == ".xlsx" and "event" in child.name.lower()
            ]
            if candidates:
                candidates.sort(key=lambda p: p.name.lower())
                session_config.events_path = candidates[0]
                return candidates[0]

        selected = self._prompt_select_events_file()
        if selected:
            session_config.events_path = selected
            return selected

        return None

    def _prompt_select_events_file(self) -> Path | None:
        control_dir = session_config.control_dir
        initial_dir = str(control_dir) if control_dir and control_dir.exists() else str(Path.home())
        events_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Events Spreadsheet",
            initial_dir,
            "Excel Files (*.xlsx *.xls);;All Files (*)",
        )
        if not events_path:
            return None
        return Path(events_path)

    def _load_events_schedule(self):
        events_path = self._find_events_path()
        if not events_path:
            return None

        try:
            return load_events(events_path)
        except Exception as exc:
            message = (
                f"Could not load events spreadsheet:\n{events_path}\n\n{exc}\n\n"
                "Would you like to choose a different events file?"
            )
            choice = QMessageBox.question(
                self,
                "Events Load Failed",
                message,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if choice == QMessageBox.Yes:
                selected = self._prompt_select_events_file()
                if selected:
                    session_config.events_path = selected
                    try:
                        return load_events(selected)
                    except Exception as exc2:
                        QMessageBox.warning(
                            self,
                            "Events Load Failed",
                            f"Could not load events spreadsheet:\n{selected}\n\n{exc2}",
                        )
            return None

    @Slot()
    def _on_view_events(self) -> None:
        if not self._require_configured("View Events"):
            return

        schedule = self._load_events_schedule()
        if schedule is None:
            QMessageBox.warning(
                self,
                "Events Not Found",
                "No events spreadsheet could be found in inputs/control."
                "\nExpected default: wrrl_events.xlsx",
            )
            return

        images_dir = Path(__file__).parent.parent / "images"
        viewer = EventsViewerWindow(
            schedule,
            year=session_config.year,
            images_dir=images_dir,
            output_dir=session_config.output_dir,
        )
        viewer.show()
        self._events_viewer = viewer

    @Slot()
    def _on_review_raes(self) -> None:
        if not self._require_configured("Data Corrections (RAES)"):
            return

        if self._raes_window is None:
            self._raes_window = RAESWindow(self)
        self._raes_window.show()
        self._raes_window.raise_()
        self._raes_window.activateWindow()

    @Slot()
    def _on_compare_raw_archive(self) -> None:
        if not self._require_configured("Compare Raw vs Archive"):
            return

        viewer = RawArchiveDiffWindow()
        viewer.show()
        self._raw_archive_diff_viewer = viewer

    @Slot()
    def _on_view_runner_history(self) -> None:
        if not self._require_configured("Runner/Club Enquiry"):
            return

        viewer = RunnerHistoryWindow(session_config.output_dir)
        viewer.show()
        self._runner_history_viewer = viewer

    @Slot()
    def _on_settings(self) -> None:
        dialog = SettingsDialog(self)
        dialog.exec()
        self._refresh_config_panel()

    @Slot()
    def _on_run_provisional_fast_track(self) -> None:
        self._run_workflow(
            script_name="publish/run_provisional_fast_track.py",
            title="Provisional Fast Track",
            initial_status="Initialising provisional fast track...",
            extra_cmd_args=[],
        )

    def _on_year_prev(self) -> None:
        years = session_config.available_years()
        idx = years.index(session_config.year) if session_config.year in years else 0
        if idx > 0:
            session_config.year = years[idx - 1]
            self.year_label.setText(str(session_config.year))

    def _on_year_next(self) -> None:
        years = session_config.available_years()
        idx = years.index(session_config.year) if session_config.year in years else 0
        if idx < len(years) - 1:
            session_config.year = years[idx + 1]
            self.year_label.setText(str(session_config.year))

    def _show_unavailable(self, feature: str) -> None:
        QMessageBox.information(
            self,
            f"{feature} not ready",
            f"The Qt UI is under development. {feature} will be available in a later update.",
        )


def launch_dashboard() -> None:
    app = QApplication(sys.argv)
    app.setWindowIcon(
        QIcon(
            str(
                _find_repository_root()
                / "league_scorer"
                / "images"
                / "WRRL shield concept.png"
            )
        )
    )
    window = QtLeagueScorerDashboard()
    window.show()
    app.exec()
