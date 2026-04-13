"""Qt window for viewing the WRRL Championship Events schedule."""

from __future__ import annotations

from pathlib import Path

from PIL.ImageQt import ImageQt
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QFont, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from league_scorer.events_loader import (
    EventEntry,
    EventsSchedule,
    STATUS_CONFIRMED,
    STATUS_PROVISIONAL,
    STATUS_TBC,
)
from league_scorer.graphical.timeline_generator import generate_timeline

WRRL_NAVY = "#3a4658"
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"

_STATUS_BG = {
    STATUS_CONFIRMED.lower(): "#e8f5e9",
    STATUS_PROVISIONAL.lower(): "#e3f2fd",
    STATUS_TBC.lower(): "#fff8e1",
}
_STATUS_FG = {
    STATUS_CONFIRMED.lower(): "#1b5e20",
    STATUS_PROVISIONAL.lower(): "#0d47a1",
    STATUS_TBC.lower(): "#e65100",
}

_COLUMNS = [
    ("race_ref", "Ref", 70),
    ("event_name", "Event", 220),
    ("distance", "Distance", 70),
    ("location", "Location", 120),
    ("organiser", "Organiser", 130),
    ("date_type", "Date Type", 90),
    ("scheduled_dates", "Dates", 180),
    ("entry_fee", "Fee", 70),
    ("scoring_basis", "Scoring", 90),
    ("status", "Status", 90),
    ("website", "Website", 180),
]


class EventsViewerWindow(QMainWindow):
    def __init__(
        self,
        schedule: EventsSchedule,
        year: int = 2026,
        images_dir: Path | None = None,
        output_dir: Path | None = None,
    ) -> None:
        super().__init__()
        self._schedule = schedule
        self._year = year
        self._images_dir = images_dir
        self._output_dir = output_dir
        self.setWindowTitle("Championship Events Schedule")
        self.resize(1200, 760)
        self._build_ui()
        self._populate(self._schedule.events)

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setObjectName("eventsViewer")
        central.setStyleSheet(f"background: {WRRL_LIGHT};")
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        layout.addWidget(self._build_toolbar())
        layout.addWidget(self._build_summary_bar())
        layout.addWidget(self._build_table())
        layout.addWidget(self._build_status_bar())

    def _build_toolbar(self) -> QWidget:
        toolbar = QWidget(self)
        toolbar.setStyleSheet(f"background: {WRRL_NAVY};")
        toolbar.setFixedHeight(52)

        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(16, 8, 16, 8)
        layout.setSpacing(10)

        title = QLabel("Championship Events Schedule", toolbar)
        title.setStyleSheet(f"color: {WRRL_WHITE};")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        layout.addWidget(title)

        layout.addStretch(1)

        if self._schedule.source_path:
            source_label = QLabel(f"Source: {self._schedule.source_path.name}", toolbar)
            source_label.setStyleSheet("color: #a0b0c0;")
            source_label.setFont(QFont("Segoe UI", 9))
            layout.addWidget(source_label)

        self._timeline_button = QPushButton("📅 Generate Timeline", toolbar)
        self._timeline_button.setCursor(Qt.PointingHandCursor)
        self._timeline_button.setStyleSheet(
            "QPushButton { background: #2d7a4a; color: #ffffff; border: none; border-radius: 8px; padding: 8px 14px; }"
            "QPushButton:hover { background: #24653d; }"
        )
        self._timeline_button.clicked.connect(self._on_generate_timeline)
        layout.addWidget(self._timeline_button)

        self._open_website_button = QPushButton("🌐 Open Website", toolbar)
        self._open_website_button.setCursor(Qt.PointingHandCursor)
        self._open_website_button.setEnabled(False)
        self._open_website_button.setStyleSheet(
            "QPushButton { background: #ffffff; color: #3a4658; border: 1px solid #ccd7e3; border-radius: 8px; padding: 8px 14px; }"
            "QPushButton:hover { background: #eef2f7; }"
        )
        self._open_website_button.clicked.connect(self._on_open_website)
        layout.addWidget(self._open_website_button)

        self._dashboard_button = QPushButton("🏠 Close", toolbar)
        self._dashboard_button.setCursor(Qt.PointingHandCursor)
        self._dashboard_button.setStyleSheet(
            "QPushButton { background: #f5f5f5; color: #3a4658; border: none; border-radius: 8px; padding: 8px 14px; }"
            "QPushButton:hover { background: #e0e6ef; }"
        )
        self._dashboard_button.clicked.connect(self._on_close)
        layout.addWidget(self._dashboard_button)

        return toolbar

    def _build_summary_bar(self) -> QWidget:
        schedule = self._schedule
        total = len(schedule.events)
        confirmed = len(schedule.confirmed)
        provisional = len(schedule.provisional)
        tbc = len(schedule.tbc)

        summary = QWidget(self)
        summary.setStyleSheet("background: #e8eaf0;")
        summary.setFixedHeight(36)

        layout = QHBoxLayout(summary)
        layout.setContentsMargins(16, 0, 16, 0)

        summary_text = (
            f"Total: {total}    "
            f"Confirmed: {confirmed}    "
            f"Provisional: {provisional}    "
            f"TBC: {tbc}"
        )
        label = QLabel(summary_text, summary)
        label.setFont(QFont("Segoe UI", 9))
        label.setStyleSheet(f"color: {WRRL_NAVY};")
        layout.addWidget(label)
        layout.addStretch(1)

        return summary

    def _build_table(self) -> QWidget:
        container = QWidget(self)
        table_layout = QVBoxLayout(container)
        table_layout.setContentsMargins(10, 10, 10, 10)

        self._table = QTableWidget(container)
        self._table.setColumnCount(len(_COLUMNS))
        self._table.setHorizontalHeaderLabels([col[1] for col in _COLUMNS])
        self._table.setSortingEnabled(True)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setDefaultAlignment(Qt.AlignLeft)

        for index, (_, _, width) in enumerate(_COLUMNS):
            self._table.setColumnWidth(index, width)

        self._table.horizontalHeader().sectionClicked.connect(self._on_sort)
        self._table.selectionModel().selectionChanged.connect(self._on_table_selection_changed)
        self._table.cellDoubleClicked.connect(self._on_table_cell_double_clicked)
        table_layout.addWidget(self._table)
        return container

    def _build_status_bar(self) -> QWidget:
        bar = QWidget(self)
        bar.setStyleSheet("background: #dde0e8;")
        bar.setFixedHeight(28)

        layout = QHBoxLayout(bar)
        layout.setContentsMargins(12, 0, 12, 0)

        self._status_label = QLabel("Click a column header to sort.", bar)
        self._status_label.setFont(QFont("Segoe UI", 8))
        self._status_label.setStyleSheet("color: #555566;")
        layout.addWidget(self._status_label)
        layout.addStretch(1)

        return bar

    def _populate(self, events: list[EventEntry]) -> None:
        self._table.setRowCount(len(events))
        for row_index, ev in enumerate(events):
            values = [
                ev.race_ref,
                ev.event_name,
                ev.distance,
                ev.location,
                ev.organiser,
                ev.date_type,
                ev.scheduled_dates,
                ev.entry_fee,
                ev.scoring_basis,
                ev.status,
                ev.website,
            ]
            status_key = ev.status.lower()
            bg = _STATUS_BG.get(status_key, "#ffffff")
            fg = _STATUS_FG.get(status_key, "#333333")
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                item.setBackground(QColor(bg))
                if col_index == len(values) - 1 and value:
                    item.setForeground(QColor("#1a0dab"))
                    font = item.font()
                    font.setUnderline(True)
                    item.setFont(font)
                    item.setToolTip("Double-click to open website")
                else:
                    item.setForeground(QColor(fg))
                item.setData(Qt.UserRole, row_index)
                self._table.setItem(row_index, col_index, item)

        self._table.resizeRowsToContents()

    def _apply_row_style(self, row_index: int, bg: str, fg: str) -> None:
        for col_index in range(self._table.columnCount()):
            item = self._table.item(row_index, col_index)
            if item is not None:
                item.setBackground(QColor(bg))
                item.setForeground(QColor(fg))

    def _on_sort(self, index: int) -> None:
        column_name = _COLUMNS[index][0]
        self._status_label.setText(f"Sorted by '{column_name}'.")

    def _selected_event_for_row(self, row: int) -> EventEntry | None:
        if row < 0 or row >= self._table.rowCount():
            return None
        item = self._table.item(row, 0)
        if item is None:
            return None
        event_index = item.data(Qt.UserRole)
        if event_index is None:
            return None
        return self._schedule.events[event_index]

    def _on_table_selection_changed(self, _selected, _deselected) -> None:
        event = self._selected_event_for_row(self._table.currentRow())
        has_website = bool(event and event.website)
        self._open_website_button.setEnabled(has_website)

        if not event:
            self._status_label.setText("Click a column header to sort.")
        elif has_website:
            self._status_label.setText(f"Website: {event.website} — double-click or press Open Website")
        else:
            self._status_label.setText("No website configured for the selected event.")

    def _on_table_cell_double_clicked(self, row: int, column: int) -> None:
        if column != len(_COLUMNS) - 1:
            return
        event = self._selected_event_for_row(row)
        if event is None or not event.website:
            return
        QDesktopServices.openUrl(QUrl(str(event.website)))

    def _on_open_website(self) -> None:
        event = self._selected_event_for_row(self._table.currentRow())
        if event is None or not event.website:
            QMessageBox.warning(self, "No Website", "The selected event has no website configured.", parent=self)
            return
        QDesktopServices.openUrl(QUrl(str(event.website)))

    def _on_generate_timeline(self) -> None:
        if not self._schedule.events:
            QMessageBox.warning(self, "No Events", "No events are loaded.", parent=self)
            return

        default_name = f"WRRL_season_timeline_{self._year}.png"
        default_dir = str(self._output_dir) if self._output_dir else str(Path.home())
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Timeline As",
            default_dir + '/' + default_name,
            "PNG Image (*.png)",
        )
        if not save_path:
            return

        try:
            img = generate_timeline(
                self._schedule,
                year=self._year,
                output_path=Path(save_path),
                images_dir=self._images_dir,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Timeline Error", f"Could not generate timeline:\n{exc}", parent=self)
            return

        self._show_timeline_preview(img, Path(save_path))

    def _show_timeline_preview(self, img, saved_path: Path) -> None:
        preview = QDialog(self)
        preview.setWindowTitle(f"Season Timeline - {saved_path.name}")
        preview.resize(1100, 800)

        layout = QVBoxLayout(preview)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        toolbar = QWidget(preview)
        toolbar.setStyleSheet(f"background: {WRRL_NAVY};")
        toolbar.setFixedHeight(42)
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 6, 12, 6)
        toolbar_layout.addWidget(QLabel(f"Saved: {saved_path}", toolbar))
        toolbar_layout.addStretch(1)
        layout.addWidget(toolbar)

        scroll = QScrollArea(preview)
        scroll.setWidgetResizable(True)
        content = QLabel()
        content.setAlignment(Qt.AlignCenter)

        qimage = ImageQt(img.convert("RGBA"))
        content.setPixmap(QPixmap.fromImage(qimage))
        content.setMinimumSize(qimage.width(), qimage.height())

        scroll.setWidget(content)
        layout.addWidget(scroll)

        preview.exec()

    def _on_close(self) -> None:
        self.close()
