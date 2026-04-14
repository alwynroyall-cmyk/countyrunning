"""Qt window for comparing raw data with archived raw data files."""

from __future__ import annotations

import html
import os
import threading
from pathlib import Path

from PySide6.QtCore import QTimer, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStyle,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from league_scorer.raw_archive_diff_service import (
    DiffRow,
    build_side_by_side_diff,
    list_comparable_file_pairs,
    load_comparable_lines,
)
from league_scorer.session_config import config as session_config

WRRL_NAVY = "#3a4658"
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"

_STATUS_COLOURS = {
    "same": "#ffffff",
    "replace": "#fff0c9",
    "delete": "#fde4e4",
    "insert": "#e4f6e8",
}


class RawArchiveDiffWindow(QMainWindow):
    status_update = Signal(str)
    message_update = Signal(str)
    diff_done = Signal(object, int, int)

    def __init__(self) -> None:
        super().__init__()
        self._pairs_by_name: dict[str, object] = {}
        self._pair_names: list[str] = []
        self._syncing = False
        self.setWindowTitle("Raw Data vs Archive Diff")
        self.resize(1280, 820)
        self.status_update.connect(self._on_status_update)
        self.message_update.connect(self._render_message)
        self.diff_done.connect(self._on_diff_done)
        self._build_ui()
        self._load_file_pairs()

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

        title = QLabel("Raw Data vs Archive Diff", header_panel)
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title.setStyleSheet(f"color: {WRRL_GREEN};")
        header_layout.addWidget(title)
        header_layout.addStretch(1)

        root_layout.addWidget(header_panel)

        controls_panel = QWidget(central)
        controls_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        controls_layout = QHBoxLayout(controls_panel)
        controls_layout.setContentsMargins(16, 12, 16, 12)
        controls_layout.setSpacing(8)

        button_style = (
            "QPushButton { background: #ffffff; color: #3a4658; border: 1px solid #ccd7e3; border-radius: 8px; padding: 8px 14px; }"
            "QPushButton:hover { background: #eef2f7; }"
        )

        refresh_btn = QPushButton("Refresh", controls_panel)
        refresh_btn.setCursor(Qt.PointingHandCursor)
        refresh_btn.setStyleSheet(button_style)
        refresh_btn.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        refresh_btn.clicked.connect(self._load_file_pairs)
        controls_layout.addWidget(refresh_btn)

        open_raw_btn = QPushButton("Open Raw", controls_panel)
        open_raw_btn.setCursor(Qt.PointingHandCursor)
        open_raw_btn.setStyleSheet(button_style)
        open_raw_btn.clicked.connect(lambda: self._open_selected_file("raw"))
        controls_layout.addWidget(open_raw_btn)

        open_archive_btn = QPushButton("Open Archive", controls_panel)
        open_archive_btn.setCursor(Qt.PointingHandCursor)
        open_archive_btn.setStyleSheet(button_style)
        open_archive_btn.clicked.connect(lambda: self._open_selected_file("archive"))
        controls_layout.addWidget(open_archive_btn)

        close_btn = QPushButton("🏠 Close", controls_panel)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(button_style)
        close_btn.clicked.connect(self.close)
        controls_layout.addWidget(close_btn)

        controls_layout.addStretch(1)

        file_label = QLabel("File:", controls_panel)
        file_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        controls_layout.addWidget(file_label)

        self._file_combo = QComboBox(controls_panel)
        self._file_combo.setMinimumWidth(520)
        self._file_combo.currentIndexChanged.connect(self._load_selected_diff)
        controls_layout.addWidget(self._file_combo)

        root_layout.addWidget(controls_panel)

        content_panel = QWidget(central)
        content_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        content_layout = QVBoxLayout(content_panel)
        content_layout.setContentsMargins(16, 16, 16, 16)
        content_layout.setSpacing(12)

        splitter = QSplitter(Qt.Horizontal, content_panel)
        splitter.setHandleWidth(6)

        self._left_text = QTextEdit(splitter)
        self._left_text.setReadOnly(True)
        self._left_text.setFont(QFont("Consolas", 10))
        self._left_text.setLineWrapMode(QTextEdit.NoWrap)
        self._left_text.verticalScrollBar().valueChanged.connect(self._sync_vertical_scroll)

        self._right_text = QTextEdit(splitter)
        self._right_text.setReadOnly(True)
        self._right_text.setFont(QFont("Consolas", 10))
        self._right_text.setLineWrapMode(QTextEdit.NoWrap)
        self._right_text.verticalScrollBar().valueChanged.connect(self._sync_vertical_scroll)

        splitter.addWidget(self._left_text)
        splitter.addWidget(self._right_text)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        content_layout.addWidget(splitter, 1)
        root_layout.addWidget(content_panel, 1)

        root_layout.addWidget(self._build_footer_panel())

    def _build_footer_panel(self) -> QWidget:
        footer = QWidget(self)
        footer.setStyleSheet("background: #1e1e1e; border-radius: 12px;")
        footer.setFixedHeight(40)

        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(16, 8, 16, 8)

        self._status_label = QLabel("Select a file to compare.", footer)
        self._status_label.setFont(QFont("Segoe UI", 9))
        self._status_label.setStyleSheet("color: #d4d4d4;")
        footer_layout.addWidget(self._status_label)
        footer_layout.addStretch(1)

        return footer

    def _results_paths(self) -> tuple[Path | None, Path | None]:
        if session_config.input_dir is None:
            return None, None
        paths = session_config.input_dir
        raw_dir = paths / "raw_data"
        archive_dir = paths / "raw_data_archive"
        return raw_dir, archive_dir

    def _load_file_pairs(self) -> None:
        input_dir = session_config.input_dir
        if input_dir is None:
            self._pairs_by_name = {}
            self._pair_names = []
            self._file_combo.clear()
            self._status_label.setText("Inputs are not configured.")
            self._render_message("Set the season data root before comparing raw and archive files.")
            return

        pairs = list_comparable_file_pairs(input_dir)
        self._pairs_by_name = {pair.filename: pair for pair in pairs}
        self._pair_names = [pair.filename for pair in pairs]
        self._file_combo.clear()
        self._file_combo.addItems(self._pair_names)

        if not self._pair_names:
            self._status_label.setText("No matching raw/archive file pairs found.")
            self._render_message("No files with the same name exist in both raw_data and raw_data_archive.")
            return

        self._status_label.setText("Select a file to compare.")
        self._load_selected_diff()

    def _selected_pair(self):
        key = self._file_combo.currentText().strip()
        return self._pairs_by_name.get(key)

    def _load_selected_diff(self) -> None:
        pair = self._selected_pair()
        if pair is None:
            self._render_message("Choose a file to compare.")
            return

        self._status_label.setText(f"Loading {pair.filename}…")
        self._render_message("Loading…")

        def worker() -> None:
            try:
                raw_lines = load_comparable_lines(pair.raw_path)
                archive_lines = load_comparable_lines(pair.archive_path)
                diff_rows = build_side_by_side_diff(raw_lines, archive_lines)
            except Exception as exc:
                self.status_update.emit(f"Failed to compare {pair.filename}")
                self.message_update.emit(f"Could not compare the selected file: {exc}")
                return

            self.diff_done.emit(diff_rows, len(raw_lines), len(archive_lines))

        threading.Thread(target=worker, daemon=True).start()

    def _on_status_update(self, message: str) -> None:
        self._status_label.setText(message)

    def _on_diff_done(self, diff_rows: list[DiffRow], raw_line_count: int, archive_line_count: int) -> None:
        differing_rows = sum(1 for row in diff_rows if row.status != "same")
        summary = (
            f"{self._file_combo.currentText().strip()}: {differing_rows} differing line(s), raw lines {raw_line_count}, archive lines {archive_line_count}"
        )
        self._status_label.setText(summary)
        self._render_diff_rows(diff_rows)

    def _render_message(self, message: str) -> None:
        html_text = self._build_html(message, "same")
        self._left_text.setHtml(html_text)
        self._right_text.setHtml(html_text)

    def _render_diff_rows(self, rows: list[DiffRow]) -> None:
        left_html = []
        right_html = []

        for row in rows:
            left_html.append(self._line_html(row.left_line_no, row.left_text, row.status))
            right_html.append(self._line_html(row.right_line_no, row.right_text, row.status))

        self._left_text.setHtml(self._wrap_html(left_html))
        self._right_text.setHtml(self._wrap_html(right_html))

    def _line_html(self, line_no: int | None, text: str, status: str) -> str:
        line_number = f"{line_no:>5}" if line_no is not None else "     "
        escaped = html.escape(text)
        return (
            f"<div style=\"background: {_STATUS_COLOURS.get(status, '#ffffff')}; padding: 2px 6px;\">"
            f"<span style=\"color: #4d5d6d;\">{line_number} | </span>"
            f"<span>{escaped}</span>"
            "</div>"
        )

    def _wrap_html(self, lines: list[str]) -> str:
        return (
            '<html><body style="margin:0; font-family: Consolas, monospace; font-size: 10pt; white-space: pre;">'
            + "".join(lines)
            + "</body></html>"
        )

    def _build_html(self, text: str, status: str) -> str:
        return self._wrap_html([self._line_html(None, text, status)])

    def _sync_vertical_scroll(self, value: int) -> None:
        if self._syncing:
            return
        self._syncing = True
        sender = self.sender()
        target = self._right_text.verticalScrollBar() if sender is self._left_text.verticalScrollBar() else self._left_text.verticalScrollBar()
        target.setValue(value)
        self._syncing = False

    def _open_selected_file(self, kind: str) -> None:
        pair = self._selected_pair()
        if pair is None:
            QMessageBox.warning(self, "No File Selected", "Choose a file pair first.", parent=self)
            return

        path = pair.raw_path if kind == "raw" else pair.archive_path
        if os.name == "nt":
            os.startfile(str(path))
            return

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
