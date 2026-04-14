"""Qt window for browsing generated autopilot reports and related outputs."""

import os
from pathlib import Path
import sys

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from league_scorer.output.output_layout import build_output_paths
from league_scorer.session_config import config as session_config


def _human_size(size: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.0f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"


class AutopilotReportsWindow(QMainWindow):
    def __init__(self, output_dir: Path | None, year: int) -> None:
        super().__init__()
        self._output_dir = output_dir
        self._year = year
        self._files: list[Path] = []
        self.setWindowTitle("Autopilot Reports")
        self.resize(1200, 760)
        self._build_ui()
        self._refresh_reports()

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setStyleSheet("background: #f5f5f5;")
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        top_bar = QWidget(central)
        top_bar.setStyleSheet("background: #ffffff; border-radius: 12px;")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(16, 16, 16, 16)
        top_layout.setSpacing(12)

        header = QLabel("Autopilot Reports", top_bar)
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

        self._refresh_btn = QPushButton("Refresh", button_bar)
        self._refresh_btn.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self._refresh_btn.clicked.connect(self._refresh_reports)
        button_layout.addWidget(self._refresh_btn)

        self._manual_btn = QPushButton("Open Manual Audit", button_bar)
        self._manual_btn.clicked.connect(self._open_manual_audit)
        button_layout.addWidget(self._manual_btn)

        self._season_btn = QPushButton("Open Season Audit", button_bar)
        self._season_btn.clicked.connect(self._open_season_audit)
        button_layout.addWidget(self._season_btn)

        self._folder_btn = QPushButton("Open Reports Folder", button_bar)
        self._folder_btn.clicked.connect(self._open_reports_folder)
        button_layout.addWidget(self._folder_btn)

        button_layout.addStretch(1)
        root_layout.addWidget(button_bar)

        content = QWidget(central)
        content_layout = QHBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(12)

        nav_panel = QWidget(content)
        nav_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        nav_layout = QVBoxLayout(nav_panel)
        nav_layout.setContentsMargins(16, 16, 16, 16)
        nav_layout.setSpacing(12)

        nav_title = QLabel("Reports", nav_panel)
        nav_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        nav_title.setStyleSheet("color: #2d7a4a;")
        nav_layout.addWidget(nav_title)

        self._reports_status_label = QLabel("", nav_panel)
        self._reports_status_label.setFont(QFont("Segoe UI", 9))
        self._reports_status_label.setStyleSheet("color: #4d5d6d;")
        nav_layout.addWidget(self._reports_status_label)

        self._table = QTableWidget(nav_panel)
        self._table.setColumnCount(2)
        self._table.setHorizontalHeaderLabels(["Name", "Modified"])
        self._table.verticalHeader().setVisible(False)
        self._table.setShowGrid(False)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.cellDoubleClicked.connect(self._open_selected_file)
        self._table.cellClicked.connect(self._preview_selected_file)
        nav_layout.addWidget(self._table, 1)

        content_layout.addWidget(nav_panel, 1)

        preview_panel = QWidget(content)
        preview_panel.setStyleSheet("background: #ffffff; border-radius: 12px;")
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(16, 16, 16, 16)
        preview_layout.setSpacing(12)

        preview_title = QLabel("Preview", preview_panel)
        preview_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        preview_title.setStyleSheet("color: #2d7a4a;")
        preview_layout.addWidget(preview_title)

        self._preview = QTextEdit(preview_panel)
        self._preview.setReadOnly(True)
        self._preview.setStyleSheet("background: #fdfdfd;")
        preview_layout.addWidget(self._preview, 1)

        content_layout.addWidget(preview_panel, 3)
        root_layout.addWidget(content, 1)

        message_bar = QWidget(central)
        message_bar.setStyleSheet("background: #1e1e1e; border-radius: 8px;")
        message_layout = QHBoxLayout(message_bar)
        message_layout.setContentsMargins(14, 10, 14, 10)

        self._message_label = QLabel("Ready", message_bar)
        self._message_label.setFont(QFont("Segoe UI", 10))
        self._message_label.setStyleSheet("color: #d4d4d4;")
        message_layout.addWidget(self._message_label)

        root_layout.addWidget(message_bar)

    def _reports_dirs(self) -> list[Path]:
        if self._output_dir is None:
            return []
        base = build_output_paths(self._output_dir)
        paths = [
            base.autopilot_runs_dir / f"year-{self._year}",
            base.quality_data_dir / f"year-{self._year}",
            base.quality_staged_checks_dir,
        ]
        return [p for p in paths if p.exists()]

    def _refresh_reports(self) -> None:
        self._files = []
        self._table.setRowCount(0)
        report_dirs = self._reports_dirs()
        if not report_dirs:
            self._preview.setPlainText("No autopilot reports found for the configured season.")
            self._reports_status_label.setText("No reports found")
            self._message_label.setText("No reports available for the configured season.")
            return

        file_rows: list[tuple[Path, str, str]] = []
        for folder in report_dirs:
            for path in sorted(folder.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True):
                file_rows.append((path, path.name, self._format_mtime(path.stat().st_mtime)))

        self._files = [row[0] for row in file_rows]
        self._table.setRowCount(len(self._files))
        self._reports_status_label.setText(f"{len(self._files)} reports found")
        self._message_label.setText(f"Loaded {len(self._files)} report(s). Double-click a row to open the selected report.")
        for row_index, (path, name, modified) in enumerate(file_rows):
            for col_index, value in enumerate((name, modified)):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                self._table.setItem(row_index, col_index, item)

        self._table.resizeColumnsToContents()
        if self._files:
            self._table.selectRow(0)
            self._preview_file(self._files[0])

    def _format_mtime(self, timestamp: float) -> str:
        from datetime import datetime

        return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")

    def _open_file_in_system(self, path: Path) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        except OSError as exc:
            QMessageBox.critical(self, "Open Failed", str(exc), parent=self)

    def _open_folder_in_system(self, path: Path) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        except OSError as exc:
            QMessageBox.critical(self, "Open Failed", str(exc), parent=self)

    def _open_reports_folder(self) -> None:
        dirs = self._reports_dirs()
        if not dirs:
            QMessageBox.warning(self, "No Reports", "No autopilot report folders were found.", parent=self)
            return
        self._open_folder_in_system(dirs[0])

    def _open_manual_audit(self) -> None:
        if self._output_dir is None:
            QMessageBox.warning(self, "Not Configured", "Output directory is not configured.", parent=self)
            return
        paths = build_output_paths(self._output_dir)
        target = paths.audit_manual_changes_dir / "manual_data_audit.xlsx"
        if not target.exists():
            QMessageBox.warning(self, "Not Found", "manual_data_audit.xlsx not found in output audit folder.", parent=self)
            return
        self._open_file_in_system(target)

    def _open_season_audit(self) -> None:
        if self._output_dir is None:
            QMessageBox.warning(self, "Not Configured", "Output directory is not configured.", parent=self)
            return
        paths = build_output_paths(self._output_dir)
        audit_dir = paths.audit_workbooks_dir
        if not audit_dir.exists():
            QMessageBox.warning(self, "Not Found", "No audit workbooks folder present.", parent=self)
            return
        xlsx_files = sorted(audit_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not xlsx_files:
            QMessageBox.warning(self, "No Files", "No audit workbook files found.", parent=self)
            return
        self._open_file_in_system(xlsx_files[0])

    def _open_selected_file(self, row: int, column: int) -> None:
        if row < 0 or row >= len(self._files):
            return
        self._open_file_in_system(self._files[row])

    def _preview_selected_file(self, row: int, column: int) -> None:
        if row < 0 or row >= len(self._files):
            return
        self._preview_file(self._files[row])

    def _preview_file(self, path: Path) -> None:
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except Exception as exc:
            self._preview.setPlainText(f"Failed to preview {path.name}: {exc}")
            return
        self._preview.setPlainText(text)
