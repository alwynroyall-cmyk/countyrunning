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
        self.setWindowTitle("View Autopilot Reports")
        self.resize(1200, 760)
        self._build_ui()
        self._refresh_reports()

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setStyleSheet("background: #f5f5f5;")
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header = QLabel("Autopilot Reports", central)
        header.setFont(QFont("Segoe UI", 16, QFont.Bold))
        header.setStyleSheet("color: #2d7a4a;")
        layout.addWidget(header)

        button_row = QWidget(central)
        button_layout = QHBoxLayout(button_row)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(8)

        self._manual_btn = QPushButton("Open Manual Audit", button_row)
        self._manual_btn.clicked.connect(self._open_manual_audit)
        button_layout.addWidget(self._manual_btn)

        self._season_btn = QPushButton("Open Season Audit", button_row)
        self._season_btn.clicked.connect(self._open_season_audit)
        button_layout.addWidget(self._season_btn)

        self._folder_btn = QPushButton("Open Reports Folder", button_row)
        self._folder_btn.clicked.connect(self._open_reports_folder)
        button_layout.addWidget(self._folder_btn)

        self._refresh_btn = QPushButton("Refresh", button_row)
        self._refresh_btn.clicked.connect(self._refresh_reports)
        button_layout.addWidget(self._refresh_btn)

        button_layout.addStretch(1)
        layout.addWidget(button_row)

        split = QSplitter(Qt.Horizontal, central)
        split.setHandleWidth(6)

        self._table = QTableWidget(split)
        self._table.setColumnCount(4)
        self._table.setHorizontalHeaderLabels(["Name", "Modified", "Size", "Folder"])
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.cellDoubleClicked.connect(self._open_selected_file)
        self._table.cellClicked.connect(self._preview_selected_file)
        split.addWidget(self._table)

        self._preview = QTextEdit(split)
        self._preview.setReadOnly(True)
        self._preview.setStyleSheet("background: #ffffff;")
        split.addWidget(self._preview)

        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)

        layout.addWidget(split, 1)

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
            return

        file_rows: list[tuple[Path, str, str, str]] = []
        for folder in report_dirs:
            for path in sorted(folder.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True):
                file_rows.append((path, path.name, self._format_mtime(path.stat().st_mtime), _human_size(path.stat().st_size), str(folder.name)))

        self._files = [row[0] for row in file_rows]
        self._table.setRowCount(len(self._files))
        for row_index, (path, name, modified, size, folder_name) in enumerate(file_rows):
            for col_index, value in enumerate((name, modified, size, folder_name)):
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
