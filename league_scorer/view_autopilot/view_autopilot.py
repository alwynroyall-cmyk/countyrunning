import os
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional

from ..output_layout import build_output_paths
from ..session_config import config as session_config


class ViewAutopilotPanel(tk.Frame):
    """Panel that lists autopilot run markdown reports and allows opening/previewing."""

    def __init__(self, parent: tk.Misc):
        super().__init__(parent, bg="#f7f9fb")
        self._reports_dir: Optional[Path] = None
        self._files: list[Path] = []
        self._status_var = tk.StringVar()
        self._configure_ui()
        self._refresh_list()

    def _configure_ui(self) -> None:
        title = tk.Label(self, text="Autopilot Reports", font=("Segoe UI", 14, "bold"), bg="#f7f9fb")
        title.pack(pady=(8, 6))

        top_frame = tk.Frame(self, bg="#f7f9fb")
        top_frame.pack(fill="x", padx=10)

        tk.Button(top_frame, text="Refresh", command=self._refresh_list, bg="#e9f0f7").pack(side="left")
        tk.Button(top_frame, text="Open Folder", command=self._open_folder, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Button(top_frame, text="Open Manual Audit", command=self._open_manual_audit, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Button(top_frame, text="Open Season Audit", command=self._open_season_audit, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Button(top_frame, text="Open Data Quality Report", command=self._open_data_quality_report, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Label(top_frame, textvariable=self._status_var, bg="#f7f9fb", fg="#55666f").pack(side="right")

        middle = tk.PanedWindow(self, orient="horizontal")
        middle.pack(fill="both", expand=True, padx=10, pady=10)

        # File list
        list_frame = tk.Frame(middle, bg="#f7f9fb")
        self._tree = ttk.Treeview(list_frame, columns=("modified", "size"), show="headings", selectmode="browse")
        self._tree.heading("modified", text="Modified")
        self._tree.heading("size", text="Size")
        self._tree.column("modified", width=140, anchor="w")
        self._tree.column("size", width=80, anchor="e")
        self._tree.pack(fill="both", expand=True, side="left")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self._tree.yview)
        scrollbar.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=scrollbar.set)
        self._tree.bind("<Double-1>", lambda e: self._open_selected())

        middle.add(list_frame, minsize=200)

        # Preview
        preview_frame = tk.Frame(middle, bg="#ffffff")
        self._preview = tk.Text(preview_frame, wrap="word", state="disabled", bg="#ffffff", relief="flat")
        self._preview.pack(fill="both", expand=True)
        middle.add(preview_frame, minsize=320)

        btn_frame = tk.Frame(self, bg="#f7f9fb")
        btn_frame.pack(fill="x", padx=10, pady=(0, 8))
        tk.Button(btn_frame, text="Open Selected", command=self._open_selected, bg="#dbe8f5").pack(side="left")
        tk.Button(btn_frame, text="Close", command=self.master.focus_set, bg="#f0f0f0").pack(side="right")

    def _resolve_reports_dir(self) -> Optional[Path]:
        if session_config.output_dir is None:
            return None
        paths = build_output_paths(session_config.output_dir)
        candidate = paths.autopilot_runs_dir / f"year-{session_config.year}"
        if not candidate.exists():
            return None
        return candidate

    def _refresh_list(self) -> None:
        self._tree.delete(*self._tree.get_children())
        self._files = []
        self._reports_dir = self._resolve_reports_dir()
        if self._reports_dir is None:
            self._status_var.set("No autopilot run directory found")
            return
        md_files = sorted(self._reports_dir.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)
        for p in md_files:
            mtime = p.stat().st_mtime
            size = p.stat().st_size
            self._tree.insert("", "end", iid=str(p.name), values=(p.name, p.stat().st_mtime, size))
            self._files.append(p)
        self._status_var.set(f"{len(self._files)} reports")
        if self._files:
            self._tree.selection_set(str(self._files[0].name))
            self._show_preview(self._files[0])

    def _open_folder(self) -> None:
        if self._reports_dir is None:
            messagebox.showwarning("No Folder", "No autopilot run folder to open.", parent=self)
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(self._reports_dir))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(self._reports_dir)], check=False)
            else:
                subprocess.run(["xdg-open", str(self._reports_dir)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _open_manual_audit(self) -> None:
        """Open the `manual_data_audit.xlsx` workbook if present in the audit manual-changes folder."""
        if session_config.output_dir is None:
            messagebox.showwarning("Not Configured", "Output directory is not configured.", parent=self)
            return
        paths = build_output_paths(session_config.output_dir)
        audit_path = paths.audit_manual_changes_dir / "manual_data_audit.xlsx"
        if not audit_path.exists():
            messagebox.showwarning("Not Found", "manual_data_audit.xlsx not found in output audit folder.", parent=self)
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(audit_path))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(audit_path)], check=False)
            else:
                subprocess.run(["xdg-open", str(audit_path)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _open_season_audit(self) -> None:
        """Open the most recent season audit workbook from the audit workbooks folder."""
        if session_config.output_dir is None:
            messagebox.showwarning("Not Configured", "Output directory is not configured.", parent=self)
            return
        paths = build_output_paths(session_config.output_dir)
        audit_dir = paths.audit_workbooks_dir
        if not audit_dir.exists():
            messagebox.showwarning("Not Found", "No audit workbooks folder present.", parent=self)
            return
        xlsx_files = sorted(audit_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not xlsx_files:
            messagebox.showwarning("No Files", "No audit workbook files found.", parent=self)
            return
        # Prefer files that contain the season/year in the name
        preferred = None
        for p in xlsx_files:
            if str(session_config.year) in p.name:
                preferred = p
                break
        target = preferred or xlsx_files[0]
        try:
            if sys.platform == "win32":
                os.startfile(str(target))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(target)], check=False)
            else:
                subprocess.run(["xdg-open", str(target)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _open_data_quality_report(self) -> None:
        """Open the data quality markdown report for the current season if present."""
        if session_config.output_dir is None:
            messagebox.showwarning("Not Configured", "Output directory is not configured.", parent=self)
            return
        paths = build_output_paths(session_config.output_dir)
        qdir = paths.quality_data_dir / f"year-{session_config.year}"
        if not qdir.exists():
            messagebox.showwarning("Not Found", "No data-quality report folder for this season.", parent=self)
            return
        # Prefer standard filename data_quality_report.md
        target = qdir / "data_quality_report.md"
        if not target.exists():
            md_files = sorted(qdir.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)
            if not md_files:
                messagebox.showwarning("No Files", "No markdown data-quality reports found.", parent=self)
                return
            target = md_files[0]
        try:
            if sys.platform == "win32":
                os.startfile(str(target))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(target)], check=False)
            else:
                subprocess.run(["xdg-open", str(target)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _open_selected(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        name = sel[0]
        path = self._reports_dir / name
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _show_preview(self, path: Path) -> None:
        try:
            text = path.read_text(encoding="utf-8")
        except Exception:
            text = "(Could not read file)"
        self._preview.configure(state="normal")
        self._preview.delete("1.0", "end")
        self._preview.insert("1.0", text[:20000])
        self._preview.configure(state="disabled")
