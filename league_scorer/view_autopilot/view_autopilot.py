import os
import subprocess
import sys
from datetime import datetime
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

        # Keep direct openers for workbooks only; markdown reports are listed below
        tk.Button(top_frame, text="Open Manual Audit", command=self._open_manual_audit, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Button(top_frame, text="Open Season Audit", command=self._open_season_audit, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Button(top_frame, text="Open Folder", command=self._open_folder, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Button(top_frame, text="Refresh", command=self._refresh_list, bg="#e9f0f7").pack(side="left", padx=6)
        tk.Label(top_frame, textvariable=self._status_var, bg="#f7f9fb", fg="#55666f").pack(side="right")

        middle = tk.PanedWindow(self, orient="horizontal")
        middle.pack(fill="both", expand=True, padx=10, pady=10)

        # File list (will include markdown reports from multiple output locations)
        list_frame = tk.Frame(middle, bg="#f7f9fb")
        self._tree = ttk.Treeview(list_frame, columns=("name", "modified", "size"), show="headings", selectmode="browse")
        self._tree.heading("name", text="Name")
        self._tree.heading("modified", text="Modified")
        self._tree.heading("size", text="Size")
        self._tree.column("name", width=320, anchor="w")
        self._tree.column("modified", width=140, anchor="w")
        self._tree.column("size", width=80, anchor="e")
        self._tree.pack(fill="both", expand=True, side="left")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self._tree.yview)
        scrollbar.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=scrollbar.set)
        self._tree.bind("<Double-1>", lambda e: self._open_selected())
        self._tree.bind("<<TreeviewSelect>>", lambda e: self._on_tree_select())

        middle.add(list_frame, minsize=200)

        # Preview
        preview_frame = tk.Frame(middle, bg="#ffffff")
        self._preview = tk.Text(preview_frame, wrap="word", state="disabled", bg="#ffffff", relief="flat")
        self._preview.pack(fill="both", expand=True)
        middle.add(preview_frame, minsize=320)

        hint = tk.Label(
            self,
            text="Double-click a report to open it. Use Open Folder to see all files in Explorer.",
            font=("Segoe UI", 9),
            bg="#f7f9fb",
            fg="#55666f",
        )
        hint.pack(fill="x", padx=10, pady=(0, 8))

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
        # collect markdown reports from multiple well-known locations
        paths = []
        base = build_output_paths(session_config.output_dir) if session_config.output_dir else None
        if base:
            # autopilot runs for this year
            ap = base.autopilot_runs_dir / f"year-{session_config.year}"
            paths.append(ap)
            # data quality for this year
            dq = base.quality_data_dir / f"year-{session_config.year}"
            paths.append(dq)
            # staged checks
            sc = base.quality_staged_checks_dir
            paths.append(sc)

        md_set = {}
        for p in paths:
            if not p or not p.exists():
                continue
            for f in p.glob("*.md"):
                md_set[str(f.resolve())] = f

        md_files = sorted(md_set.values(), key=lambda p: p.stat().st_mtime, reverse=True)
        for f in md_files:
            mtime = f.stat().st_mtime
            size = f.stat().st_size
            iid = str(f.resolve())
            self._tree.insert("", "end", iid=iid, values=(f.name, self._format_mtime(mtime), size))
            self._files.append(f)
        count = len(self._files)
        if self._reports_dir is None:
            self._reports_dir = self._resolve_reports_dir()
        location = str(self._reports_dir) if self._reports_dir is not None else "No reports folder"
        self._status_var.set(f"{count} reports  •  {location}")
        if self._files:
            first_iid = str(self._files[0].resolve())
            self._tree.selection_set(first_iid)
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

    def _open_staged_checks_report(self) -> None:
        """Open the staged checks markdown report (staged_checks_report.md)."""
        if session_config.output_dir is None:
            messagebox.showwarning("Not Configured", "Output directory is not configured.", parent=self)
            return
        paths = build_output_paths(session_config.output_dir)
        staged_dir = paths.quality_staged_checks_dir
        if not staged_dir.exists():
            messagebox.showwarning("Not Found", "No staged-checks folder present.", parent=self)
            return
        target = staged_dir / "staged_checks_report.md"
        if not target.exists():
            md_files = sorted(staged_dir.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)
            if not md_files:
                messagebox.showwarning("No Files", "No staged-check markdown reports found.", parent=self)
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
        iid = sel[0]
        path = Path(iid)
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _on_tree_select(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        iid = sel[0]
        path = Path(iid)
        if path.exists():
            self._show_preview(path)

    def _show_preview(self, path: Path) -> None:
        try:
            text = path.read_text(encoding="utf-8")
        except Exception:
            text = "(Could not read file)"
        self._preview.configure(state="normal")
        self._preview.delete("1.0", "end")
        self._preview.insert("1.0", text[:20000])
        self._preview.configure(state="disabled")

    def _format_mtime(self, mtime: float) -> str:
        try:
            dt = datetime.fromtimestamp(float(mtime))
            return dt.strftime("%d-%m-%Y %H:%M:%S")
        except Exception:
            return str(mtime)
