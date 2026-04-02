"""Issue Review panel — load actionable audit issues and navigate to source."""
from __future__ import annotations

import os
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog, ttk
from typing import Callable

import pandas as pd

from ..audit_data_service import ACTIONABLE_COLUMNS, find_latest_audit_workbook, load_actionable_issues
from ..issue_resolution_service import (
    apply_quick_fix_for_issue,
    quick_fix_prompt,
    quick_fix_requires_input,
    supports_quick_fix,
)
from ..issue_tracking import build_issue_identity
from ..session_config import config as session_config
from .dashboard import WRRL_GREEN, WRRL_LIGHT, WRRL_NAVY, WRRL_WHITE

# ── constants ─────────────────────────────────────────────────────────────────

_COLUMNS = ACTIONABLE_COLUMNS

_DISPLAY_COLS = ["Issue Code", "Severity", "Race", "Name", "Club", "Message"]

_COL_WIDTHS = {
    "Issue Code": 120,
    "Severity": 70,
    "Race": 160,
    "Name": 160,
    "Club": 140,
    "Message": 340,
}

# ── panel ─────────────────────────────────────────────────────────────────────


class IssueReviewPanel(tk.Frame):
    """Inline panel showing all actionable audit issues with navigation to source."""

    def __init__(
        self,
        parent: tk.Widget,
        back_callback: Callable[[], None] | None = None,
        view_runner_callback: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__(parent, bg=WRRL_LIGHT)
        self._back_callback = back_callback
        self._view_runner_callback = view_runner_callback
        self._workbook_path: Path | None = None
        self._df: pd.DataFrame = pd.DataFrame(columns=_COLUMNS)
        self._filtered_df: pd.DataFrame = pd.DataFrame(columns=_COLUMNS)
        self._style = ttk.Style(self)
        self._configure_styles()
        self._build_ui()
        self._load_issues()

    # ── styles ────────────────────────────────────────────────────────────────

    def _configure_styles(self) -> None:
        self._style.configure(
            "Issues.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#22313f",
            rowheight=26,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self._style.configure(
            "Issues.Treeview.Heading",
            background="#dbe5ee",
            foreground="#22313f",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        self._style.map(
            "Issues.Treeview",
            background=[("selected", "#8fb3d1")],
            foreground=[("selected", "#102030")],
        )

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._build_header()
        self._build_filter_bar()
        self._build_table()
        self._build_detail_pane()
        self._build_action_bar()

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=WRRL_NAVY, padx=14, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Issue Review",
            font=("Segoe UI", 15, "bold"),
            bg=WRRL_NAVY,
            fg=WRRL_WHITE,
        ).pack(side="left")

        if self._back_callback:
            tk.Button(
                header,
                text="\u25c4 Dashboard",
                font=("Segoe UI", 10, "bold"),
                bg=WRRL_GREEN,
                fg=WRRL_WHITE,
                relief="flat",
                padx=10,
                pady=4,
                cursor="hand2",
                command=self._back_callback,
                activebackground="#1f5632",
                activeforeground=WRRL_WHITE,
            ).pack(side="right")

        self._count_var = tk.StringVar(value="")
        tk.Label(
            header,
            textvariable=self._count_var,
            font=("Segoe UI", 10),
            bg=WRRL_NAVY,
            fg="#b8cfe0",
        ).pack(side="right", padx=(0, 16))

    def _build_filter_bar(self) -> None:
        bar = tk.Frame(self, bg="#dbe1e8", padx=14, pady=8)
        bar.pack(fill="x")

        tk.Label(
            bar,
            text="Filter by code:",
            font=("Segoe UI", 10),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
        ).pack(side="left")

        self._code_var = tk.StringVar(value="All")
        self._code_combo = ttk.Combobox(
            bar,
            textvariable=self._code_var,
            state="readonly",
            width=20,
        )
        self._code_combo.pack(side="left", padx=(4, 16))
        self._code_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_filter())

        tk.Label(
            bar,
            text="Race:",
            font=("Segoe UI", 10),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
        ).pack(side="left")

        self._race_var = tk.StringVar(value="All")
        self._race_combo = ttk.Combobox(
            bar,
            textvariable=self._race_var,
            state="readonly",
            width=28,
        )
        self._race_combo.pack(side="left", padx=(4, 16))
        self._race_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_filter())

        tk.Button(
            bar,
            text="Reload",
            font=("Segoe UI", 9),
            bg="#c5cdd6",
            fg=WRRL_NAVY,
            relief="flat",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._load_issues,
        ).pack(side="left", padx=(0, 4))

        tk.Button(
            bar,
            text="Clear Filter",
            font=("Segoe UI", 9),
            bg="#c5cdd6",
            fg=WRRL_NAVY,
            relief="flat",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._clear_filter,
        ).pack(side="left")

    def _build_table(self) -> None:
        outer = tk.Frame(self, bg=WRRL_LIGHT)
        outer.pack(fill="both", expand=True, padx=14, pady=(8, 0))

        scrollbar = ttk.Scrollbar(outer, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        hscrollbar = ttk.Scrollbar(outer, orient="horizontal")
        hscrollbar.pack(side="bottom", fill="x")

        self._tree = ttk.Treeview(
            outer,
            columns=_DISPLAY_COLS,
            show="headings",
            style="Issues.Treeview",
            yscrollcommand=scrollbar.set,
            xscrollcommand=hscrollbar.set,
            selectmode="browse",
        )
        self._tree.pack(fill="both", expand=True)
        scrollbar.config(command=self._tree.yview)
        hscrollbar.config(command=self._tree.xview)

        for col in _DISPLAY_COLS:
            self._tree.heading(col, text=col, anchor="w")
            self._tree.column(col, width=_COL_WIDTHS.get(col, 120), anchor="w", stretch=True)

        self._tree.tag_configure("warning", foreground="#8b4500")
        self._tree.tag_configure("error", foreground="#8b0000")
        self._tree.tag_configure("info", foreground="#1a4a7a")

        self._tree.bind("<<TreeviewSelect>>", self._on_select)

    def _build_detail_pane(self) -> None:
        detail_frame = tk.Frame(self, bg="#eef2f6", relief="flat", bd=0)
        detail_frame.pack(fill="x", padx=14, pady=(4, 0))

        # Issue code + key header
        top_row = tk.Frame(detail_frame, bg="#eef2f6")
        top_row.pack(fill="x", padx=10, pady=(6, 0))

        self._detail_code_var = tk.StringVar(value="")
        tk.Label(
            top_row,
            textvariable=self._detail_code_var,
            font=("Segoe UI", 11, "bold"),
            bg="#eef2f6",
            fg=WRRL_NAVY,
            anchor="w",
        ).pack(side="left")

        self._detail_key_var = tk.StringVar(value="")
        tk.Label(
            top_row,
            textvariable=self._detail_key_var,
            font=("Segoe UI", 10),
            bg="#eef2f6",
            fg="#555555",
            anchor="w",
        ).pack(side="left", padx=(14, 0))

        # Full message
        self._detail_msg_var = tk.StringVar(value="Select an issue to see details.")
        tk.Label(
            detail_frame,
            textvariable=self._detail_msg_var,
            font=("Segoe UI", 10),
            bg="#eef2f6",
            fg="#22313f",
            anchor="w",
            wraplength=900,
            justify="left",
        ).pack(fill="x", padx=10, pady=(2, 0))

        # Next step
        self._detail_next_var = tk.StringVar(value="")
        tk.Label(
            detail_frame,
            textvariable=self._detail_next_var,
            font=("Segoe UI", 9, "italic"),
            bg="#eef2f6",
            fg="#1a5632",
            anchor="w",
            wraplength=900,
            justify="left",
        ).pack(fill="x", padx=10, pady=(2, 6))

    def _build_action_bar(self) -> None:
        self._action_bar = tk.Frame(self, bg=WRRL_LIGHT, padx=14, pady=8)
        self._action_bar.pack(fill="x")

        self._open_file_btn = tk.Button(
            self._action_bar,
            text="📂 Open Source File",
            font=("Segoe UI", 10, "bold"),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2",
            state="disabled",
            command=self._open_source_file,
            activebackground="#c5cdd6",
            activeforeground=WRRL_NAVY,
        )
        self._open_file_btn.pack(side="left", padx=(0, 8))

        self._runner_history_btn = tk.Button(
            self._action_bar,
            text="🏃 View in Runner History",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2",
            state="disabled",
            command=self._go_to_runner_history,
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
        )
        self._runner_history_btn.pack(side="left")

        self._quick_fix_btn = tk.Button(
            self._action_bar,
            text="⚡ Quick Fix",
            font=("Segoe UI", 10, "bold"),
            bg="#0f6a8c",
            fg=WRRL_WHITE,
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2",
            state="disabled",
            command=self._apply_quick_fix,
            activebackground="#0c5875",
            activeforeground=WRRL_WHITE,
        )
        self._quick_fix_btn.pack(side="left", padx=(8, 0))

        # label to show the file path of the selected row
        self._source_path_var = tk.StringVar(value="")
        tk.Label(
            self._action_bar,
            textvariable=self._source_path_var,
            font=("Segoe UI", 8),
            bg=WRRL_LIGHT,
            fg="#888888",
            anchor="w",
        ).pack(side="left", padx=(16, 0))

    # ── data loading ──────────────────────────────────────────────────────────

    def _load_issues(self) -> None:
        """Locate the audit workbook and load all actionable issues."""
        workbook = find_latest_audit_workbook()
        if workbook is None:
            self._workbook_path = None
            self._df = pd.DataFrame(columns=_COLUMNS)
            self._count_var.set("No audit workbook found — run Audit Races first.")
            self._code_combo["values"] = ["All"]
            self._code_var.set("All")
            self._race_combo["values"] = ["All"]
            self._race_var.set("All")
            self._render_tree(self._df)
            return

        try:
            self._workbook_path = workbook
            self._df = load_actionable_issues(workbook)
        except Exception as exc:
            messagebox.showerror("Load Error", f"Could not read audit workbook:\n{exc}")
            self._workbook_path = workbook
            self._df = pd.DataFrame(columns=_COLUMNS)

        # Populate filter dropdowns
        codes = ["All"] + sorted(
            [value for value in self._df["Issue Code"].dropna().unique().tolist() if str(value).strip()]
        )
        self._code_combo["values"] = codes
        if self._code_var.get() not in codes:
            self._code_var.set("All")

        races = ["All"] + sorted(
            [value for value in self._df["Race"].dropna().unique().tolist() if str(value).strip()]
        )
        self._race_combo["values"] = races
        if self._race_var.get() not in races:
            self._race_var.set("All")

        self._apply_filter()

    def _apply_filter(self) -> None:
        df = self._df.copy()
        code = self._code_var.get()
        race = self._race_var.get()
        if code != "All":
            df = df[df["Issue Code"] == code]
        if race != "All":
            df = df[df["Race"] == race]
        self._filtered_df = df.reset_index(drop=True)
        self._render_tree(self._filtered_df)

    def _clear_filter(self) -> None:
        self._code_var.set("All")
        self._race_var.set("All")
        self._apply_filter()

    # ── rendering ─────────────────────────────────────────────────────────────

    def _render_tree(self, df: pd.DataFrame) -> None:
        self._tree.delete(*self._tree.get_children())
        total = len(self._df)
        shown = len(df)
        workbook_name = self._workbook_path.name if self._workbook_path else ""
        base = (
            f"{shown} issue{'s' if shown != 1 else ''} shown  (total {total})"
            if total > 0
            else "No actionable issues found."
        )
        self._count_var.set(f"{base}  |  {workbook_name}" if workbook_name else base)

        for _, row in df.iterrows():
            code = str(row.get("Issue Code", ""))
            severity = str(row.get("Severity", "")).lower()
            tag = "error" if severity == "error" else ("warning" if severity == "warning" else "info")
            values = [str(row.get(c, "")) for c in _DISPLAY_COLS]
            self._tree.insert("", "end", values=values, tags=(tag,))

        # Clear detail pane
        self._detail_code_var.set("")
        self._detail_key_var.set("")
        self._detail_msg_var.set("Select an issue to see details.")
        self._detail_next_var.set("")
        self._source_path_var.set("")
        self._open_file_btn.config(state="disabled")
        self._runner_history_btn.config(state="disabled")
        self._quick_fix_btn.config(state="disabled", text="⚡ Quick Fix")

    # ── selection handling ────────────────────────────────────────────────────

    def _on_select(self, _event: tk.Event) -> None:
        selection = self._tree.selection()
        if not selection:
            return
        item = selection[0]
        idx = self._tree.index(item)
        if idx >= len(self._filtered_df):
            return
        row = self._filtered_df.iloc[idx]
        self._show_detail(row)

    def _show_detail(self, row: pd.Series) -> None:
        code = str(row.get("Issue Code", ""))
        key = str(row.get("Key", ""))
        name = str(row.get("Name", ""))
        race = str(row.get("Race", ""))
        message = str(row.get("Message", ""))
        next_step = str(row.get("Next Step", ""))

        self._detail_code_var.set(f"{code}")
        key_parts = []
        if key:
            key_parts.append(f"Key: {key}")
        if name:
            key_parts.append(f"Runner: {name}")
        self._detail_key_var.set("  |  ".join(key_parts))
        self._detail_msg_var.set(message or "—")
        self._detail_next_var.set(f"Next step: {next_step}" if next_step else "")

        # Resolve the source file path
        source_row = str(row.get("Source Row", ""))
        source_path = self._resolve_source_file(race)
        if source_path:
            self._source_path_var.set(str(source_path))
            self._open_file_btn.config(state="normal")
        else:
            self._source_path_var.set("")
            self._open_file_btn.config(state="disabled")

        # Enable runner history button when we have a name
        if name and self._view_runner_callback:
            self._runner_history_btn.config(state="normal")
        else:
            self._runner_history_btn.config(state="disabled")

        if supports_quick_fix(code):
            self._quick_fix_btn.config(state="normal", text=f"⚡ Quick Fix {code}")
        else:
            self._quick_fix_btn.config(state="disabled", text="⚡ Quick Fix")

        # Store currently selected row data for action callbacks
        self._selected_row = row

    # ── source file resolution ────────────────────────────────────────────────

    def _resolve_source_file(self, race_value: str) -> Path | None:
        """Try to find the input race file matching the Race column value."""
        if not race_value or not session_config.input_dir:
            return None
        input_dir = Path(session_config.input_dir)
        if not input_dir.exists():
            return None
        # The Race column contains the workbook stem (e.g. "Race 3 - Malmesbury")
        # Try direct stem match first, then partial match
        for suffix in (".xlsx", ".xls", ".xlsm"):
            exact = input_dir / f"{race_value}{suffix}"
            if exact.exists():
                return exact
        # Partial match — race column value as substring of filename
        for f in input_dir.iterdir():
            if f.suffix.lower() in {".xlsx", ".xls", ".xlsm"}:
                if race_value.lower() in f.stem.lower() or f.stem.lower() in race_value.lower():
                    return f
        return None

    # ── actions ───────────────────────────────────────────────────────────────

    def _open_source_file(self) -> None:
        path_text = self._source_path_var.get()
        if not path_text:
            messagebox.showwarning("No Source File", "Could not locate the source file for this issue.")
            return
        path = Path(path_text)
        if not path.exists():
            messagebox.showerror("File Not Found", f"Source file not found:\n{path}")
            return
        try:
            os.startfile(str(path))  # Windows
        except AttributeError:
            subprocess.run(["open" if os.name == "posix" else "xdg-open", str(path)], check=False)

    def _go_to_runner_history(self) -> None:
        if not hasattr(self, "_selected_row"):
            return
        name = str(self._selected_row.get("Name", "")).strip()
        if not name:
            return
        if self._view_runner_callback:
            self._view_runner_callback(name)
        else:
            messagebox.showinfo("Runner History", f"Select runner: {name}\nin the Runner History panel.")

    def _apply_quick_fix(self) -> None:
        if not hasattr(self, "_selected_row"):
            return

        issue = {
            col: str(self._selected_row.get(col, "")).strip()
            for col in _COLUMNS
        }
        issue_code = issue.get("Issue Code", "")
        runner_name = issue.get("Name", "runner")

        if not supports_quick_fix(issue_code):
            messagebox.showinfo(
                "Manual Review Required",
                f"No quick fix is available for {issue_code}.",
                parent=self,
            )
            return

        target_value = None
        if quick_fix_requires_input(issue_code):
            title, prompt = quick_fix_prompt(issue_code, runner_name)
            entered = simpledialog.askstring(title, prompt, parent=self)
            if entered is None:
                return
            target_value = entered.strip()

        result = apply_quick_fix_for_issue(
            issue,
            input_dir=session_config.input_dir,
            target_value=target_value,
        )

        if result.success:
            if result.verified_resolved and result.issue_identity:
                identities = self._df.apply(lambda row: build_issue_identity(row.to_dict()), axis=1)
                self._df = self._df.loc[~identities.eq(result.issue_identity.lower())].reset_index(drop=True)
                self._apply_filter()
            messagebox.showinfo(
                "Quick Fix Applied",
                result.message + "\n\nRe-run Audit to refresh this issue list.",
                parent=self,
            )
        else:
            messagebox.showwarning("Quick Fix Not Applied", result.message, parent=self)
