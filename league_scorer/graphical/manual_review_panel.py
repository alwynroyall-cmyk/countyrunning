from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Dict, List

import pandas as pd

from ..name_lookup import append_name_corrections, read_name_lookup_state
from ..session_config import config as session_config
from .runner_history_viewer import RunnerHistoryPanel
from .club_match_dialog import (
    _append_club_conversions,
    _build_write_result_text,
    _read_club_lookup_state,
    _summarise_selection,
    _set_widget_tree_bg as _legacy_set_widget_tree_bg,  # kept for compatibility
    ClubMatchCandidate,
)
from .manual_review_dialog import (
    _build_club_candidates,
    _build_name_candidates,
    _summarise_name_selection,
    _build_confirmation_text,
    _build_name_write_result_text,
)
from .results_workbook import find_latest_results_workbook, sorted_race_sheet_names

from .manual_review_helpers import (
    WRRL_NAVY,
    WRRL_GREEN,
    WRRL_LIGHT,
    WRRL_WHITE,
    _DEFAULT_BG,
    _HIGHLIGHT_BG,
    proper_case,
    make_header_label,
    build_scroll_frame,
    set_widget_tree_bg,
    scan_workbook_for_runner_state,
    detect_runner_anomalies,
    add_club_row,
    add_name_row,
)


class ManualReviewPanel(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        club_df: pd.DataFrame | None = None,
        clubs_path: Path | None = None,
        name_df: pd.DataFrame | None = None,
        names_path: Path | None = None,
        back_callback=None,
    ) -> None:
        super().__init__(parent, bg=_DEFAULT_BG)
        self._changed_runners: set[str] = set()
        self._back_callback = back_callback
        self._clubs_path = clubs_path
        self._names_path = names_path
        self._club_candidates: List[ClubMatchCandidate] = _build_club_candidates(club_df)
        self._name_candidates = _build_name_candidates(name_df)
        self._club_vars: Dict[str, tk.BooleanVar] = {}
        self._name_vars: Dict[str, tk.BooleanVar] = {}
        self._name_value_vars: Dict[str, tk.StringVar] = {}
        self._club_row_frames: Dict[str, tk.Frame] = {}
        self._name_row_frames: Dict[str, tk.Frame] = {}
        self._runner_panel: RunnerHistoryPanel | None = None
        self._anomaly_tree: ttk.Treeview | None = None
        self._anomaly_status_var: tk.StringVar | None = None
        self._anomaly_find_btn: tk.Button | None = None
        self._anomaly_scan_queue: "queue.Queue | None" = None
        self._split: tk.PanedWindow | None = None
        self._notebook: ttk.Notebook | None = None
        self._anomaly_tab: tk.Frame | None = None
        self._apply_updates_btn: tk.Button | None = None
        self._build_ui()

    # -----------------------------
    # UI construction
    # -----------------------------

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=WRRL_NAVY, padx=14, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Manual Review",
            font=("Segoe UI", 15, "bold"),
            bg=WRRL_NAVY,
            fg=WRRL_WHITE,
        ).pack(side="left")

        if self._back_callback:
            tk.Button(
                header,
                text="🏠 Dashboard",
                font=("Segoe UI", 10, "bold"),
                bg=WRRL_LIGHT,
                fg=WRRL_GREEN,
                relief="flat",
                padx=10,
                pady=4,
                cursor="hand2",
                command=self._back_callback,
                activebackground="#1f5632",
                activeforeground=WRRL_GREEN,
            ).pack(side="right")

        subtitle = tk.Label(
            self,
            text="Review suggestions on the left and apply runner corrections on the right.",
            font=("Segoe UI", 10),
            bg=_DEFAULT_BG,
            fg="#4a5868",
            anchor="w",
            justify="left",
        )
        subtitle.pack(fill="x", padx=16, pady=(12, 10))

        split = tk.PanedWindow(self, orient="horizontal", sashrelief="flat", bg=_DEFAULT_BG)
        split.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self._split = split

        left = tk.Frame(split, bg=_DEFAULT_BG)
        right = tk.Frame(split, bg=_DEFAULT_BG)
        split.add(left, minsize=420, width=560)
        split.add(right, minsize=420)
        self.after(0, self._set_equal_split)

        notebook = ttk.Notebook(left)
        notebook.pack(fill="both", expand=True)
        self._notebook = notebook
        self._anomaly_tab = None
        self._apply_updates_btn = None

        added_tabs = 0

        if self._club_candidates:
            club_tab = tk.Frame(notebook, bg=_DEFAULT_BG)
            notebook.add(club_tab, text=f"Clubs ({len(self._club_candidates)})")
            self._build_club_tab(club_tab)
            added_tabs += 1

        if self._name_candidates:
            name_tab = tk.Frame(notebook, bg=_DEFAULT_BG)
            notebook.add(name_tab, text=f"Names ({len(self._name_candidates)})")
            self._build_name_tab(name_tab)
            added_tabs += 1

        anomaly_tab = tk.Frame(notebook, bg=_DEFAULT_BG)
        self._anomaly_tab = anomaly_tab
        notebook.add(anomaly_tab, text="Runner Anomalies")
        self._build_runner_anomalies_tab(anomaly_tab)
        added_tabs += 1

        if not added_tabs:
            notebook.destroy()
            tk.Label(
                self,
                text="No manual review suggestions are available in this workbook.",
                font=("Segoe UI", 10),
                bg=_DEFAULT_BG,
                fg="#4a5868",
            ).pack(anchor="w", padx=16, pady=(0, 12))

        action_bar = tk.Frame(left, bg=_DEFAULT_BG)
        action_bar.pack(fill="x", pady=(8, 0))

        self._apply_updates_btn = tk.Button(
            action_bar,
            text="Apply Selected Updates",
            command=self._open_confirmation,
            font=("Segoe UI", 10, "bold"),
            bg="#2d7a4a",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=7,
        )
        self._apply_updates_btn.pack(side="right")

        notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._runner_panel = RunnerHistoryPanel(right, back_callback=None, parent_panel=self)
        self._runner_panel.pack(fill="both", expand=True)

    def _set_equal_split(self) -> None:
        if self._split is None:
            return
        try:
            total_width = self._split.winfo_width()
            if total_width > 0:
                self._split.sash_place(0, total_width // 2, 0)
        except Exception:
            return

    def _on_tab_changed(self, _event=None):
        if not self._notebook or not self._anomaly_tab or not self._apply_updates_btn:
            return
        selected = self._notebook.select()
        if selected == str(self._anomaly_tab):
            self._apply_updates_btn.pack_forget()
        else:
            if not self._apply_updates_btn.winfo_ismapped():
                self._apply_updates_btn.pack(side="right")

    # -----------------------------
    # Tabs
    # -----------------------------

    def _build_runner_anomalies_tab(self, parent: tk.Misc) -> None:
        intro = tk.Label(
            parent,
            text="Runners with cross-race club/gender/category anomalies. Select a row to open on the right.",
            font=("Segoe UI", 9),
            bg=_DEFAULT_BG,
            fg="#4a5868",
            anchor="w",
            justify="left",
        )
        intro.pack(fill="x", padx=10, pady=(8, 4))

        toolbar = tk.Frame(parent, bg=_DEFAULT_BG)
        toolbar.pack(fill="x", padx=10, pady=(0, 6))
        self._anomaly_find_btn = tk.Button(
            toolbar,
            text="Find Anomalies",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=14,
            pady=5,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
            command=self._refresh_runner_anomaly_list,
        )
        self._anomaly_find_btn.pack(side="left")

        self._anomaly_status_var = tk.StringVar(
            value="Click ‘Find Anomalies’ to scan for cross-race inconsistencies."
        )
        tk.Label(
            toolbar,
            textvariable=self._anomaly_status_var,
            font=("Segoe UI", 9, "italic"),
            bg=_DEFAULT_BG,
            fg="#4a5868",
        ).pack(side="left", padx=(12, 0))

        table_wrap = tk.Frame(parent, bg=_DEFAULT_BG)
        table_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        tree = ttk.Treeview(
            table_wrap,
            show="headings",
            columns=("Status", "Runner", "Anomalies", "Details"),
            style="RunnerHistory.Treeview",
        )
        tree.heading("Status", text="✓", anchor="center")
        tree.heading("Runner", text="Runner", anchor="w")
        tree.heading("Anomalies", text="Anomalies", anchor="w")
        tree.heading("Details", text="Details", anchor="w")
        tree.column("Status", width=28, anchor="center", stretch=False)
        tree.column("Runner", width=220, anchor="w", stretch=False)
        tree.column("Anomalies", width=150, anchor="w", stretch=False)
        tree.column("Details", width=520, anchor="w", stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        tree.bind("<<TreeviewSelect>>", self._on_runner_anomaly_selected)
        self._anomaly_tree = tree

    def _build_club_tab(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg="#dbe5ee")
        header.pack(fill="x")
        make_header_label(header, "Tick", 0, 8)
        make_header_label(header, "Current Club", 1, 24)
        make_header_label(header, "Proposed Club", 2, 24)
        make_header_label(header, "Message", 3, 48)

        rows_frame = build_scroll_frame(parent)
        for candidate in self._club_candidates:
            var = tk.BooleanVar(value=False)
            self._club_vars[candidate.current_club] = var
            add_club_row(rows_frame, candidate, var, self._club_row_frames)

    def _build_name_tab(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg="#dbe5ee")
        header.pack(fill="x")
        make_header_label(header, "Tick", 0, 8)
        make_header_label(header, "Current Name", 1, 24)
        make_header_label(header, "Proposed Name", 2, 28)
        make_header_label(header, "Message", 3, 44)

        rows_frame = build_scroll_frame(parent)
        for candidate in self._name_candidates:
            var = tk.BooleanVar(value=False)
            value_var = tk.StringVar(value=candidate.proposed_name)
            self._name_vars[candidate.current_name] = var
            self._name_value_vars[candidate.current_name] = value_var
            add_name_row(rows_frame, candidate, var, value_var, self._name_row_frames)

    # -----------------------------
    # Anomaly scanning
    # -----------------------------

    def _refresh_runner_anomaly_list(self) -> None:
        if self._anomaly_tree is None:
            return
        if self._anomaly_scan_queue is not None:
            # Scan already in progress
            return

        self._anomaly_tree.delete(*self._anomaly_tree.get_children())
        if self._anomaly_find_btn is not None:
            self._anomaly_find_btn.config(state="disabled", text="Scanning…")
        if self._anomaly_status_var is not None:
            self._anomaly_status_var.set("Scanning…")

        q: queue.Queue = queue.Queue()
        self._anomaly_scan_queue = q

        def worker():
            state = scan_workbook_for_runner_state()
            anomalies = detect_runner_anomalies(state)
            q.put(anomalies)

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_anomaly_scan)

    def _poll_anomaly_scan(self) -> None:
        q = self._anomaly_scan_queue
        if q is None or self._anomaly_tree is None:
            return
        try:
            anomalies = q.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_anomaly_scan)
            return

        self._anomaly_scan_queue = None
        tree = self._anomaly_tree
        tree.delete(*tree.get_children())

        for idx, row in enumerate(anomalies):
            tag = "even" if idx % 2 else "odd"
            runner_name = row["runner"]
            status = "✓" if runner_name in self._changed_runners else ""
            tree.insert(
                "",
                "end",
                values=(status, runner_name, row["anomalies"], row["details"]),
                tags=(tag,),
            )

        if anomalies:
            if self._anomaly_status_var is not None:
                self._anomaly_status_var.set(f"{len(anomalies)} runner(s) with anomalies found.")
        else:
            if self._anomaly_status_var is not None:
                self._anomaly_status_var.set("No anomalies found.")
            tree.insert("", "end", values=("No anomalies found", "", "", ""), tags=("odd",))

        if self._anomaly_find_btn is not None:
            self._anomaly_find_btn.config(state="normal", text="Find Anomalies")

    def _on_runner_anomaly_selected(self, _event=None) -> None:
        if self._anomaly_tree is None or self._runner_panel is None:
            return
        selected = self._anomaly_tree.selection()
        if not selected:
            return
        values = self._anomaly_tree.item(selected[0], "values")
        if not values:
            return
        runner_name = str(values[1]).strip() if len(values) > 1 else ""
        if not runner_name or runner_name.lower() == "no anomalies found":
            return
        self._runner_panel.select_runner(runner_name)

        # Auto-scroll to selected row for better UX
        self._anomaly_tree.see(selected[0])

    # -----------------------------
    # Confirmation / apply updates
    # -----------------------------

    def _open_confirmation(self) -> None:
        selected_clubs = [
            candidate
            for candidate in self._club_candidates
            if self._club_vars.get(candidate.current_club, tk.BooleanVar()).get()
        ]
        selected_names = [
            {
                "current_name": candidate.current_name,
                "proposed_name": self._name_value_vars[candidate.current_name].get().strip(),
                "message": candidate.message,
            }
            for candidate in self._name_candidates
            if self._name_vars.get(candidate.current_name, tk.BooleanVar()).get()
        ]

        if not selected_clubs and not selected_names:
            messagebox.showwarning(
                "No updates selected",
                "Tick at least one club or name correction before continuing.",
                parent=self,
            )
            return

        if selected_clubs and self._clubs_path is None:
            messagebox.showerror(
                "Club Lookup Missing",
                "clubs.xlsx could not be resolved from the control folder, so club corrections cannot be written yet.",
                parent=self,
            )
            return

        if selected_names and self._names_path is None:
            messagebox.showerror(
                "Name Corrections Path Missing",
                "name_corrections.xlsx could not be resolved from the control folder, so name corrections cannot be written yet.",
                parent=self,
            )
            return

        club_summary = None
        if selected_clubs and self._clubs_path is not None:
            club_summary = _summarise_selection(
                selected_clubs,
                _read_club_lookup_state(self._clubs_path),
            )

        name_summary = None
        if selected_names and self._names_path is not None:
            name_summary = _summarise_name_selection(
                selected_names,
                read_name_lookup_state(self._names_path),
            )

        _ManualReviewConfirmationDialog(self, club_summary, name_summary)

    def highlight_club_conflicts(self, conflict_clubs: List[str]) -> None:
        conflicts = set(conflict_clubs)
        for current_club, frame in self._club_row_frames.items():
            bg = _HIGHLIGHT_BG if current_club in conflicts else _DEFAULT_BG
            set_widget_tree_bg(frame, bg)

    def highlight_name_conflicts(self, conflict_names: List[str]) -> None:
        conflicts = set(conflict_names)
        for current_name, frame in self._name_row_frames.items():
            bg = _HIGHLIGHT_BG if current_name in conflicts else _DEFAULT_BG
            set_widget_tree_bg(frame, bg)


class _ManualReviewConfirmationDialog(tk.Toplevel):
    def __init__(self, panel: ManualReviewPanel, club_summary: dict | None, name_summary: dict | None) -> None:
        super().__init__(panel)
        self.title("Confirm Manual Review Updates")
        self.geometry("860x520")
        self.configure(bg=_DEFAULT_BG)
        self.transient(panel.winfo_toplevel())
        self._panel = panel
        self._club_summary = club_summary
        self._name_summary = name_summary

        self._build_ui()
        self.grab_set()

    def destroy(self):
        if self._panel and hasattr(self._panel, "_runner_panel") and self._panel._runner_panel is not None:
            runner_name = self._panel._runner_panel._runner_var.get().strip()
            if runner_name:
                pc_runner = proper_case(runner_name)
                self._panel._changed_runners.add(pc_runner)
                self._panel._refresh_runner_anomaly_list()
                self._panel._runner_panel.select_runner(runner_name)
        super().destroy()

    def _build_ui(self) -> None:
        tk.Label(
            self,
            text="Confirm Manual Review Updates",
            font=("Segoe UI", 15, "bold"),
            bg=_DEFAULT_BG,
            fg="#22313f",
        ).pack(anchor="w", padx=16, pady=(14, 8))

        text = tk.Text(self, wrap="word", font=("Segoe UI", 10), bg="#ffffff", fg="#22313f")
        text.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        text.insert("1.0", _build_confirmation_text(self._club_summary, self._name_summary))
        text.config(state="disabled")

        action_bar = tk.Frame(self, bg=_DEFAULT_BG)
        action_bar.pack(fill="x", padx=16, pady=(0, 16))

        tk.Button(
            action_bar,
            text="Back To Selection",
            command=self._back_to_selection,
            font=("Segoe UI", 10),
            bg="#d0d4de",
            fg="#22313f",
            relief="flat",
            padx=12,
            pady=6,
        ).pack(side="left")

        has_conflicts = bool(
            (self._club_summary and self._club_summary["conflicts"])
            or (self._name_summary and self._name_summary["conflicts"])
        )
        tk.Button(
            action_bar,
            text="Confirm Write",
            command=self._confirm_write,
            state="disabled" if has_conflicts else "normal",
            font=("Segoe UI", 10, "bold"),
            bg="#2d7a4a",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=7,
        ).pack(side="right")

    def _back_to_selection(self) -> None:
        if self._club_summary:
            self._panel.highlight_club_conflicts(
                [item.current_club for item in self._club_summary["conflicts"]]
            )
        if self._name_summary:
            self._panel.highlight_name_conflicts(
                [item["current_name"] for item in self._name_summary["conflicts"]]
            )
        self.destroy()

    def _confirm_write(self) -> None:
        messages = []
        if self._club_summary and self._panel._clubs_path is not None:
            club_result = _append_club_conversions(
                self._panel._clubs_path,
                self._club_summary["new_rows"],
            )
            messages.append(_build_write_result_text(club_result))
        if self._name_summary and self._panel._names_path is not None:
            name_result = append_name_corrections(
                self._panel._names_path,
                self._name_summary["new_rows"],
            )
            messages.append(_build_name_write_result_text(name_result))

        self.destroy()
        messagebox.showinfo("Manual Review Updated", "\n\n".join(messages), parent=self._panel)