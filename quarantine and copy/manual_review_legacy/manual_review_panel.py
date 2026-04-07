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
