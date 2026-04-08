"""
dashboard.py — Main dashboard for WRRL League AI.

Provides a professional console with branded header and execution options.
"""

import tkinter as tk
import os
import queue
import subprocess
import sys
import threading
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

from .gui import LeagueScorerApp
from .events_viewer import EventsViewerPanel
from ..events_loader import load_events
from ..common_files import race_discovery_exclusions
from ..raceroster_import import (
    SporthiveRaceNotDirectlyImportableError,
    import_raceroster_results,
)
from ..input_layout import sort_existing_input_files
from ..output_layout import build_output_paths, ensure_output_subdirs, sort_existing_output_files
from ..session_config import config as session_config
from ..source_loader import discover_race_files
from ..structured_logging import log_event
from .import_helpers import (
    RaceImportRequest,
    ask_multiline_page_text,
    prompt_race_import_request,
    run_manual_sporthive_import,
)


# ──────────────────────────────────────────────────────────────────────────────
# Define WRRL brand colors
# ──────────────────────────────────────────────────────────────────────────────

WRRL_NAVY = "#3a4658"      # Dark navy blue from shield
WRRL_GREEN = "#2d7a4a"     # Forest green from shield
WRRL_LIGHT = "#f5f5f5"     # Light gray for text background
WRRL_WHITE = "#ffffff"     # Pure white for text
WRRL_AMBER = "#e6a817"    # Amber used as a warning accent


# ──────────────────────────────────────────────────────────────────────────────
# Autopilot progress dialog
# ──────────────────────────────────────────────────────────────────────────────

class _AutopilotProgressDialog(tk.Toplevel):
    """Modern animated progress dialog: stage timeline + race chip grid."""

    _DEFAULT_STAGE_LABELS = ["Audit", "Safe Fixes", "Quality Checks"]

    # Node colours
    _C_PEND      = "#d9e2ec"
    _C_PEND_TXT  = "#8a9aaa"
    _C_ACTIVE_A  = "#2d7a4a"   # dim pulse
    _C_ACTIVE_B  = "#4bbf74"   # bright pulse
    _C_DONE_NODE = "#1aaa60"
    _C_DONE_LINE = "#1aaa60"
    _C_LINE_IDLE = "#d0dae6"

    # Race chip colours
    _C_CHIP_PEND    = "#dce3ea"
    _C_CHIP_PEND_FG = "#5a6878"
    _C_CHIP_ACT     = "#4bbf74"
    _C_CHIP_ACT_FG  = "#ffffff"
    _C_CHIP_DONE    = "#2d7a4a"
    _C_CHIP_DONE_FG = "#ffffff"

    def __init__(
        self,
        parent: tk.Tk,
        year: int,
        *,
        window_title: str = "Autopilot - WRRL League AI",
        header_text: str = "WRRL League AI Autopilot",
        stage_labels: list[str] | None = None,
        initial_status: str = "Initialising autopilot...",
    ) -> None:
        super().__init__(parent)
        self.title(window_title)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.configure(bg=WRRL_LIGHT)
        self.geometry("600x380")
        self.protocol("WM_DELETE_WINDOW", self._on_close_attempt)

        self._year = year
        self._header_text = header_text
        self._stage_labels = stage_labels or list(self._DEFAULT_STAGE_LABELS)
        self._current_stage = 0
        self._stages_done: set[int] = set()
        self._race_chips: list[tk.Label] = []
        self._pulse_on = True
        self._animating = True
        self._status_var = tk.StringVar(value=initial_status)
        self._substep_var = tk.StringVar(value="")

        self._build()
        self._pulse()

    # ── layout ────────────────────────────────────────────────────────────────

    def _build(self) -> None:
        header = tk.Frame(self, bg=WRRL_NAVY, padx=16, pady=12)
        header.pack(fill="x")
        tk.Label(
            header, text=f"⚡  {self._header_text}",
            font=("Segoe UI", 13, "bold"), bg=WRRL_NAVY, fg=WRRL_WHITE,
        ).pack(side="left")
        tk.Label(
            header, text=f"Season {self._year}",
            font=("Segoe UI", 10), bg=WRRL_NAVY, fg="#a0b0c0",
        ).pack(side="right", padx=(0, 4))

        tk.Frame(self, bg=WRRL_GREEN, height=3).pack(fill="x")

        body = tk.Frame(self, bg=WRRL_LIGHT, padx=24, pady=14)
        body.pack(fill="both", expand=True)

        tk.Label(
            body, text="STAGE PROGRESS",
            font=("Segoe UI", 8, "bold"), bg=WRRL_LIGHT, fg="#8a9aaa",
        ).pack(anchor="w")

        self._canvas = tk.Canvas(
            body, width=548, height=86, bg=WRRL_LIGHT, highlightthickness=0,
        )
        self._canvas.pack(pady=(4, 14), anchor="w")
        self._redraw_timeline()

        tk.Label(
            body, text="RACES",
            font=("Segoe UI", 8, "bold"), bg=WRRL_LIGHT, fg="#8a9aaa",
        ).pack(anchor="w")

        self._chips_container = tk.Frame(body, bg=WRRL_LIGHT)
        self._chips_container.pack(fill="x", pady=(4, 14), anchor="w", ipady=2)
        tk.Label(
            self._chips_container, text="Discovering race files…",
            font=("Segoe UI", 9, "italic"), bg=WRRL_LIGHT, fg="#aaaaaa",
        ).pack(anchor="w")

        tk.Frame(body, bg="#d0d8e0", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(
            body, textvariable=self._status_var,
            font=("Segoe UI", 9), bg=WRRL_LIGHT, fg=WRRL_NAVY,
            anchor="w", wraplength=548,
        ).pack(anchor="w")

        tk.Label(
            body, textvariable=self._substep_var,
            font=("Segoe UI", 9, "italic"), bg=WRRL_LIGHT, fg="#5d6e82",
            anchor="w", wraplength=548,
        ).pack(anchor="w", pady=(4, 0))

    # ── timeline canvas ───────────────────────────────────────────────────────

    def _redraw_timeline(self) -> None:
        c = self._canvas
        c.delete("all")
        cx = [90, 274, 458]
        cy, r = 36, 20

        for i in range(2):
            done = (i + 1) in self._stages_done
            c.create_line(
                cx[i] + r + 2, cy, cx[i + 1] - r - 2, cy,
                fill=self._C_DONE_LINE if done else self._C_LINE_IDLE,
                width=3, capstyle="round",
            )

        for i, label in enumerate(self._stage_labels):
            sn, x = i + 1, cx[i]
            if sn in self._stages_done:
                fill, txt, txt_c, lbl_c = self._C_DONE_NODE, "✓", "white", self._C_DONE_NODE
            elif sn == self._current_stage:
                fill  = self._C_ACTIVE_B if self._pulse_on else self._C_ACTIVE_A
                txt, txt_c, lbl_c = str(sn), "white", WRRL_NAVY
            else:
                fill, txt, txt_c, lbl_c = self._C_PEND, str(sn), self._C_PEND_TXT, self._C_PEND_TXT

            c.create_oval(x - r, cy - r, x + r, cy + r, fill=fill, outline="", width=0)
            c.create_text(x, cy, text=txt, font=("Segoe UI", 12, "bold"), fill=txt_c)
            c.create_text(x, cy + r + 14, text=label, font=("Segoe UI", 9), fill=lbl_c)

    # ── pulse animation ───────────────────────────────────────────────────────

    def _pulse(self) -> None:
        if not self._animating:
            return
        self._pulse_on = not self._pulse_on
        self._redraw_timeline()
        self.after(540, self._pulse)

    # ── public update API ─────────────────────────────────────────────────────

    def set_races(self, race_names: list[str]) -> None:
        """Populate the race chip grid."""
        for w in self._chips_container.winfo_children():
            w.destroy()
        self._race_chips = []
        if not race_names:
            tk.Label(
                self._chips_container, text="No race files found.",
                font=("Segoe UI", 9, "italic"), bg=WRRL_LIGHT, fg="#aaaaaa",
            ).pack(anchor="w")
            return
        row_f: tk.Frame | None = None
        for i, name in enumerate(race_names):
            if i % 10 == 0:
                row_f = tk.Frame(self._chips_container, bg=WRRL_LIGHT)
                row_f.pack(anchor="w", pady=(0, 4))
            chip = tk.Label(
                row_f, text=f"R{i + 1}",
                font=("Segoe UI", 8, "bold"),
                bg=self._C_CHIP_PEND, fg=self._C_CHIP_PEND_FG,
                padx=7, pady=4, relief="flat",
            )
            chip.pack(side="left", padx=(0, 5))
            chip.bind("<Enter>", lambda e, n=name, k=i: self._status_var.set(f"Race {k + 1}: {n}"))
            self._race_chips.append(chip)

    def set_stage(self, stage_num: int, label: str = "") -> None:
        """Advance to a new stage."""
        if 0 < self._current_stage < stage_num:
            self._stages_done.add(self._current_stage)
        self._current_stage = stage_num
        self._redraw_timeline()
        self._status_var.set(f"Stage {stage_num}: {label}" if label else f"Stage {stage_num}")
        if stage_num != 3:
            self._substep_var.set("")
        if stage_num == 1 and self._race_chips:
            self._animate_chips(0)

    def stage_done(self, stage_num: int) -> None:
        """Mark a stage as completed."""
        self._stages_done.add(stage_num)
        if stage_num == 1:
            self._mark_all_chips_done()
        self._redraw_timeline()

    def set_status(self, text: str) -> None:
        self._status_var.set(text)

    def set_substep(self, text: str) -> None:
        self._substep_var.set(text)

    def finish(self, success: bool) -> None:
        """Freeze animation and show final state."""
        self._animating = False
        for s in (1, 2, 3):
            self._stages_done.add(s)
        self._current_stage = 0
        self._mark_all_chips_done()
        self._redraw_timeline()
        icon = "✓" if success else "⚠"
        msg  = "Autopilot completed successfully." if success else "Autopilot completed with issues."
        self._status_var.set(f"{icon}  {msg}")
        self._substep_var.set("")

    # ── chip helpers ──────────────────────────────────────────────────────────

    def _animate_chips(self, idx: int) -> None:
        if not self._animating or not self._race_chips:
            return
        n = len(self._race_chips)
        interval = max(280, 5500 // max(n, 1))
        if idx > 0:
            prev = self._race_chips[idx - 1]
            if prev.winfo_exists():
                prev.config(bg=self._C_CHIP_DONE, fg=self._C_CHIP_DONE_FG)
        if idx < n:
            chip = self._race_chips[idx]
            if chip.winfo_exists():
                chip.config(bg=self._C_CHIP_ACT, fg=self._C_CHIP_ACT_FG)
            self.after(interval, lambda: self._animate_chips(idx + 1))

    def _mark_all_chips_done(self) -> None:
        for chip in self._race_chips:
            if chip.winfo_exists():
                chip.config(bg=self._C_CHIP_DONE, fg=self._C_CHIP_DONE_FG)

    def _on_close_attempt(self) -> None:
        messagebox.showinfo(
            "Autopilot Running",
            "Autopilot is still running. Please wait for it to complete.",
            parent=self,
        )


# ──────────────────────────────────────────────────────────────────────────────
# Main dashboard window
# ──────────────────────────────────────────────────────────────────────────────

class LeagueScorerDashboard(tk.Tk):
    """Professional dashboard for WRRL League AI."""

    def __init__(self) -> None:
        super().__init__()
        self.title("WRRL League AI")
        self.geometry("900x700")
        self.minsize(800, 600)
        self.resizable(True, True)
        self.configure(bg=WRRL_NAVY)

        # Load both logos
        self.shield_img = None
        self.waa_img = None
        self._load_logos()

        # Loaded events schedule (path stored in session_config)
        self._events_schedule = None

        self._build_header()
        self._build_main_content()
        self._build_footer()
        self._restore_events_schedule()

        # Start maximised on Windows; fall back to large geometry on other OS
        try:
            self.state('zoomed')
        except tk.TclError:
            self.geometry("1200x800")
            x = (self.winfo_screenwidth() // 2) - 600
            y = (self.winfo_screenheight() // 2) - 400
            self.geometry(f"+{x}+{y}")

    # ── logo loading ──────────────────────────────────────────────────────────

    def _load_logos(self) -> None:
        """Load and resize both header logos."""
        try:
            images_dir = Path(__file__).parent.parent / "images"

            # WRRL shield — large, left side
            shield_path = images_dir / "WRRL shield concept.png"
            if shield_path.exists():
                img = Image.open(shield_path)
                img.thumbnail((120, 120), Image.Resampling.LANCZOS)
                self.shield_img = ImageTk.PhotoImage(img)

            # WAA / wide logo — small, top right
            waa_path = images_dir / "WRRL_logo-629x400.png"
            if waa_path.exists():
                img = Image.open(waa_path)
                img.thumbnail((100, 60), Image.Resampling.LANCZOS)
                self.waa_img = ImageTk.PhotoImage(img)
        except ImportError:
            pass  # PIL not available

    # ── header section ────────────────────────────────────────────────────────

    def _build_header(self) -> None:
        """Build the top header with logos and title."""
        header = tk.Frame(self, bg=WRRL_NAVY, height=140)
        header.pack(side="top", fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        # Right-side stack for WAA logo + subtle help action
        right_stack = tk.Frame(header, bg=WRRL_NAVY)
        right_stack.pack(side="right", padx=20, pady=10, anchor="n")

        # WAA logo — small, top right
        if self.waa_img:
            waa_label = tk.Label(right_stack, image=self.waa_img, bg=WRRL_NAVY)
            waa_label.pack(anchor="e")

        # Minimal help action under WAA logo
        help_btn = tk.Button(
            right_stack,
            text="Help",
            command=self._on_help,
            font=("Segoe UI", 9),
            bg=WRRL_NAVY,
            fg="#a0b0c0",
            activebackground=WRRL_NAVY,
            activeforeground=WRRL_WHITE,
            relief="flat",
            bd=0,
            padx=2,
            pady=0,
            cursor="hand2",
            highlightthickness=0,
        )
        help_btn.pack(anchor="e", pady=(6, 0))

        # WRRL shield — large, left side
        if self.shield_img:
            shield_label = tk.Label(header, image=self.shield_img, bg=WRRL_NAVY)
            shield_label.pack(side="left", padx=20, pady=10)

        # Title and subtitle in the centre
        title_frame = tk.Frame(header, bg=WRRL_NAVY)
        title_frame.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        title = tk.Label(
            title_frame,
            text="WRRL League AI",
            font=("Segoe UI", 32, "bold"),
            bg=WRRL_NAVY,
            fg=WRRL_WHITE,
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            title_frame,
            text="WRRL League AI for Wiltshire Road and Running League operations",
            font=("Segoe UI", 11),
            bg=WRRL_NAVY,
            fg="#a0b0c0",
        )
        subtitle.pack(anchor="w")

        # Green accent bar
        accent = tk.Frame(self, bg=WRRL_GREEN, height=4)
        accent.pack(side="top", fill="x")

    # ── main content section ──────────────────────────────────────────────────

    def _build_main_content(self) -> None:
        """Build the main content area with config panel and execution options."""
        content = tk.Frame(self, bg=WRRL_LIGHT)
        content.pack(side="top", fill="both", expand=True, padx=20, pady=(12, 20))

        self._build_config_panel(content)

        # Page container — home view or scorer view swap here
        self._page_container = tk.Frame(content, bg=WRRL_LIGHT)
        self._page_container.pack(fill="both", expand=True)

        self._build_home_page()

    def _build_home_page(self) -> None:
        """Build the Quick Actions home page inside the page container."""
        self._home_frame = tk.Frame(self._page_container, bg=WRRL_LIGHT)
        self._home_frame.pack(fill="both", expand=True)

        # Section title
        section_title = tk.Label(
            self._home_frame,
            text="Quick Actions",
            font=("Segoe UI", 14, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        )
        section_title.pack(anchor="w", pady=(12, 8))

        # Button grid
        button_frame = tk.Frame(self._home_frame, bg=WRRL_LIGHT)
        button_frame.pack(fill="both", expand=True)

        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        proc_lbl = tk.Label(
            button_frame,
            text="Processing",
            font=("Segoe UI", 11, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        )
        proc_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=(0, 2))

        view_lbl = tk.Label(
            button_frame,
            text="Views",
            font=("Segoe UI", 11, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        )
        view_lbl.grid(row=0, column=1, sticky="w", padx=10, pady=(0, 2))



        # Three-column main panel layout, no headings
        for col in range(3):
            button_frame.grid_columnconfigure(col, weight=1)

        # Column 1
        # Keep a reference to the Run Autopilot card so we can update its
        # visual state when data are dirty (RAES/autopilot dirty flag).
        self._run_autopilot_card = self._create_action_button(
            button_frame, "Run Autopilot", "Run audit, safe auto-fixes, and staged checks", self._on_run_autopilot, 0, 0, tone="primary"
        )
        self._create_action_button(
            button_frame, "Publish Results", "Publish final results from audited files (includes PDF packs)", self._on_publish_results, 1, 0, tone="primary"
        )
        self._create_action_button(
            button_frame, "⬇ Fetch Results", "Download results from Race Roster into this season", self._on_import_raceroster, 2, 0, tone="primary"
        )

        # Column 2
        self._create_action_button(
            button_frame, "📝 View Autopilot Report", "Open latest autopilot summary report", self._on_view_autopilot_report, 0, 1, tone="secondary"
        )
        self._create_action_button(
            button_frame, "📊 View Results", "Open generated standings and reports", self._on_view_results, 1, 1, tone="secondary"
        )
        self._create_action_button(
            button_frame, "📋 View Events", "Browse loaded events schedule", self._on_view_events, 2, 1, tone="secondary"
        )

        # Column 3
        # Replace single 'Manual Corrections' card with two compact actions: Classic and RAES
        btn_frame_cell = tk.Frame(button_frame, bg=WRRL_LIGHT)
        btn_frame_cell.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")
        button_frame.grid_rowconfigure(0, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)

        # Card-style container to hold two buttons side-by-side
        card = tk.Frame(
            btn_frame_cell,
            bg="#ffffff",
            cursor="hand2",
            highlightthickness=1,
            highlightbackground="#d4dce4",
            highlightcolor="#d4dce4",
            padx=8,
            pady=8,
        )
        card.pack(fill="both", expand=True)

        inner = tk.Frame(card, bg="#ffffff")
        inner.pack(fill="both", expand=True)

        # Classic manual corrections retired — only RAES remains

        # Data Corrections (RAES) action card
        self._create_action_button(
            button_frame,
            "Data Corrections (RAES)",
            "Review and apply runner-level corrections using the RAES two-pane editor.",
            self._on_review_raes,
            0,
            2,
            tone="secondary",
        )
        self._create_action_button(
            button_frame, "Compare Raw vs Archive", "Inspect line-by-line changes against the raw-data archive", self._on_compare_raw_archive, 1, 2, tone="secondary"
        )
        self._create_action_button(
            button_frame, "🔎 Runner/Club Enquiry", "Search published results by runner or club", self._on_view_runner_history, 2, 2, tone="secondary"
        )


        # Add small, dark, subtle buttons to the bottom, styled like the header
        self._add_bottom_action_buttons()
    def _add_bottom_action_buttons(self):
        """Add small, dark, subtle action buttons to the bottom area, styled like the header."""
        bottom_frame = tk.Frame(self._home_frame, bg=WRRL_NAVY, pady=10)
        bottom_frame.pack(side="bottom", fill="x")

        btn_cfg = {
            "font": ("Segoe UI", 9),
            "bg": WRRL_NAVY,
            "fg": "#a0b0c0",
            "activebackground": WRRL_GREEN,
            "activeforeground": WRRL_WHITE,
            "relief": "flat",
            "bd": 0,
            "padx": 4,
            "pady": 2,
            "cursor": "hand2",
            "highlightthickness": 0,
            "width": 14,
        }

        # Inner frame for tight grouping
        inner = tk.Frame(bottom_frame, bg=WRRL_NAVY)
        inner.pack(anchor="w", pady=2, padx=10)

        btn1 = tk.Button(
            inner,
            text="⚙️ Settings",
            command=self._on_settings,
            **btn_cfg,
        )
        btn1.pack(side="left")

        sep1 = tk.Frame(inner, bg="#b0c4de", width=4, height=26)
        sep1.pack(side="left", padx=2, pady=0)

        btn2 = tk.Button(
            inner,
            text="▶ Classic Scorer",
            command=self._on_run_scorer,
            **btn_cfg,
        )
        btn2.pack(side="left")

        sep2 = tk.Frame(inner, bg="#b0c4de", width=4, height=26)
        sep2.pack(side="left", padx=2, pady=0)

        btn3 = tk.Button(
            inner,
            text="Publish Provisional",
            command=self._on_run_provisional_fast_track,
            **btn_cfg,
        )
        btn3.pack(side="left")

    # ── config panel ──────────────────────────────────────────────────────────

    def _build_config_panel(self, parent: tk.Frame) -> None:
        """Build the season configuration panel."""
        panel = tk.Frame(parent, bg=WRRL_NAVY, padx=12, pady=10)
        panel.pack(fill="x", pady=(0, 4))

        # ── Season picker ────────────────────────────────────────────────────
        season_frame = tk.Frame(panel, bg=WRRL_NAVY)
        season_frame.pack(anchor="w")

        tk.Label(
            season_frame, text="Season",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_NAVY, fg="#a0b0c0",
        ).pack()

        picker = tk.Frame(season_frame, bg=WRRL_NAVY)
        picker.pack(pady=(2, 0))

        btn_cfg = dict(
            font=("Segoe UI", 14, "bold"),
            bg=WRRL_NAVY, fg=WRRL_GREEN,
            relief="flat", bd=0, padx=4, pady=0, cursor="hand2",
            activebackground=WRRL_NAVY, activeforeground="#6fc98a",
        )
        tk.Button(picker, text="\u25c4", command=self._on_year_prev, **btn_cfg).pack(side="left")

        self._year_label = tk.Label(
            picker,
            text=str(session_config.year),
            font=("Segoe UI", 22, "bold"),
            bg=WRRL_NAVY, fg=WRRL_WHITE,
            width=5, anchor="center",
        )
        self._year_label.pack(side="left", padx=6)

        tk.Button(picker, text="\u25ba", command=self._on_year_next, **btn_cfg).pack(side="left")

        freshness = tk.Frame(panel, bg=WRRL_NAVY)
        freshness.pack(anchor="w", pady=(8, 0))
        self._freshness_light = tk.Canvas(
            freshness,
            width=16,
            height=16,
            bg=WRRL_NAVY,
            highlightthickness=0,
            bd=0,
        )
        self._freshness_light.pack(side="left")
        self._freshness_dot = self._freshness_light.create_oval(2, 2, 14, 14, fill="#cc4444", outline="")
        self._freshness_var = tk.StringVar(value="Data status unknown")
        tk.Label(
            freshness,
            textvariable=self._freshness_var,
            font=("Segoe UI", 9),
            bg=WRRL_NAVY,
            fg="#d7dfeb",
            anchor="w",
        ).pack(side="left", padx=(6, 0))

        self._refresh_config_panel()

    def _refresh_config_panel(self) -> None:
        """Update config bar indicators to reflect current session_config state."""
        if hasattr(self, "_year_label"):
            self._year_label.config(text=str(session_config.year))
        if hasattr(self, "_freshness_var") and hasattr(self, "_freshness_light"):
            is_current, detail = self._compute_data_freshness()
            colour = WRRL_GREEN if is_current else "#d94f4f"
            self._freshness_light.itemconfig(self._freshness_dot, fill=colour)
            self._freshness_var.set(detail)
            # Also update Run Autopilot action card to visually reflect the same
            # freshness/dirty state so the prominent action is consistent.
            try:
                if hasattr(self, "_run_autopilot_card") and self._run_autopilot_card.winfo_exists():
                    # Find the title and subtitle labels inside the card
                    children = self._run_autopilot_card.winfo_children()
                    title_lbl = children[0] if len(children) > 0 else None
                    subtitle_lbl = children[1] if len(children) > 1 else None
                    if is_current:
                        self._run_autopilot_card.config(bg=WRRL_GREEN)
                        if title_lbl is not None:
                            title_lbl.config(bg=WRRL_GREEN, fg=WRRL_WHITE)
                        if subtitle_lbl is not None:
                            subtitle_lbl.config(bg=WRRL_GREEN, fg="#d9efe2")
                    else:
                        self._run_autopilot_card.config(bg=WRRL_AMBER)
                        if title_lbl is not None:
                            title_lbl.config(bg=WRRL_AMBER, fg=WRRL_NAVY)
                        if subtitle_lbl is not None:
                            subtitle_lbl.config(bg=WRRL_AMBER, fg="#6f5a10")
            except Exception:
                pass

    def _compute_data_freshness(self) -> tuple[bool, str]:
        if not session_config.input_dir:
            return False, "No input folder configured"

        # If RAES has written a dirty marker, treat data as stale until autopilot runs
        out = session_config.output_dir
        if out is not None:
            raes_dirty = Path(out) / "raes" / "dirty"
            if raes_dirty.exists():
                return False, "RAES edits pending — run Autopilot"

        raw_data_dir = session_config.raw_data_dir
        audited_dir = session_config.audited_dir
        if not raw_data_dir or not audited_dir:
            return False, "Input folders are not available"

        raw_files = discover_race_files(raw_data_dir, excluded_names=race_discovery_exclusions())
        audited_files = discover_race_files(audited_dir, excluded_names=race_discovery_exclusions())

        if not raw_files:
            return True, "No raw race files yet"

        stale: list[int] = []
        for race_num, raw_path in raw_files.items():
            audited_path = audited_files.get(race_num)
            if audited_path is None:
                stale.append(race_num)
                continue
            if raw_path.stat().st_mtime > audited_path.stat().st_mtime:
                stale.append(race_num)

        if stale:
            return False, f"Audit required for race(s): {', '.join(str(num) for num in stale[:6])}"
        return True, "All audited files are current"

    def _on_year_prev(self) -> None:
        """Step the season year back by one."""
        years = session_config.available_years()
        idx = years.index(session_config.year) if session_config.year in years else 0
        if idx > 0:
            session_config.year = years[idx - 1]
            self._restore_events_schedule()

    def _on_year_next(self) -> None:
        """Step the season year forward by one."""
        years = session_config.available_years()
        idx = years.index(session_config.year) if session_config.year in years else 0
        if idx < len(years) - 1:
            session_config.year = years[idx + 1]
            self._restore_events_schedule()



    # ── action buttons ──────────────────────────────────────────────────────

    def _create_action_button(
        self,
        parent: tk.Frame,
        text: str,
        subtitle: str,
        command,
        row: int,
        col: int,
        tone: str = "secondary",
    ) -> None:
        """Create a modern action card with title, subtitle, and hover styling."""
        btn_frame = tk.Frame(parent, bg=WRRL_LIGHT)
        btn_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        parent.grid_rowconfigure(row, weight=1)
        parent.grid_columnconfigure(col, weight=1)

        if tone == "primary":
            card_bg = WRRL_GREEN
            title_fg = WRRL_WHITE
            subtitle_fg = "#d9efe2"
            hover_bg = "#24653d"
            border = "#1f5632"
        else:
            card_bg = "#ffffff"
            title_fg = WRRL_NAVY
            subtitle_fg = "#66788d"
            hover_bg = "#eef2f7"
            border = "#d4dce4"

        card = tk.Frame(
            btn_frame,
            bg=card_bg,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=border,
            padx=16,
            pady=14,
        )
        card.pack(fill="both", expand=True)

        title_lbl = tk.Label(
            card,
            text=text,
            font=("Segoe UI", 13, "bold"),
            bg=card_bg,
            fg=title_fg,
            anchor="w",
            justify="left",
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = tk.Label(
            card,
            text=subtitle,
            font=("Segoe UI", 9),
            bg=card_bg,
            fg=subtitle_fg,
            anchor="w",
            justify="left",
            wraplength=320,
        )
        subtitle_lbl.pack(anchor="w", pady=(6, 0))

        def _set_bg(bg: str) -> None:
            card.configure(bg=bg)
            title_lbl.configure(bg=bg)
            subtitle_lbl.configure(bg=bg)

        def _on_enter(_event) -> None:
            _set_bg(hover_bg)

        def _on_leave(_event) -> None:
            _set_bg(card_bg)

        for widget in (card, title_lbl, subtitle_lbl):
            widget.bind("<Button-1>", lambda _e: command())
            widget.bind("<Enter>", _on_enter)
            widget.bind("<Leave>", _on_leave)

        # Return the card so callers can keep a reference for dynamic updates.
        return card

    # ── footer section ────────────────────────────────────────────────────────

    def _build_footer(self) -> None:
        """Build the footer with version info."""
        footer = tk.Frame(self, bg=WRRL_NAVY, height=50)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        from .. import __version__
        footer_text = tk.Label(
            footer,
            text=f"© 2026 Wiltshire Athletics Assoc. | WRRL League AI v{__version__}",
            font=("Segoe UI", 9),
            bg=WRRL_NAVY,
            fg="#707080",
        )
        footer_text.pack(pady=10)

    # ── action handlers ───────────────────────────────────────────────────────

    # ── guards ────────────────────────────────────────────────────────────

    def _require_configured(self, action: str = "this action") -> bool:
        """Return True if a data root is set; otherwise show a warning."""
        if not session_config.is_configured:
            messagebox.showwarning(
                "Not Configured",
                f"Please set a Data Root folder before using {action}.\n\n"
                "Use Settings (⚙️) to configure your data paths.",
            )
            return False
        return True

    # ── action handlers ────────────────────────────────────────────────────

    def _on_run_scorer(self) -> None:
        """Show the WRRL League AI scorer panel inline within the dashboard."""
        if not self._require_configured("Run WRRL League AI"):
            return
        log_event("dashboard_open_scorer", year=session_config.year)
        session_config.ensure_dirs()
        if session_config.input_dir:
            sort_existing_input_files(session_config.input_dir)
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)
        self._home_frame.pack_forget()
        scorer = LeagueScorerApp(
            self._page_container,
            input_dir=session_config.input_dir,
            output_dir=session_config.output_dir,
            year=session_config.year,
            back_callback=self._on_scorer_back,
        )
        scorer.pack(fill="both", expand=True)
        self._scorer_frame = scorer

    def _on_scorer_back(self) -> None:
        """Return from the scorer panel to the home page."""
        if hasattr(self, "_scorer_frame"):
            self._scorer_frame.destroy()
            del self._scorer_frame
        self._home_frame.pack(fill="both", expand=True)

    def _on_run_autopilot(self) -> None:
        """Run full automation pipeline in background and report outcome."""
        if not self._require_configured("Run Autopilot"):
            return
        if session_config.data_root is None:
            messagebox.showerror("Data Root Missing", "Set Data Root before running autopilot.", parent=self)
            return

        session_config.ensure_dirs()
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)

        output_paths = ensure_output_subdirs(session_config.output_dir)
        dlg = _AutopilotProgressDialog(self, year=session_config.year)

        def _show_result(code: int, stdout_text: str, stderr_text: str) -> None:
            if dlg.winfo_exists():
                dlg.grab_release()
                dlg.destroy()
            report_path = self._resolve_autopilot_report_path()
            staged_report_path = self._resolve_staged_checks_report_path()
            review_path = None
            if report_path is not None and report_path.exists():
                review_path = report_path
            elif staged_report_path is not None and staged_report_path.exists():
                review_path = staged_report_path
            self._show_autopilot_result_dialog(success=(code == 0), review_path=review_path)

        self._run_workflow(
            script_name="run_full_autopilot.py",
            dlg=dlg,
            extra_cmd_args=[
                "--mode", "apply-safe-fixes",
                "--staged-report-dir", str(output_paths.quality_staged_checks_dir),
                "--data-quality-output-dir", str(output_paths.quality_data_dir),
            ],
            error_title="Autopilot Failed",
            show_result_fn=_show_result,
            handle_substep=True,
            finish_delay_ms=1400,
        )

    def _on_run_provisional_fast_track(self) -> None:
        """Publish provisional results without the full audit and staged checks."""
        if not self._require_configured("Publish Provisional Results"):
            return
        if session_config.data_root is None:
            messagebox.showerror("Data Root Missing", "Set Data Root before publishing provisional results.", parent=self)
            return
        if not messagebox.askyesno(
            "Publish Provisional Results",
            "This fast track skips audit review and quality checks.\n\nUse it only when you want a provisional publish from the current raw data. Continue?",
            parent=self,
        ):
            return

        session_config.ensure_dirs()
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)

        dlg = _AutopilotProgressDialog(
            self,
            year=session_config.year,
            window_title="Provisional Fast Track",
            header_text="WRRL Provisional Fast Track",
            stage_labels=["Refresh Audited", "Publish Results", "Write Summary"],
            initial_status="Initialising provisional publish...",
        )

        def _show_result(code: int, stdout_text: str, stderr_text: str) -> None:
            if dlg.winfo_exists():
                dlg.grab_release()
                dlg.destroy()
            review_path = self._resolve_provisional_fast_track_report_path()
            self._show_workflow_result_dialog(
                title="Provisional Results Complete",
                success=(code == 0),
                review_path=review_path if review_path is not None and review_path.exists() else None,
                success_headline="Provisional publish complete",
                failure_headline="Provisional publish stopped",
                success_body="The current raw data has been turned into published provisional results.\nReview the summary if you want the details.",
                failure_body="The provisional fast track did not finish cleanly.\nReview the summary for the failure details before retrying.",
                review_button_text="Review Summary",
            )

        self._run_workflow(
            script_name="run_provisional_fast_track.py",
            dlg=dlg,
            extra_cmd_args=[],
            error_title="Provisional Fast Track Failed",
            show_result_fn=_show_result,
        )

    def _on_publish_results(self) -> None:
        """Publish final results from audited files (PDF generation enabled)."""
        if not self._require_configured("Publish Results"):
            return
        if session_config.data_root is None:
            messagebox.showerror("Data Root Missing", "Set Data Root before publishing results.", parent=self)
            return
        if not messagebox.askyesno(
            "Publish Results",
            "Publish final results from audited files now?\n\nThis run writes the full publish pack including PDF outputs.",
            parent=self,
        ):
            return

        session_config.ensure_dirs()
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)

        dlg = _AutopilotProgressDialog(
            self,
            year=session_config.year,
            window_title="Publish Results",
            header_text="WRRL Final Publish",
            stage_labels=["Load Audited", "Publish Results", "Write Summary"],
            initial_status="Initialising final publish...",
        )

        def _show_result(code: int, stdout_text: str, stderr_text: str) -> None:
            if dlg.winfo_exists():
                dlg.grab_release()
                dlg.destroy()
            review_path = self._resolve_publish_results_report_path()
            self._show_workflow_result_dialog(
                title="Publish Results Complete",
                success=(code == 0),
                review_path=review_path if review_path is not None and review_path.exists() else None,
                success_headline="Final publish complete",
                failure_headline="Final publish stopped",
                success_body="Published final results from audited files, including PDF outputs.",
                failure_body="The final publish did not finish cleanly.\nReview the summary for details before retrying.",
                review_button_text="Review Summary",
            )

        self._run_workflow(
            script_name="run_publish_results.py",
            dlg=dlg,
            extra_cmd_args=[],
            error_title="Publish Results Failed",
            show_result_fn=_show_result,
        )

    def _run_workflow(
        self,
        *,
        script_name: str,
        dlg: "_AutopilotProgressDialog",
        extra_cmd_args: list[str],
        error_title: str,
        show_result_fn: Callable[[int, str, str], None],
        handle_substep: bool = False,
        finish_delay_ms: int = 1200,
    ) -> None:
        """Shared subprocess runner for all three workflow handlers."""
        result_queue: queue.Queue = queue.Queue()

        def _worker() -> None:
            script_path = str(Path(__file__).resolve().parents[2] / "scripts" / script_name)
            output_paths = ensure_output_subdirs(session_config.output_dir)
            report_base = output_paths.autopilot_runs_dir
            cmd = [
                sys.executable, "-u", script_path,
                "--year", str(session_config.year),
                "--data-root", str(session_config.data_root),
                "--report-dir", str(report_base),
                *extra_cmd_args,
            ]
            # Build a clean env snapshot so GUI-specific env mutations don't
            # bleed into the worker process, and exclude interpreter startup
            # hooks that are irrelevant in a non-interactive subprocess.
            _child_env = {k: v for k, v in os.environ.items() if k != "PYTHONSTARTUP"}
            try:
                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, encoding="utf-8", errors="replace",
                    env=_child_env,
                )
                stderr_lines: list[str] = []

                def _drain_stderr() -> None:
                    if proc.stderr is None:
                        return
                    for ln in proc.stderr:
                        stderr_lines.append(ln)

                threading.Thread(target=_drain_stderr, daemon=True).start()

                stdout_lines: list[str] = []
                if proc.stdout is not None:
                    for raw in proc.stdout:
                        line = raw.rstrip("\n")
                        stdout_lines.append(line)
                        if line.startswith("PROGRESS:"):
                            result_queue.put(("progress", line))

                proc.wait()
                result_queue.put((
                    "done", proc.returncode,
                    "\n".join(stdout_lines),
                    "".join(stderr_lines),
                ))
            except Exception as exc:
                result_queue.put(("error", str(exc)))

        threading.Thread(target=_worker, daemon=True).start()

        def _handle_progress(line: str) -> None:
            try:
                parts = line.split(":", 3)
                if len(parts) < 3 or not dlg.winfo_exists():
                    return
                kind = parts[1]
                rest = parts[2:]
                if kind == "RACES":
                    names = [n for n in rest[0].split("|") if n]
                    dlg.set_races(names)
                    dlg.set_status(f"Discovered {len(names)} race file{'s' if len(names) != 1 else ''}")
                elif kind == "STAGE":
                    try:
                        sn = int(rest[0])
                    except (ValueError, IndexError):
                        return
                    dlg.set_stage(sn, rest[1] if len(rest) > 1 else "")
                elif kind == "STAGE_DONE":
                    try:
                        sn = int(rest[0])
                    except (ValueError, IndexError):
                        return
                    dlg.stage_done(sn)
                elif handle_substep and kind == "SUBSTEP":
                    try:
                        sn = int(rest[0])
                    except (ValueError, IndexError):
                        return
                    if sn == 3:
                        dlg.set_substep(f"• {rest[1] if len(rest) > 1 else ''}")
            except tk.TclError:
                return

        def _poll() -> None:
            if not dlg.winfo_exists():
                return
            while True:
                try:
                    payload = result_queue.get_nowait()
                except queue.Empty:
                    self.after(80, _poll)
                    return
                if payload[0] == "progress":
                    _handle_progress(payload[1])
                elif payload[0] == "done":
                    _, code, stdout_text, stderr_text = payload
                    dlg.finish(code == 0)
                    self.after(finish_delay_ms, lambda c=code, o=stdout_text, e=stderr_text: show_result_fn(c, o, e))
                    return
                elif payload[0] == "error":
                    if dlg.winfo_exists():
                        dlg.grab_release()
                        dlg.destroy()
                    messagebox.showerror(error_title, payload[1], parent=self)
                    return
                else:
                    continue

        self.after(80, _poll)

    def _resolve_autopilot_report_path(self) -> Path | None:
        """Resolve autopilot report location in outputs/autopilot/runs."""
        if session_config.output_dir is None:
            return None
        return (
            build_output_paths(session_config.output_dir)
            .autopilot_runs_dir
            / f"year-{session_config.year}"
            / "autopilot_report.md"
        )

    def _resolve_staged_checks_report_path(self) -> Path | None:
        """Resolve staged-check markdown report in outputs/quality/staged-checks."""
        if session_config.output_dir is None:
            return None
        return build_output_paths(session_config.output_dir).quality_staged_checks_dir / "staged_checks_report.md"

    def _resolve_provisional_fast_track_report_path(self) -> Path | None:
        if session_config.output_dir is None:
            return None
        return (
            build_output_paths(session_config.output_dir)
            .autopilot_runs_dir
            / f"year-{session_config.year}"
            / "provisional_fast_track.md"
        )

    def _resolve_publish_results_report_path(self) -> Path | None:
        if session_config.output_dir is None:
            return None
        return (
            build_output_paths(session_config.output_dir)
            .autopilot_runs_dir
            / f"year-{session_config.year}"
            / "publish_results.md"
        )

    def _show_workflow_result_dialog(
        self,
        *,
        title: str,
        success: bool,
        review_path: Path | None,
        success_headline: str,
        failure_headline: str,
        success_body: str,
        failure_body: str,
        review_button_text: str = "Review Messages",
    ) -> None:
        """Show a friendly completion dialog with an optional review action."""
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=WRRL_LIGHT)

        frame = tk.Frame(dialog, bg=WRRL_LIGHT, padx=18, pady=16)
        frame.pack(fill="both", expand=True)

        headline = success_headline if success else failure_headline
        body = success_body if success else failure_body
        if review_path is None:
            body += "\n\nNo message file is available yet."

        tk.Label(
            frame,
            text=headline,
            font=("Segoe UI", 13, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
            anchor="w",
        ).pack(anchor="w")

        tk.Label(
            frame,
            text=body,
            font=("Segoe UI", 10),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
            justify="left",
            anchor="w",
            wraplength=430,
        ).pack(anchor="w", pady=(8, 14))

        btn_row = tk.Frame(frame, bg=WRRL_LIGHT)
        btn_row.pack(fill="x")

        def _on_review_messages() -> None:
            if review_path is None:
                return
            try:
                self._open_file_in_system(review_path)
            except Exception as exc:
                messagebox.showerror("Open Failed", str(exc), parent=dialog)

        review_btn = tk.Button(
            btn_row,
            text=review_button_text,
            command=_on_review_messages,
            font=("Segoe UI", 10, "bold"),
            bg="#2d7a4a",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground="#ffffff",
            state="normal" if review_path is not None else "disabled",
        )
        review_btn.pack(side="left")

        close_btn = ttk.Button(btn_row, text="Close", command=dialog.destroy)
        close_btn.pack(side="right")

    def _show_autopilot_result_dialog(self, *, success: bool, review_path: Path | None) -> None:
        """Show a friendly autopilot completion dialog with optional review action."""
        self._show_workflow_result_dialog(
            title="Autopilot Complete",
            success=success,
            review_path=review_path,
            success_headline="All done",
            failure_headline="Autopilot finished",
            success_body="Everything ran successfully.\nYou can review the messages if you want a summary.",
            failure_body="Some items need your attention.\nPlease review messages for a clear summary of what to do next.",
        )

    def _open_file_in_system(self, path: Path) -> None:
        if sys.platform == "win32":
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)

    def _on_view_autopilot_report(self) -> None:
        """Open the latest autopilot markdown report in the default editor/app."""
        if not self._require_configured("View Autopilot Report"):
            return
        if session_config.output_dir is None:
            messagebox.showerror("Output Not Configured", "Output folder is not configured.", parent=self)
            return

        report_path = self._resolve_autopilot_report_path()
        if report_path is None or not report_path.exists():
            messagebox.showwarning(
                "Report Not Found",
                f"No autopilot report found for season {session_config.year}.\nRun Autopilot first.",
                parent=self,
            )
            return

        try:
            self._open_file_in_system(report_path)
        except Exception as exc:
            messagebox.showerror("Open Failed", str(exc), parent=self)

    def _on_load_events(self) -> None:
        """Browse for and load an events XLSX file."""
        if not self._require_configured("Load Events"):
            return
        session_config.ensure_dirs()
        if session_config.input_dir:
            sort_existing_input_files(session_config.input_dir)
        initial_path = session_config.events_path
        if initial_path and initial_path.parent.exists():
            initial_dir = str(initial_path.parent)
        else:
            initial_dir = str(session_config.control_dir) if session_config.control_dir else "/"
        path_str = filedialog.askopenfilename(
            title="Select Events Spreadsheet",
            initialdir=initial_dir,
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path_str:
            return
        try:
            events_path = Path(path_str)
            self._events_schedule = load_events(events_path)
            session_config.events_path = events_path
            self._refresh_config_panel()
            count = len(self._events_schedule.events)
            messagebox.showinfo(
                "Events Loaded",
                f"Loaded {count} event(s) from:\n{events_path.name}",
            )
        except Exception as exc:
            messagebox.showerror("Load Failed", str(exc))
            self._events_schedule = None
            session_config.events_path = None
            self._refresh_config_panel()

    def _restore_events_schedule(self) -> bool:
        """Restore the remembered events file for the active season if possible."""
        events_path = session_config.events_path
        if (not events_path or not events_path.exists()) and session_config.control_dir is not None:
            control_dir = session_config.control_dir
            default_path = control_dir / "wrrl_events.xlsx"
            if default_path.exists():
                events_path = default_path
                session_config.events_path = default_path
            else:
                candidates = sorted(control_dir.glob("*event*.xlsx"), key=lambda p: p.name.lower())
                if candidates:
                    events_path = candidates[0]
                    session_config.events_path = events_path

        if not events_path or not events_path.exists():
            self._events_schedule = None
            self._refresh_config_panel()
            return False

        try:
            self._events_schedule = load_events(events_path)
        except Exception:
            self._events_schedule = None
            self._refresh_config_panel()
            return False

        self._refresh_config_panel()
        return True

    def _on_view_events(self) -> None:
        """Show the events viewer panel inline within the dashboard."""
        if not self._require_configured("View Events"):
            return
        if self._events_schedule is None:
            if not self._restore_events_schedule():
                messagebox.showwarning(
                    "Events Not Found",
                    "No events spreadsheet could be found in inputs/control.\nExpected default: wrrl_events.xlsx",
                    parent=self,
                )
                return
        images_dir = Path(__file__).parent.parent / "images"
        self._home_frame.pack_forget()
        panel = EventsViewerPanel(
            self._page_container,
            self._events_schedule,
            year=session_config.year,
            images_dir=images_dir,
            output_dir=session_config.output_dir,
            on_return_dashboard=self.show_home_panel
        )
        panel.pack(fill="both", expand=True)

    def show_home_panel(self):
        """Show the dashboard home panel."""
        # Remove all children from page container except _home_frame
        for child in self._page_container.winfo_children():
            if child is not self._home_frame:
                child.destroy()
        self._home_frame.pack(fill="both", expand=True)

    def _on_view_results(self) -> None:
        """Show the results viewer panel inline within the dashboard."""
        if not self._require_configured("View Results"):
            return
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)
        self._home_frame.pack_forget()
        from .results_viewer import ResultsViewerPanel
        panel = ResultsViewerPanel(self._page_container)
        panel.pack(fill="both", expand=True)
        self._results_panel = panel
        def on_close():
            if hasattr(self, "_results_panel"):
                self._results_panel.destroy()
                del self._results_panel
            self._home_frame.pack(fill="both", expand=True)
        # Always-visible return control (top-right overlay)
        close_btn = tk.Button(
            panel,
            text="🏠 Dashboard",
            command=on_close,
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_GREEN,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_GREEN,
        )
        close_btn.place(relx=1.0, x=-12, y=10, anchor="ne")

    def _on_review_manual_corrections(self) -> None:
        # Legacy manual corrections retired; this action is removed.
        messagebox.showinfo("Retired", "The classic manual corrections flow has been retired. Use RAES instead.", parent=self)

    def _on_review_raes(self) -> None:
        """Open the RAES manual correction panel."""
        if not self._require_configured("RAES Manual Corrections"):
            return
        log_event("dashboard_open_manual_corrections_raes", year=session_config.year)
        # Open the RAESPanel inline (two-pane view)
        self._home_frame.pack_forget()
        try:
            from ..raes.raes_panel import RAESPanel

            panel = RAESPanel(self._page_container, back_callback=self.show_home_panel)
            panel.pack(fill="both", expand=True)
            self._raes_panel = panel
        except Exception as exc:
            messagebox.showerror("RAES Error", f"Failed to open RAES panel: {exc}", parent=self)

    def _on_manual_review_back(self) -> None:
        # No-op for retired manual review panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_settings(self) -> None:
        """Show the settings panel inline within the dashboard."""
        if not self._require_configured("Settings"):
            return
        self._home_frame.pack_forget()
        from .settings_dialog import SettingsPanel

        def on_close():
            if hasattr(self, "_settings_panel"):
                self._settings_panel.destroy()
                del self._settings_panel
            self._home_frame.pack(fill="both", expand=True)

        panel = SettingsPanel(
            self._page_container,
            on_close=on_close,
            on_open_logs=self._on_open_logs_from_settings,
        )
        panel.pack(fill="both", expand=True)
        self._settings_panel = panel

    def _on_open_logs_from_settings(self) -> None:
        if not hasattr(self, "_settings_panel"):
            return
        self._settings_panel.pack_forget()
        try:
            from .log_viewer import LogViewerPanel

            panel = LogViewerPanel(
                self._page_container,
                back_callback=self._on_logs_back_to_settings,
                dashboard_callback=self._on_logs_back_to_dashboard,
            )
            panel.pack(fill="both", expand=True)
            self._log_viewer_panel = panel
        except Exception as exc:
            self._settings_panel.pack(fill="both", expand=True)
            messagebox.showerror("Log Viewer Error", str(exc), parent=self)

    def _on_logs_back_to_settings(self) -> None:
        if hasattr(self, "_log_viewer_panel"):
            self._log_viewer_panel.destroy()
            del self._log_viewer_panel
        if hasattr(self, "_settings_panel"):
            self._settings_panel.pack(fill="both", expand=True)

    def _on_logs_back_to_dashboard(self) -> None:
        if hasattr(self, "_log_viewer_panel"):
            self._log_viewer_panel.destroy()
            del self._log_viewer_panel
        if hasattr(self, "_settings_panel"):
            self._settings_panel.destroy()
            del self._settings_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_audit_runners(self) -> None:
        """Show the audit runner panel inline within the dashboard."""
        if not self._require_configured("Audit Runners"):
            return
        session_config.ensure_dirs()
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)
        self._home_frame.pack_forget()
        from .audit_gui import LeagueAuditApp
        panel = LeagueAuditApp(
            self._page_container,
            input_dir=session_config.input_dir,
            output_dir=session_config.output_dir,
            year=session_config.year,
            back_callback=self._on_audit_back,
            completion_callback=self._on_audit_complete_view,
        )
        panel.pack(fill="both", expand=True)
        self._audit_panel = panel

    def _on_view_audit(self, preferred_workbook=None, return_to_audit: bool = False) -> None:
        """Show the audit viewer panel inline within the dashboard."""
        if not self._require_configured("View Audit"):
            return
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)
        if return_to_audit and hasattr(self, "_audit_panel"):
            self._audit_panel.pack_forget()
        else:
            self._home_frame.pack_forget()
        from .audit_viewer import AuditViewerPanel
        panel = AuditViewerPanel(self._page_container, preferred_workbook=preferred_workbook)
        panel.pack(fill="both", expand=True)
        self._audit_view_panel = panel

        def on_close():
            if hasattr(self, "_audit_view_panel"):
                self._audit_view_panel.destroy()
                del self._audit_view_panel
            if return_to_audit and hasattr(self, "_audit_panel"):
                self._audit_panel.pack(fill="both", expand=True)
            else:
                self._home_frame.pack(fill="both", expand=True)

        close_text = "\u25c4 Audit" if return_to_audit else "🏠 Dashboard",
        close_btn = tk.Button(
            panel,
            text=close_text,
            command=on_close,
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
        )
        close_btn.place(relx=1.0, x=-12, y=10, anchor="ne")

    def _on_audit_back(self) -> None:
        if hasattr(self, "_audit_panel"):
            self._audit_panel.destroy()
            del self._audit_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_view_runner_history_back(self) -> None:
        if hasattr(self, "_runner_history_panel"):
            self._runner_history_panel.destroy()
            del self._runner_history_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_compare_raw_archive_back(self) -> None:
        if hasattr(self, "_raw_archive_diff_panel"):
            self._raw_archive_diff_panel.destroy()
            del self._raw_archive_diff_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_view_runner_history(self) -> None:
        if not self._require_configured("Runner/Club Enquiry"):
            return
        self._home_frame.pack_forget()
        from .runner_history_viewer import RunnerHistoryPanel

        panel = RunnerHistoryPanel(
            self._page_container,
            back_callback=self._on_view_runner_history_back,
        )
        panel.pack(fill="both", expand=True)
        self._runner_history_panel = panel

    def _on_compare_raw_archive(self) -> None:
        if not self._require_configured("Compare Raw vs Archive"):
            return
        self._home_frame.pack_forget()
        from .raw_archive_diff_viewer import RawArchiveDiffPanel

        panel = RawArchiveDiffPanel(
            self._page_container,
            back_callback=self._on_compare_raw_archive_back,
        )
        panel.pack(fill="both", expand=True)
        self._raw_archive_diff_panel = panel

    def _on_issue_review(self) -> None:
        """Show the Issue Review panel inline within the dashboard."""
        if not self._require_configured("Review Issues"):
            return
        self._home_frame.pack_forget()
        from .issue_reviewer import IssueReviewPanel
        panel = IssueReviewPanel(
            self._page_container,
            back_callback=self._on_issue_review_back,
            view_runner_callback=self._on_issue_review_open_runner,
        )
        panel.pack(fill="both", expand=True)
        self._issue_review_panel = panel

    def _on_issue_review_back(self) -> None:
        if hasattr(self, "_issue_review_panel"):
            self._issue_review_panel.destroy()
            del self._issue_review_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_issue_review_open_runner(self, runner_name: str) -> None:
        """Close Issue Review panel and open Runner History pre-loaded for runner_name."""
        if hasattr(self, "_issue_review_panel"):
            self._issue_review_panel.destroy()
            del self._issue_review_panel
        from .runner_history_viewer import RunnerHistoryPanel
        panel = RunnerHistoryPanel(
            self._page_container,
            back_callback=self._on_view_runner_history_back,
            initial_runner=runner_name,
        )
        panel.pack(fill="both", expand=True)
        self._runner_history_panel = panel

    def _on_audit_complete_view(self, preferred_workbook=None) -> None:
        self._on_view_audit(preferred_workbook=preferred_workbook, return_to_audit=True)

    def _on_help(self) -> None:
        """Show help information."""
        docs_dir = Path(__file__).resolve().parents[2] / "documents"
        dependencies_doc = docs_dir / "dependencies.md"
        ops_doc = docs_dir / "operational_dependencies.md"

        help_text = (
            "WRRL League AI Help\n\n"
            "• Settings (⚙️): Configure data paths and league scoring parameters.\n"
            "  Set your Data Root once; folders are created as: {root}/{year}/inputs and {root}/{year}/outputs\n"
            "  Inputs are structured as: raw_data, series, control, audited, raw_data_archive\n\n"
            "• Season: Select the current league year.\n\n"
            "• Events are auto-loaded from inputs/control (default filename: wrrl_events.xlsx).\n"
            "• View Events: Browse the loaded events schedule.\n\n"
            "• Import Race Roster: Paste a Race Roster URL and save results directly\n"
            "  into inputs/raw_data.\n\n"
            "• Run Autopilot: Execute audit, safe auto-fixes, and staged checks.\n"
            "  Autopilot suppresses PDF generation to keep turnaround fast.\n"
            "• Publish Results: Build final publish outputs from audited files (includes PDFs).\n"
            "• View Autopilot Report: Open the latest automation summary.\n"
            "• Run WRRL League AI: Execute the classic scoring pipeline manually.\n"
            "• View Results: Browse generated results.\n\n"
            "Operational dependencies docs:\n"
            f"• {dependencies_doc}\n"
            f"• {ops_doc}\n\n"
            "Key notes:\n"
            "• PDF output requires Microsoft Word (docx2pdf).\n"
            "• Browser automation features may require: playwright install chromium"
        )
        messagebox.showinfo("Help", help_text)

    def _on_import_raceroster(self) -> None:
        """Import Race Roster results directly into the active season input folder."""
        if not self._require_configured("Import Race Roster"):
            return

        session_config.ensure_dirs()
        if session_config.input_dir:
            sort_existing_input_files(session_config.input_dir)
        input_dir = session_config.raw_data_dir
        if not input_dir or not input_dir.exists():
            messagebox.showerror("Input not found", f"Raw data folder does not exist:\n{input_dir}")
            return

        request = prompt_race_import_request(self)
        if request is None:
            return

        log_event(
            "dashboard_import_raceroster_requested",
            year=session_config.year,
            race_number=request.race_number,
            race_url=request.race_url,
            has_sporthive_hint=request.sporthive_race_hint is not None,
        )

        self._run_raceroster_import_async(
            request=request,
            input_dir=input_dir,
        )

    def _run_raceroster_import_async(
        self,
        request: RaceImportRequest,
        input_dir: Path,
    ) -> None:
        """Run Race Roster import off the UI thread with progress feedback."""
        result_queue: queue.Queue = queue.Queue()

        progress = tk.Toplevel(self)
        progress.title("Importing Race Roster")
        progress.transient(self)
        progress.grab_set()
        progress.resizable(False, False)
        progress.configure(bg=WRRL_LIGHT)
        progress.geometry("480x140")

        tk.Label(
            progress,
            text="Importing race results. This can take a moment...",
            font=("Segoe UI", 10),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
            wraplength=440,
            justify="left",
        ).pack(fill="x", padx=16, pady=(16, 8))

        bar = ttk.Progressbar(progress, mode="indeterminate", length=430)
        bar.pack(padx=16, pady=(0, 12))
        bar.start(12)

        tk.Label(
            progress,
            text="Please wait...",
            font=("Segoe UI", 9, "italic"),
            bg=WRRL_LIGHT,
            fg="#666666",
        ).pack(padx=16, pady=(0, 8))

        def _ignore_close() -> None:
            messagebox.showinfo(
                "Import Running",
                "Import is still running. Please wait for it to finish.",
                parent=progress,
            )

        progress.protocol("WM_DELETE_WINDOW", _ignore_close)

        def _worker() -> None:
            try:
                output_path, count, history_path = import_raceroster_results(
                    race_url=request.race_url,
                    input_dir=input_dir,
                    league_race_number=request.race_number,
                    race_name_override=request.race_name,
                    sporthive_race_id_hint=request.sporthive_race_hint,
                )
            except SporthiveRaceNotDirectlyImportableError:
                result_queue.put(("manual", None))
            except Exception as exc:
                result_queue.put(("error", str(exc)))
            else:
                result_queue.put(("ok", (output_path, count, history_path)))

        threading.Thread(target=_worker, daemon=True).start()

        def _poll_result() -> None:
            if not progress.winfo_exists():
                return
            try:
                status, payload = result_queue.get_nowait()
            except queue.Empty:
                self.after(120, _poll_result)
                return

            bar.stop()
            progress.grab_release()
            progress.destroy()

            if status == "manual":
                log_event(
                    "dashboard_import_raceroster_manual_required",
                    level="WARNING",
                    year=session_config.year,
                    race_number=request.race_number,
                    race_url=request.race_url,
                )
                use_manual = messagebox.askyesno(
                    "Sporthive Manual Import",
                    "This Sporthive race cannot be imported directly via API.\n\n"
                    "Switch to manual page-paste mode for this race?",
                    parent=self,
                )
                if use_manual:
                    self._run_manual_sporthive_import_dashboard(
                        request=request,
                        input_dir=input_dir,
                    )
                return

            if status == "error":
                log_event(
                    "dashboard_import_raceroster_failed",
                    level="ERROR",
                    year=session_config.year,
                    race_number=request.race_number,
                    race_url=request.race_url,
                    error=str(payload),
                )
                messagebox.showerror("Race Roster import failed", str(payload), parent=self)
                return

            output_path, count, history_path = payload
            self._refresh_config_panel()
            log_event(
                "dashboard_import_raceroster_completed",
                year=session_config.year,
                race_number=request.race_number,
                imported_rows=count,
                output_path=output_path,
                history_path=history_path,
            )
            messagebox.showinfo(
                "Import Complete",
                f"Imported {count} rows to:\n{output_path}\n\n"
                f"Import history:\n{history_path}",
                parent=self,
            )

        self.after(120, _poll_result)

    def _run_manual_sporthive_import_dashboard(
        self,
        request: RaceImportRequest,
        input_dir: Path,
    ) -> None:
        def _ask_page(page_no: int) -> str | None:
            return ask_multiline_page_text(
                self,
                title=f"Sporthive Page {page_no}",
                prompt=(
                    f"Paste results table text for Sporthive page {page_no}.\n"
                    "Copy the rows shown on screen (including pipe-delimited lines) and click 'Use This Page'."
                ),
                bg=WRRL_LIGHT,
                fg=WRRL_NAVY,
                panel_bg="#dbe1e8",
                accent_bg=WRRL_GREEN,
                accent_fg=WRRL_WHITE,
            )

        result = run_manual_sporthive_import(
            self,
            input_dir=input_dir,
            request=request,
            ask_page_text=_ask_page,
        )
        if result is None:
            log_event(
                "dashboard_import_sporthive_manual_cancelled",
                level="WARNING",
                year=session_config.year,
                race_number=request.race_number,
            )
            return

        output_path, count, history_path = result
        self._refresh_config_panel()
        log_event(
            "dashboard_import_sporthive_manual_completed",
            year=session_config.year,
            race_number=request.race_number,
            imported_rows=count,
            output_path=output_path,
            history_path=history_path,
        )

        messagebox.showinfo(
            "Import Complete",
            f"Imported {count} rows to:\n{output_path}\n\n"
            f"Import history:\n{history_path}",
            parent=self,
        )


# ── entry point ───────────────────────────────────────────────────────────────

def launch_dashboard() -> None:
    """Create and run the dashboard application."""
    app = LeagueScorerDashboard()
    app.mainloop()
