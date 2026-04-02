"""
dashboard.py — Main dashboard for the Wiltshire Road and Running League Management.

Provides a professional console with branded header and execution options.
"""

import tkinter as tk
import queue
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

from .gui import LeagueScorerApp
from .events_viewer import EventsViewerWindow
from ..events_loader import load_events
from ..raceroster_import import (
    SporthiveRaceNotDirectlyImportableError,
    import_raceroster_results,
)
from ..session_config import config as session_config
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


# ──────────────────────────────────────────────────────────────────────────────
# Main dashboard window
# ──────────────────────────────────────────────────────────────────────────────

class LeagueScorerDashboard(tk.Tk):
    """Professional dashboard for the Wiltshire League Management."""

    def __init__(self) -> None:
        super().__init__()
        self.title("WRRL League Management")
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
            text="WRRL League Management",
            font=("Segoe UI", 32, "bold"),
            bg=WRRL_NAVY,
            fg=WRRL_WHITE,
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            title_frame,
            text="Wiltshire Road and Running League Management System",
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

        buttons = [
            ("📅 Load Events", "Load season events spreadsheet", self._on_load_events, 1, 0, "secondary"),
            ("📋 View Events", "Browse loaded events schedule", self._on_view_events, 1, 1, "secondary"),

            ("🧪 Audit Races", "Audit and standardise race workbook names", self._on_audit_runners, 2, 0, "secondary"),
            ("🔎 View Audit", "Review latest audit workbook", self._on_view_audit, 2, 1, "secondary"),

            ("▶ Create League Results", "Run scoring and generate outputs", self._on_run_scorer, 3, 0, "primary"),
            ("📊 View Results", "Open generated standings and reports", self._on_view_results, 3, 1, "secondary"),

            ("✏️ Edit Clubs", "Manually correct runner club assignments", self._on_edit_clubs, 4, 0, "secondary"),
            ("🏟️ Club History", "View one club across all races", self._on_view_club_history, 4, 1, "secondary"),

            ("🧩 Check All Runners", "Suggest blank-club fixes across races", self._on_check_all_runners, 5, 0, "secondary"),
            ("🏃 Runner History", "View one runner across all races", self._on_view_runner_history, 5, 1, "secondary"),

            ("🌐 Import Race Roster", "Pull race results into this season", self._on_import_raceroster, 6, 0, "primary"),
            ("⚙️ Settings", "League scoring and report options", self._on_settings, 6, 1, "secondary"),
        ]

        for text, subtitle, cmd, row, col, tone in buttons:
            self._create_action_button(button_frame, text, subtitle, cmd, row, col, tone=tone)

    # ── config panel ──────────────────────────────────────────────────────────

    def _build_config_panel(self, parent: tk.Frame) -> None:
        """Build the season / path configuration panel."""
        panel = tk.Frame(parent, bg=WRRL_NAVY, padx=12, pady=10)
        panel.pack(fill="x", pady=(0, 4))

        # Two-column layout: season picker (left) | data-root (right)
        cols = tk.Frame(panel, bg=WRRL_NAVY)
        cols.pack(fill="x")
        cols.columnconfigure(1, weight=1)

        # ── Season picker ────────────────────────────────────────────────────
        season_frame = tk.Frame(cols, bg=WRRL_NAVY)
        season_frame.grid(row=0, column=0, padx=(0, 24), sticky="ns")

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

        # ── Data Root ────────────────────────────────────────────────────────
        root_frame = tk.Frame(cols, bg=WRRL_NAVY)
        root_frame.grid(row=0, column=1, sticky="nsew")

        tk.Label(root_frame, text="Data Root:", font=("Segoe UI", 10, "bold"),
                 bg=WRRL_NAVY, fg="#a0b0c0").pack(anchor="w")

        root_row = tk.Frame(root_frame, bg=WRRL_NAVY)
        root_row.pack(fill="x", pady=(2, 0))

        self._root_label = tk.Label(
            root_row,
            text=str(session_config.data_root) if session_config.data_root else "Not set",
            font=("Segoe UI", 9),
            bg=WRRL_NAVY,
            fg="#ffcc44" if not session_config.data_root else "#90ee90",
            anchor="w",
        )
        self._root_label.pack(side="left", fill="x", expand=True, padx=(0, 10))

        tk.Button(
            root_row, text="Set Root\u2026",
            font=("Segoe UI", 9, "bold"),
            bg=WRRL_GREEN, fg=WRRL_WHITE,
            relief="flat", padx=8, pady=2, cursor="hand2",
            command=self._on_set_data_root,
            activebackground="#1f5632", activeforeground=WRRL_WHITE,
        ).pack(side="right")

        # ── Status row (input / output / events) ────────────────────────────
        row2 = tk.Frame(panel, bg=WRRL_NAVY)
        row2.pack(fill="x", pady=(8, 0))

        self._input_status  = self._make_status_label(row2, "Input:",  "left")
        self._output_status = self._make_status_label(row2, "Output:", "left")
        self._events_status = self._make_status_label(row2, "Events:", "left")
        self._refresh_config_panel()

    def _make_status_label(self, parent: tk.Frame, prefix: str, side: str) -> tk.Label:
        tk.Label(parent, text=prefix, font=("Segoe UI", 9, "bold"),
                 bg=WRRL_NAVY, fg=WRRL_WHITE).pack(side=side, padx=(0, 4))
        lbl = tk.Label(parent, text="—", font=("Segoe UI", 9),
                       bg=WRRL_NAVY, fg="#ffcc44", anchor="w")
        lbl.pack(side=side, padx=(0, 16))
        return lbl

    def _refresh_config_panel(self) -> None:
        """Update config bar indicators to reflect current session_config state."""
        if not hasattr(self, "_root_label"):
            return
        if hasattr(self, "_year_label"):
            self._year_label.config(text=str(session_config.year))
        if session_config.data_root:
            self._root_label.config(text=str(session_config.data_root), fg="#90ee90")
        else:
            self._root_label.config(text="Not set", fg="#ffcc44")

        in_path  = session_config.input_dir
        out_path = session_config.output_dir
        ev_path  = session_config.events_path

        self._input_status.config(
            text=str(in_path) if in_path else "—",
            fg="#90ee90" if (in_path and in_path.exists()) else "#ffcc44",
        )
        self._output_status.config(
            text=str(out_path) if out_path else "—",
            fg="#90ee90" if (out_path and out_path.exists()) else "#ffcc44",
        )
        self._events_status.config(
            text=ev_path.name if ev_path else "—",
            fg="#90ee90" if (ev_path and ev_path.exists()) else "#ffcc44",
        )

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

    def _on_set_data_root(self) -> None:
        """Browse for and set the data root folder."""
        initial = str(session_config.data_root) if session_config.data_root else "/"
        folder = filedialog.askdirectory(
            title="Select Data Root Folder (parent of all year folders)",
            initialdir=initial,
        )
        if folder:
            session_config.data_root = Path(folder)
            session_config.ensure_dirs()
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

    # ── footer section ────────────────────────────────────────────────────────

    def _build_footer(self) -> None:
        """Build the footer with version info."""
        footer = tk.Frame(self, bg=WRRL_NAVY, height=50)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        from .. import __version__
        footer_text = tk.Label(
            footer,
            text=f"© 2026 Wiltshire Athletics Assoc. | Wiltshire League Scorer v{__version__}",
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
                "Use the 'Set Root…' button in the configuration panel above.",
            )
            return False
        return True

    # ── action handlers ────────────────────────────────────────────────────

    def _on_run_scorer(self) -> None:
        """Show the League Management scorer panel inline within the dashboard."""
        if not self._require_configured("Run League Management"):
            return
        session_config.ensure_dirs()
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

    def _on_load_events(self) -> None:
        """Browse for and load an events XLSX file."""
        if not self._require_configured("Load Events"):
            return
        initial_path = session_config.events_path
        if initial_path and initial_path.parent.exists():
            initial_dir = str(initial_path.parent)
        else:
            initial_dir = str(session_config.input_dir) if session_config.input_dir else "/"
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
        """Open the events viewer window."""
        if not self._require_configured("View Events"):
            return
        if self._events_schedule is None:
            if messagebox.askyesno(
                "No Events Loaded",
                "No events spreadsheet is currently loaded.\nWould you like to load one now?",
            ):
                self._on_load_events()
            if self._events_schedule is None:
                return
        images_dir = Path(__file__).parent.parent / "images"
        EventsViewerWindow(
            self,
            self._events_schedule,
            year=session_config.year,
            images_dir=images_dir,
            output_dir=session_config.output_dir,
        )

    def _on_view_results(self) -> None:
        """Show the results viewer panel inline within the dashboard."""
        if not self._require_configured("View Results"):
            return
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
            text="\u25c4 Dashboard",
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
        panel = SettingsPanel(self._page_container, on_close=on_close)
        panel.pack(fill="both", expand=True)
        self._settings_panel = panel

    def _on_audit_runners(self) -> None:
        """Show the audit runner panel inline within the dashboard."""
        if not self._require_configured("Audit Runners"):
            return
        session_config.ensure_dirs()
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

        close_text = "\u25c4 Audit" if return_to_audit else "\u25c4 Dashboard"
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

    def _on_edit_clubs(self) -> None:
        """Show the club assignment editor panel inline within the dashboard."""
        if not self._require_configured("Edit Clubs"):
            return
        session_config.ensure_dirs()
        self._home_frame.pack_forget()
        from .club_editor import ClubEditorPanel
        panel = ClubEditorPanel(
            self._page_container,
            back_callback=self._on_edit_clubs_back,
        )
        panel.pack(fill="both", expand=True)
        self._club_editor_panel = panel

    def _on_edit_clubs_back(self) -> None:
        if hasattr(self, "_club_editor_panel"):
            self._club_editor_panel.destroy()
            del self._club_editor_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_check_all_runners(self) -> None:
        """Show cross-race runner club checks panel inline within the dashboard."""
        if not self._require_configured("Check All Runners"):
            return
        session_config.ensure_dirs()
        self._home_frame.pack_forget()
        from .check_all_runners import CheckAllRunnersPanel
        panel = CheckAllRunnersPanel(
            self._page_container,
            back_callback=self._on_check_all_runners_back,
        )
        panel.pack(fill="both", expand=True)
        self._check_all_runners_panel = panel

    def _on_check_all_runners_back(self) -> None:
        if hasattr(self, "_check_all_runners_panel"):
            self._check_all_runners_panel.destroy()
            del self._check_all_runners_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_view_runner_history(self) -> None:
        """Show runner history viewer panel inline within the dashboard."""
        if not self._require_configured("Runner History"):
            return
        self._home_frame.pack_forget()
        from .runner_history_viewer import RunnerHistoryPanel
        panel = RunnerHistoryPanel(
            self._page_container,
            back_callback=self._on_view_runner_history_back,
        )
        panel.pack(fill="both", expand=True)
        self._runner_history_panel = panel

    def _on_view_runner_history_back(self) -> None:
        if hasattr(self, "_runner_history_panel"):
            self._runner_history_panel.destroy()
            del self._runner_history_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_view_club_history(self) -> None:
        """Show club history viewer panel inline within the dashboard."""
        if not self._require_configured("Club History"):
            return
        self._home_frame.pack_forget()
        from .club_history_viewer import ClubHistoryPanel
        panel = ClubHistoryPanel(
            self._page_container,
            back_callback=self._on_view_club_history_back,
        )
        panel.pack(fill="both", expand=True)
        self._club_history_panel = panel

    def _on_view_club_history_back(self) -> None:
        if hasattr(self, "_club_history_panel"):
            self._club_history_panel.destroy()
            del self._club_history_panel
        self._home_frame.pack(fill="both", expand=True)

    def _on_audit_complete_view(self, preferred_workbook=None) -> None:
        self._on_view_audit(preferred_workbook=preferred_workbook, return_to_audit=True)

    def _on_help(self) -> None:
        """Show help information."""
        docs_dir = Path(__file__).resolve().parents[2] / "documents"
        dependencies_doc = docs_dir / "dependencies.md"
        ops_doc = docs_dir / "operational_dependencies.md"

        help_text = (
            "WRRL League Management Help\n\n"
            "• Set Root…: Choose your base data folder once.\n"
            "  Folders are created as: {root}/{year}/inputs and {root}/{year}/outputs\n\n"
            "• Season: Select the current league year.\n\n"
            "• Load Events: Select the events XLSX from the season inputs folder.\n"
            "  The filename is remembered and reused for the active season folder on later runs.\n"
            "• View Events: Browse the loaded events schedule.\n\n"
            "• Import Race Roster: Paste a Race Roster URL and save it directly\n"
            "  as a race workbook in the active season inputs folder.\n\n"
            "• Run League Management: Execute the scoring pipeline.\n"
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
        input_dir = session_config.input_dir
        if not input_dir or not input_dir.exists():
            messagebox.showerror("Input not found", f"Input folder does not exist:\n{input_dir}")
            return

        request = prompt_race_import_request(self)
        if request is None:
            return

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
                messagebox.showerror("Race Roster import failed", str(payload), parent=self)
                return

            output_path, count, history_path = payload
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
            return

        output_path, count, history_path = result

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
