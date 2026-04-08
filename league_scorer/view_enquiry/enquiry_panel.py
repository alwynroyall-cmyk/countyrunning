import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional
import io
import tempfile
from pathlib import Path
import re
import string

import pandas as pd

from ..session_config import config as session_config
from ..graphical.results_workbook import find_latest_results_workbook, sorted_race_sheet_names
from ..events_loader import load_events
from ..race_processor import extract_race_number
from ..club_loader import load_clubs


def _norm_for_compare(s: str) -> str:
    if s is None:
        return ""
    # remove punctuation, lower-case and collapse spaces
    s2 = re.sub(r"[^A-Za-z0-9 ]+", "", str(s))
    return re.sub(r"\s+", " ", s2).strip().lower()


def _format_time_str(value) -> str:
    """Return time formatted as H:MM:SS.0 when possible, else original string."""
    if value is None:
        return ""
    try:
        # If it's already a pandas Timedelta-like
        if hasattr(value, "total_seconds"):
            total = float(value.total_seconds())
        else:
            td = pd.to_timedelta(value)
            total = float(td.total_seconds())
        total_s = int(total)
        hh = total_s // 3600
        mm = (total_s % 3600) // 60
        ss = total_s % 60
        return f"{hh}:{mm:02d}:{ss:02d}.0"
    except Exception:
        # fallback: try to extract parts with regex
        s = str(value).strip()
        m = re.match(r"^(?:(\d+):)?(\d{1,2}):(\d{2})(?:[.,](\d+))?$", s)
        if m:
            hh = int(m.group(1) or 0)
            mm = int(m.group(2))
            ss = int(m.group(3))
            return f"{hh}:{mm:02d}:{ss:02d}.0"
        return s


def _extract_distance_label(name: str) -> str:
    """Try to identify a short distance label from an event or sheet name.
    Examples: '5k', '10k', '5 mile'. Returns 'Other' when unknown.
    """
    if not name:
        return "Other"
    s = str(name).lower()
    # look for kilometres
    m = re.search(r"(\d+(?:\.\d+)?)\s*km\b", s)
    if not m:
        m = re.search(r"(\d+(?:\.\d+)?)\s*k\b", s)
    if m:
        try:
            k = float(m.group(1))
            if k.is_integer():
                return f"{int(k)}k"
            return f"{k}k"
        except Exception:
            return "Other"
    # look for miles
    m = re.search(r"(\d+(?:\.\d+)?)\s*mile\b", s)
    if not m:
        m = re.search(r"(\d+(?:\.\d+)?)\s*mi\b", s)
    if m:
        try:
            mv = float(m.group(1))
            if mv.is_integer():
                return f"{int(mv)} mile"
            return f"{mv} mile"
        except Exception:
            return "Other"
    # common short tokens like '5mile', '10mile' or '5mile' without space
    m = re.search(r"\b(\d+)[^\d\s]*(k|km)\b", s)
    if m:
        return f"{int(m.group(1))}k"
    m = re.search(r"\b(\d+)[^\d\s]*mile\b", s)
    if m:
        return f"{int(m.group(1))} mile"
    return "Other"


def _time_to_seconds(tval) -> Optional[float]:
    """Convert various time representations to seconds, or return None."""
    if tval is None:
        return None
    s = tval
    # numeric
    try:
        if isinstance(s, (int, float)):
            return float(s)
    except Exception:
        pass
    # pandas Timedelta-like
    try:
        if hasattr(s, "total_seconds"):
            return float(s.total_seconds())
    except Exception:
        pass
    # string parsing
    try:
        ss = str(s).strip()
        if ss == "":
            return None
        try:
            td = pd.to_timedelta(ss)
            return float(td.total_seconds())
        except Exception:
            pass
        m = re.match(r"(?:(\d+):)?(\d{1,2}):(\d{2})(?:[.,](\d+))?$", ss)
        if m:
            hh = int(m.group(1) or 0)
            mm = int(m.group(2))
            sec = int(m.group(3))
            return hh * 3600 + mm * 60 + sec
        m2 = re.match(r"^(\d+):(\d{2})(?:[.,](\d+))?$", ss)
        if m2:
            mm = int(m2.group(1))
            sec = int(m2.group(2))
            return mm * 60 + sec
    except Exception:
        pass
    return None


class RunnerClubEnquiryPanel(tk.Frame):
    """Panel to search published results by runner or club.

    Runner tab: type-ahead combobox for runner names, a compact summary
    card and a detailed per-race table. Data is read-only and sourced from
    the latest published standings workbook.
    """

    def __init__(self, parent, back_callback=None, initial_runner: Optional[str] = None):
        super().__init__(parent, bg="#f7f9fb")
        self._back_callback = back_callback
        self._initial_runner = initial_runner
        self._xl_cache: tuple[pd.ExcelFile | None, Path | None, float] = (None, None, 0.0)
        self._all_names: list[str] = []
        self._name_map: dict[str, str] = {}
        self._build_ui()
        # best-effort pre-select
        if initial_runner:
            self.select_runner(initial_runner)

    def _get_xl(self, results_path: Path) -> pd.ExcelFile | None:
        cached_xl, cached_path, cached_mtime = self._xl_cache
        try:
            mtime = results_path.stat().st_mtime
        except OSError:
            return None
        if cached_xl is not None and cached_path == results_path and cached_mtime == mtime:
            return cached_xl
        if cached_xl is not None:
            try:
                cached_xl.close()
            except Exception:
                pass
        try:
            xl = pd.ExcelFile(results_path)
            self._xl_cache = (xl, results_path, mtime)
            return xl
        except Exception:
            self._xl_cache = (None, None, 0.0)
            return None

    def _build_ui(self):
        header = tk.Frame(self, bg="#22313f", padx=12, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Runner / Club Enquiry",
            font=("Segoe UI", 14, "bold"),
            bg="#22313f",
            fg="#ffffff",
        ).pack(side="left")

        if self._back_callback:
            tk.Button(
                header,
                text="🏠 Dashboard",
                command=self._back_callback,
                font=("Segoe UI", 10, "bold"),
                bg="#f4f6f8",
                fg="#2d7a4a",
                relief="flat",
                padx=8,
                pady=4,
            ).pack(side="right")

        # Notebook for Runner / Club (Club tab left as placeholder)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=12)

        # Runner tab
        runner_frame = tk.Frame(nb, bg="#f7f9fb")
        # keep reference so we can capture only this panel
        self._runner_frame = runner_frame
        nb.add(runner_frame, text="Runner")

        # Search combobox variables — widget created in the right column below
        self._runner_var = tk.StringVar()
        self._runner_combo = None

        # Refresh removed to allow fluid typing — names load automatically

        # Two-column layout: left summary card, right search + details
        content = tk.Frame(runner_frame, bg="#f7f9fb")
        content.pack(fill="both", expand=True, pady=(8, 0))

        left = tk.Frame(content, bg="#ffffff", padx=16, pady=16, width=320)
        left.pack(side="left", fill="y", padx=(0, 12))
        # keep reference for screenshotting
        self._summary_frame = left

        # centre (middle) panel: reduce size so right panel can be larger
        right = tk.Frame(content, bg="#f7f9fb", width=640)
        right.pack(side="left", fill="both")
        right.pack_propagate(False)

        # Summary card (left)
        self._summary_vars = {
            "NameClub": tk.StringVar(value=""),
            "Club": tk.StringVar(value=""),
            "Category": tk.StringVar(value=""),
            "Points": tk.StringVar(value=""),
            "Races": tk.StringVar(value=""),
        }

        # Avatar row
        avatar_row = tk.Frame(left, bg="#ffffff")
        avatar_row.pack(fill="x")
        avatar_canvas = tk.Canvas(avatar_row, width=64, height=64, bg="#2d7a4a", highlightthickness=0)
        avatar_canvas.create_oval(4, 4, 60, 60, fill="#2d7a4a", outline="")
        avatar_canvas.create_text(32, 34, text="🏃", font=("Segoe UI Emoji", 20))
        avatar_canvas.pack(side="left")

        name_frame = tk.Frame(avatar_row, bg="#ffffff")
        name_frame.pack(side="left", padx=12)
        tk.Label(name_frame, textvariable=self._summary_vars["NameClub"], bg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(name_frame, textvariable=self._summary_vars["Club"], bg="#ffffff", font=("Segoe UI", 10)).pack(anchor="w")

        # Stats
        stats = tk.Frame(left, bg="#ffffff")
        stats.pack(fill="x", pady=(12, 0))
        tk.Label(stats, text="Category:", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(stats, textvariable=self._summary_vars["Category"], bg="#ffffff", font=("Segoe UI", 12)).grid(row=0, column=1, sticky="w", padx=(8,0))
        tk.Label(stats, text="Points:", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w", pady=(8,0))
        tk.Label(stats, textvariable=self._summary_vars["Points"], bg="#ffffff", font=("Segoe UI", 12)).grid(row=1, column=1, sticky="w", padx=(8,0), pady=(8,0))
        tk.Label(stats, text="Races:", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(8,0))
        tk.Label(stats, textvariable=self._summary_vars["Races"], bg="#ffffff", font=("Segoe UI", 12)).grid(row=2, column=1, sticky="w", padx=(8,0), pady=(8,0))

        tk.Button(left, text="📋 Copy", command=self._copy_details, font=("Segoe UI", 9), bg="#e8f0f7", fg="#22313f", relief="flat").pack(anchor="w", pady=(12,0))

        # Right: create search bar above the details table
        search_bar_right = tk.Frame(right, bg="#f7f9fb", padx=6, pady=6)
        search_bar_right.pack(fill="x")
        tk.Label(search_bar_right, text="Runner:", bg="#f7f9fb").pack(side="left")
        self._runner_combo = ttk.Combobox(search_bar_right, textvariable=self._runner_var)
        self._runner_combo.pack(side="left", padx=(8, 8), fill="x", expand=True)
        self._runner_combo.bind("<<ComboboxSelected>>", lambda e: self._on_runner_selected())
        self._runner_combo.bind("<KeyRelease>", self._on_runner_key)
        # helper for debouncing the filter while the user types
        self._filter_after_id = None
        self._runner_combo.bind("<Return>", self._on_runner_return)

        # style Treeview for a lighter, modern look
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Enquiry.Treeview", background="#ffffff", fieldbackground="#ffffff", rowheight=28)
        style.configure("Enquiry.Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#f1f5f9")

        details_frame = tk.Frame(right, bg="#f7f9fb")
        details_frame.pack(fill="both", expand=True, pady=(8,0))

        cols = ["Race", "Club", "Time", "Points"]
        self._tree = ttk.Treeview(details_frame, columns=cols, show="headings", style="Enquiry.Treeview")
        for c in cols:
            self._tree.heading(c, text=c, anchor="center")
            if c == "Race":
                self._tree.column(c, width=180, anchor="w")
            elif c == "Club":
                self._tree.column(c, width=200, anchor="w")
            elif c == "Time":
                self._tree.column(c, width=100, anchor="center")
            else:
                # make Points slightly wider so values are visible
                self._tree.column(c, width=120, anchor="center")
        self._tree.pack(fill="both", expand=True, side="left")

        # alternating row colours
        self._tree.tag_configure("odd", background="#ffffff")
        self._tree.tag_configure("even", background="#f6f8fa")

        yscroll = ttk.Scrollbar(details_frame, orient="vertical", command=self._tree.yview)
        yscroll.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=yscroll.set)

        # Right-hand panel: Best times by distance
        # enlarge right-hand Best-by-Distance panel
        right_panel = tk.Frame(content, bg="#ffffff", padx=12, pady=12, width=360)
        right_panel.pack(side="left", fill="y", padx=(12, 0))
        right_panel.pack_propagate(False)

        tk.Label(right_panel, text="Best by Distance", bg="#ffffff", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        cols2 = ["Distance", "Best Time", "Event"]
        self._best_tree = ttk.Treeview(right_panel, columns=cols2, show="headings", height=10)
        for c in cols2:
            self._best_tree.heading(c, text=c, anchor="center")
            if c == "Distance":
                self._best_tree.column(c, width=80, anchor="center")
            elif c == "Best Time":
                self._best_tree.column(c, width=100, anchor="center")
            else:
                self._best_tree.column(c, width=220, anchor="w")
        self._best_tree.pack(fill="both", expand=True, pady=(8,0))
        self._best_tree.tag_configure("odd", background="#ffffff")
        self._best_tree.tag_configure("even", background="#f6f8fa")

        # Club tab (runners view)
        club_frame = tk.Frame(nb, bg="#f7f9fb")
        nb.add(club_frame, text="Club - Runners")

        # Club search bar
        club_search = tk.Frame(club_frame, bg="#f7f9fb", padx=6, pady=6)
        club_search.pack(fill="x")
        tk.Label(club_search, text="Club:", bg="#f7f9fb").pack(side="left")
        self._club_var = tk.StringVar()
        self._club_combo = ttk.Combobox(club_search, textvariable=self._club_var)
        self._club_combo.pack(side="left", padx=(8,8), fill="x", expand=True)
        self._club_combo.bind("<<ComboboxSelected>>", lambda e: self._on_club_selected())

        # Club content: left summary, right members
        club_content = tk.Frame(club_frame, bg="#f7f9fb")
        club_content.pack(fill="both", expand=True, padx=12, pady=(8,0))

        club_left = tk.Frame(club_content, bg="#ffffff", padx=12, pady=12, width=300)
        club_left.pack(side="left", fill="y", padx=(0,12))
        club_left.pack_propagate(False)

        club_right = tk.Frame(club_content, bg="#f7f9fb")
        club_right.pack(side="left", fill="both", expand=True)

        # Club summary — copy of the Club - Races left panel (image, runners, aggregated scores, male/female, team summaries)
        self._club_vars = {
            "ClubName": tk.StringVar(value=""),
            "Members": tk.StringVar(value=""),
            "TotalPoints": tk.StringVar(value=""),
            "MaleScore": tk.StringVar(value=""),
            "FemaleScore": tk.StringVar(value=""),
            "TeamA_Label": tk.StringVar(value=""),
            "TeamB_Label": tk.StringVar(value=""),
            "TeamA_Summary": tk.StringVar(value=""),
            "TeamB_Summary": tk.StringVar(value=""),
        }
        # optional header image
        self._club_image_photo = None
        self._club_image_label = tk.Label(club_left, bg="#ffffff")
        self._club_image_label.pack(fill="x", pady=(0,8))

        tk.Label(club_left, textvariable=self._club_vars["ClubName"], bg="#ffffff", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(6,0))
        tk.Label(club_left, text="Runners:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12,0))
        tk.Label(club_left, textvariable=self._club_vars["Members"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        tk.Label(club_left, text="Aggregated Scores:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12,0))
        tk.Label(club_left, textvariable=self._club_vars["TotalPoints"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        tk.Label(club_left, text="Male Scorers:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10,0))
        tk.Label(club_left, textvariable=self._club_vars["MaleScore"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        tk.Label(club_left, text="Female Scorers:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8,0))
        tk.Label(club_left, textvariable=self._club_vars["FemaleScore"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        tk.Label(club_left, textvariable=self._club_vars["TeamA_Label"], bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10,0))
        tk.Label(club_left, textvariable=self._club_vars["TeamA_Summary"], bg="#ffffff", font=("Segoe UI", 10), justify="left", anchor="w").pack(anchor="w")
        tk.Label(club_left, textvariable=self._club_vars["TeamB_Label"], bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8,0))
        tk.Label(club_left, textvariable=self._club_vars["TeamB_Summary"], bg="#ffffff", font=("Segoe UI", 10), justify="left", anchor="w").pack(anchor="w")

        mcols = ["Member", "Category", "Points", "Races"]
        self._club_tree = ttk.Treeview(club_right, columns=mcols, show="headings")
        for c in mcols:
            self._club_tree.heading(c, text=c)
            if c == "Member":
                self._club_tree.column(c, width=240, anchor="w")
            else:
                self._club_tree.column(c, width=100, anchor="center")
        self._club_tree.pack(fill="both", expand=True, side="left")
        club_ys = ttk.Scrollbar(club_right, orient="vertical", command=self._club_tree.yview)
        club_ys.pack(side="right", fill="y")
        self._club_tree.configure(yscrollcommand=club_ys.set)

        # populate clubs now
        self._refresh_clubs()

        # Club - Races tab: per-club per-race summary (three-column layout)
        club_races_frame = tk.Frame(nb, bg="#f7f9fb")
        nb.add(club_races_frame, text="Club - Races")

        container = tk.Frame(club_races_frame, bg="#f7f9fb", padx=12, pady=12)
        container.pack(fill="both", expand=True)

        left_col = tk.Frame(container, bg="#ffffff", padx=12, pady=12, width=300)
        left_col.pack(side="left", fill="y", padx=(0,12))
        left_col.pack_propagate(False)

        center_col = tk.Frame(container, bg="#f7f9fb", padx=12, pady=12, width=400)
        center_col.pack(side="left", fill="both")
        center_col.pack_propagate(False)
        # centre header for Team A
        tk.Label(center_col, text="Team A", bg="#f7f9fb", font=("Segoe UI", 11, "bold")).pack(anchor="nw", pady=(0,6))

        right_col = tk.Frame(container, bg="#f7f9fb", padx=12, pady=12, width=400)
        right_col.pack(side="left", fill="y", padx=(12,0))
        right_col.pack_propagate(False)

        # right header for Team B will sit above the scroller
        tk.Label(right_col, text="Team B", bg="#f7f9fb", font=("Segoe UI", 11, "bold")).pack(anchor="nw", pady=(0,6))

        # left summary (Team A)
        self._club_races_photo = None
        self._club_races_image_label = tk.Label(left_col, bg="#ffffff")
        self._club_races_image_label.pack(fill="x", pady=(0,8))
        # (header image above; no separate 'Team A Summary' heading)

        self._club_races_vars = {
            "ClubName": tk.StringVar(value=""),
            "Members": tk.StringVar(value=""),
            "TotalPoints": tk.StringVar(value=""),
            "TeamA": tk.StringVar(value=""),
            "TeamB": tk.StringVar(value=""),
            "TeamA_Pos": tk.StringVar(value=""),
            "TeamB_Pos": tk.StringVar(value=""),
            "MaleScore": tk.StringVar(value=""),
            "FemaleScore": tk.StringVar(value=""),
            "TeamA_Label": tk.StringVar(value=""),
            "TeamB_Label": tk.StringVar(value=""),
            "TeamA_Summary": tk.StringVar(value=""),
            "TeamB_Summary": tk.StringVar(value=""),
        }
        tk.Label(left_col, textvariable=self._club_races_vars["ClubName"], bg="#ffffff", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(6,0))
        tk.Label(left_col, text="Runners:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12,0))
        tk.Label(left_col, textvariable=self._club_races_vars["Members"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        tk.Label(left_col, text="Aggregated Scores:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12,0))
        tk.Label(left_col, textvariable=self._club_races_vars["TotalPoints"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        # male/female breakdown rows (renamed)
        tk.Label(left_col, text="Male Scorers:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10,0))
        tk.Label(left_col, textvariable=self._club_races_vars["MaleScore"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        tk.Label(left_col, text="Female Scorers:", bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8,0))
        tk.Label(left_col, textvariable=self._club_races_vars["FemaleScore"], bg="#ffffff", font=("Segoe UI", 11)).pack(anchor="w")
        # Team summaries: show Team A then Team B as a short list (position and points)
        tk.Label(left_col, textvariable=self._club_races_vars["TeamA_Label"], bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10,0))
        tk.Label(left_col, textvariable=self._club_races_vars["TeamA_Summary"], bg="#ffffff", font=("Segoe UI", 10), justify="left", anchor="w").pack(anchor="w")
        tk.Label(left_col, textvariable=self._club_races_vars["TeamB_Label"], bg="#ffffff", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8,0))
        tk.Label(left_col, textvariable=self._club_races_vars["TeamB_Summary"], bg="#ffffff", font=("Segoe UI", 10), justify="left", anchor="w").pack(anchor="w")

        # center: scrollable container for Team A (per-race expandable cards)
        _a_scroller = tk.Frame(center_col, bg="#f7f9fb")
        _a_scroller.pack(fill="both", expand=True)
        _a_canvas = tk.Canvas(_a_scroller, bg="#f7f9fb", highlightthickness=0)
        _a_vsb = ttk.Scrollbar(_a_scroller, orient="vertical", command=_a_canvas.yview)
        _a_canvas.configure(yscrollcommand=_a_vsb.set)
        _a_vsb.pack(side="right", fill="y")
        _a_canvas.pack(side="left", fill="both", expand=True)
        _a_inner = tk.Frame(_a_canvas, bg="#f7f9fb")
        _a_win = _a_canvas.create_window((0, 0), window=_a_inner, anchor="nw")
        def _a_on_config(e, c=_a_canvas):
            try:
                c.configure(scrollregion=c.bbox("all"))
            except Exception:
                pass
        _a_inner.bind("<Configure>", _a_on_config)
        def _a_canvas_resize(e, c=_a_canvas, w=_a_win):
            try:
                c.itemconfig(w, width=e.width)
            except Exception:
                pass
        _a_canvas.bind("<Configure>", _a_canvas_resize)
        # expose the inner frame for content population
        self._team_a_center = _a_inner

        # right summary placeholder (removed empty label to keep alignment with centre)
        # below the summary, a scrollable container for Team B per-race cards
        # right: scrollable container for Team B per-race cards
        _b_scroller = tk.Frame(right_col, bg="#f7f9fb")
        _b_scroller.pack(fill="both", expand=True)
        _b_canvas = tk.Canvas(_b_scroller, bg="#f7f9fb", highlightthickness=0)
        _b_vsb = ttk.Scrollbar(_b_scroller, orient="vertical", command=_b_canvas.yview)
        _b_canvas.configure(yscrollcommand=_b_vsb.set)
        _b_vsb.pack(side="right", fill="y")
        _b_canvas.pack(side="left", fill="both", expand=True)
        _b_inner = tk.Frame(_b_canvas, bg="#f7f9fb")
        _b_win = _b_canvas.create_window((0, 0), window=_b_inner, anchor="nw")
        def _b_on_config(e, c=_b_canvas):
            try:
                c.configure(scrollregion=c.bbox("all"))
            except Exception:
                pass
        _b_inner.bind("<Configure>", _b_on_config)
        def _b_canvas_resize(e, c=_b_canvas, w=_b_win):
            try:
                c.itemconfig(w, width=e.width)
            except Exception:
                pass
        _b_canvas.bind("<Configure>", _b_canvas_resize)
        self._team_b_center = _b_inner


        # status
        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, bg="#f7f9fb", fg="#66707c").pack(fill="x")

        # populate names now
        self._refresh_names()

    def _on_runner_key(self, event):
        """Filter the combobox values as the user types (simple type-ahead)."""
        # ignore navigation/modifier keys to avoid unnecessary filtering
        if event and getattr(event, 'keysym', None) in (
            'Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'):
            return

        q = self._runner_var.get().strip().lower()

        # debounce: wait 150ms after typing stops to update values
        if self._filter_after_id:
            try:
                self.after_cancel(self._filter_after_id)
            except Exception:
                pass
        self._filter_after_id = self.after(150, lambda: self._apply_name_filter(q))

    def _apply_name_filter(self, q: str):
        """Apply filtering to the combobox values (debounced)."""
        self._filter_after_id = None
        if not q:
            values = []
        else:
            # normalize query for prefix lookup
            qnorm = _norm_for_compare(q).replace(" ", "")
            values = []
            # use prefix index for fast starts-with matches (try longest available up to 3)
            try:
                if hasattr(self, "_prefix_index") and qnorm:
                    key = qnorm[:3] if len(qnorm) >= 3 else qnorm
                    candidates = self._prefix_index.get(key, [])
                    # narrow candidates to those that actually start with the query
                    for n in candidates:
                        nk = _norm_for_compare(n).replace(" ", "")
                        if nk.startswith(qnorm):
                            values.append(n)
            except Exception:
                values = []
            # if not enough candidates, also include contains matches from full list
            if len(values) < 100:
                contains = []
                ql = q.lower()
                for n in self._all_names:
                    nl = n.lower()
                    if ql in nl and n not in values:
                        contains.append(n)
                values = values + contains
        # limit suggestions for performance/UI
        max_suggestions = 100
        if len(values) > max_suggestions:
            values = values[:max_suggestions]
        # only update if changed to avoid UI thrash
        try:
            current = list(self._runner_combo['values'])
        except Exception:
            current = []
        if values != current:
            self._runner_combo['values'] = values
            if values:
                try:
                    self._runner_combo.focus_set()
                    # open dropdown shortly after update
                    self.after(40, lambda: self._runner_combo.event_generate('<Down>'))
                except Exception:
                    pass

    def _on_runner_return(self, event):
        # Treat Enter as selection
        self._on_runner_selected()

    def _refresh_names(self):
        results_path = self._find_results_workbook()
        # fallback: if session config not set, try searching common output folders
        if results_path is None:
            root = Path.cwd()
            candidates = list(root.glob("output/**/*.xlsx"))
            if candidates:
                # prefer files with 'standings' or 'summary' in name
                pref = [p for p in candidates if any(x in p.stem.lower() for x in ("standings", "summary", "results"))]
                pick = pref[0] if pref else max(candidates, key=lambda p: p.stat().st_mtime)
                results_path = pick
        if results_path is None:
            self._status_var.set("No standings workbook found")
            self._all_names = []
            self._runner_combo["values"] = []
            return
        xl = self._get_xl(results_path)
        if xl is None:
            self._status_var.set(f"Failed to read workbook: {results_path.name}")
            self._runner_combo["values"] = []
            return
        # Try to read a Summary or Standings sheet (common for published standings)
        try:
            chosen = None
            for s in xl.sheet_names:
                n = s.lower()
                if "summary" in n or "standings" in n:
                    chosen = s
                    break
            if chosen:
                df = xl.parse(chosen)
            else:
                # fallback: try first sheet
                df = xl.parse(xl.sheet_names[0])
        except Exception as exc:
            self._status_var.set(f"Failed to parse summary: {exc}")
            self._runner_combo["values"] = []
            return

        # Guess name column by common names, case-insensitive match first
        name_col = None
        for candidate in ("Name", "Runner", "Full Name", "fullname"):
            if candidate in df.columns:
                name_col = candidate
                break
        if name_col is None:
            for c in df.columns:
                if str(c).lower() in ("name", "runner", "full name", "fullname"):
                    name_col = c
                    break

        # Heuristic fallback: pick the column with most alphabetic-containing values
        if name_col is None:
            best = None
            best_count = 0
            for c in df.columns:
                try:
                    cnt = int(df[c].dropna().astype(str).apply(lambda s: bool(re.search(r"[A-Za-z]", s))).sum())
                except Exception:
                    cnt = 0
                if cnt > best_count:
                    best_count = cnt
                    best = c
            if best_count > 0:
                name_col = best

        # Pair-column fallback: combine two columns (forename + surname)
        if name_col is None:
            cols = list(df.columns)
            best_pair = None
            best_pair_count = 0
            for i in range(len(cols)):
                for j in range(i + 1, len(cols)):
                    try:
                        series = (df[cols[i]].fillna("").astype(str).str.strip() + " " + df[cols[j]].fillna("").astype(str).str.strip())
                        cnt = int(series.apply(lambda s: bool(re.search(r"[A-Za-z]", s))).sum())
                    except Exception:
                        cnt = 0
                    if cnt > best_pair_count:
                        best_pair_count = cnt
                        best_pair = (cols[i], cols[j])
            if best_pair_count > 0 and best_pair is not None:
                # Create a temporary combined column
                df["__combined_name"] = (
                    df[best_pair[0]].fillna("").astype(str).str.strip() + " " + df[best_pair[1]].fillna("").astype(str).str.strip()
                )
                name_col = "__combined_name"

        if name_col is None:
            self._status_var.set("Could not locate name column in summary")
            self._all_names = []
            self._runner_combo["values"] = []
            return

        raw_names = df[name_col].dropna().astype(str).unique().tolist()

        # Normalize to Proper Case and dedupe preserving order; build map to raw values
        seen = set()
        norms: list[str] = []
        name_map: dict[str, str] = {}
        for raw in raw_names:
            norm = string.capwords(raw.strip())
            key = norm.lower()
            if key in seen:
                continue
            seen.add(key)
            norms.append(norm)
            name_map[key] = raw.strip()

        names = norms

        # If the names look like sheet names (e.g., 'Race 1', 'Race 2'),
        # aggregate runner names from race sheets instead.
        sheet_like_re = re.compile(r"^(race\b|standings\b|summary\b|results\b)", re.IGNORECASE)
        if all(sheet_like_re.search(n) for n in names[:5]) if names else False:
            agg = []
            agg_set = set()
            for s in xl.sheet_names:
                try:
                    s_df = xl.parse(s)
                except Exception:
                    continue
                candidate_cols = [c for c in s_df.columns if str(c).lower() in ("name","runner","full name","fullname")]
                if not candidate_cols:
                    # fallback heuristic: pick column with alphabetic values
                    for c in s_df.columns:
                        try:
                            cnt = int(s_df[c].dropna().astype(str).apply(lambda v: bool(re.search(r"[A-Za-z]", v))).sum())
                        except Exception:
                            cnt = 0
                        if cnt > 0:
                            candidate_cols.append(c)
                            break
                if candidate_cols:
                    col = candidate_cols[0]
                    vals = s_df[col].dropna().astype(str).tolist()
                    for v in vals:
                        n = string.capwords(v.strip())
                        k = n.lower()
                        if k and k not in agg_set:
                            agg_set.add(k)
                            agg.append(n)
                            name_map[k] = v.strip()
            names = agg

        # filter out obvious sheet/file-like entries and sort alphabetically
        bad_token_re = re.compile(r"\b(race|standings|summary|results)\b", re.IGNORECASE)
        cleaned = []
        seen2 = set()
        for n in names:
            s = str(n).strip()
            if not s:
                continue
            # skip obvious filenames or sheet-like entries
            if s.lower().endswith('.xlsx'):
                continue
            if sheet_like_re.search(s):
                continue
            if bad_token_re.search(s):
                continue
            key = s.lower()
            if key in seen2:
                continue
            seen2.add(key)
            cleaned.append(s)
        if not cleaned:
            cleaned = names
        # sort alphabetically for predictable selection
        names_sorted = sorted(cleaned, key=lambda x: x.lower())
        # rebuild name_map for the sorted list
        name_map2 = {}
        for n in names_sorted:
            name_map2[n.lower()] = name_map.get(n.lower(), n)

        # final store
        self._all_names = names_sorted
        self._name_map = name_map2
        # don't pre-fill the combobox with all names (keeps typing responsive)
        self._runner_combo["values"] = []
        self._status_var.set(f"Workbook: {results_path.name} — {len(names_sorted)} runners")

        # Build a small prefix index for quick lookup (1..3 chars)
        self._prefix_index: dict[str, list[str]] = {}
        try:
            for nm in self._all_names:
                # normalized key without spaces for prefixing
                nk = _norm_for_compare(nm).replace(" ", "")
                for l in range(1, min(3, len(nk)) + 1):
                    k = nk[:l]
                    self._prefix_index.setdefault(k, []).append(nm)
        except Exception:
            self._prefix_index = {}

        # if initial runner requested, try to pre-select
        if self._initial_runner and self._initial_runner in names:
            self._runner_var.set(self._initial_runner)
            self._on_runner_selected()

    def _find_results_workbook(self):
        return find_latest_results_workbook(session_config.output_dir)

    def _refresh_clubs(self):
        """Populate club list from the latest workbook (best-effort)."""
        results_path = self._find_results_workbook()
        clubs = []
        if results_path is None:
            self._club_combo["values"] = []
            return
        xl = self._get_xl(results_path)
        if xl is None:
            self._club_combo["values"] = []
            return
        # try summary sheet first for club/team column
        try:
            df = None
            for s in xl.sheet_names:
                n = s.lower()
                if "summary" in n or "standings" in n:
                    df = xl.parse(s)
                    break
            if df is None:
                df = xl.parse(xl.sheet_names[0])
        except Exception:
            self._club_combo["values"] = []
            return
        club_col = None
        for c in ("Club", "Team", "Affiliation"):
            if c in df.columns:
                club_col = c
                break
        if club_col is None:
            for c in df.columns:
                if str(c).lower() in ("club", "team", "affiliation"):
                    club_col = c
                    break
        if club_col is None:
            # fallback: scan race sheets for a 'Club' column
            found = set()
            for s in xl.sheet_names:
                try:
                    s_df = xl.parse(s)
                except Exception:
                    continue
                for c in s_df.columns:
                    if str(c).lower() in ("club", "team"):
                        vals = s_df[c].dropna().astype(str).tolist()
                        for v in vals:
                            found.add(v.strip())
            clubs = sorted([c for c in found if c])
        else:
            try:
                clubs = sorted(df[club_col].dropna().astype(str).str.strip().unique().tolist())
            except Exception:
                clubs = []
        # Filter out team-suffixed entries like 'Club -- A' or 'Club -- B'
        def _strip_team_suffix(name: str) -> str:
            if not name:
                return name
            # remove occurrences of ' -- A' or ' -- B'
            m = re.match(r"^(.*)\s+--\s+[AB]$", name)
            if m:
                return m.group(1).strip()
            return name

        cleaned = []
        seen = set()
        for c in clubs:
            s = _strip_team_suffix(c)
            if not s:
                continue
            key = s.lower()
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(s)

        self._club_combo["values"] = cleaned

    def _on_club_selected(self):
        club = self._club_var.get().strip()
        if not club:
            return
        # clear tree
        for iid in self._club_tree.get_children():
            self._club_tree.delete(iid)
        # clear per-race center widgets (both Team A and Team B)
        try:
            if hasattr(self, "_team_a_center"):
                for child in list(self._team_a_center.winfo_children()):
                    child.destroy()
            if hasattr(self, "_team_b_center"):
                for child in list(self._team_b_center.winfo_children()):
                    child.destroy()
        except Exception:
            pass
        # load workbook and try to find members in summary sheet
        results_path = self._find_results_workbook()
        if results_path is None:
            messagebox.showwarning("No workbook", "No standings workbook available.", parent=self)
            return
        xl = self._get_xl(results_path)
        if xl is None:
            messagebox.showerror("Read Failed", f"Failed to open workbook: {results_path.name}", parent=self)
            return
        members = []
        total_points = 0
        members_male_total = 0
        members_female_total = 0
        per_race = {}
        try:
            # try summary/standings for members and points
            summary = None
            for s in xl.sheet_names:
                n = s.lower()
                if "summary" in n or "standings" in n:
                    summary = xl.parse(s)
                    break
            if summary is None:
                summary = xl.parse(xl.sheet_names[0])
            # find name and club and points columns
            name_col = None
            club_col = None
            pts_col = None
            for c in summary.columns:
                cl = str(c).lower()
                if cl in ("name","runner","full name","fullname") and name_col is None:
                    name_col = c
                if cl in ("club","team","affiliation") and club_col is None:
                    club_col = c
                if cl in ("total points","points","score","total") and pts_col is None:
                    pts_col = c
            # detect gender column in summary if present
            gender_col = None
            for c in summary.columns:
                if str(c).strip().lower() in ("gender", "sex", "m/f", "male/female"):
                    gender_col = c
                    break
            if club_col is not None and name_col is not None:
                for _, r in summary.iterrows():
                    try:
                        rc = str(r.get(club_col, "")).strip()
                    except Exception:
                        rc = ""
                    if rc and rc.lower() == club.lower():
                        name = str(r.get(name_col, "")).strip()
                        pts = r.get(pts_col, None) if pts_col is not None else None
                        pts_s = ""
                        if pts is not None and pd.notna(pts) and str(pts).strip() != "":
                            try:
                                pts_s = str(int(float(pts)))
                                total_points += int(float(pts))
                                # add to gender subtotals if gender known
                                try:
                                    if gender_col is not None:
                                        gv = r.get(gender_col, "")
                                        if pd.notna(gv) and str(gv).strip():
                                            g0 = str(gv).strip()[0].upper()
                                            if g0 == 'M':
                                                members_male_total += int(float(pts))
                                            elif g0 == 'F':
                                                members_female_total += int(float(pts))
                                except Exception:
                                    pass
                            except Exception:
                                pts_s = str(pts)
                        catv = str(r.get("Category", r.get("Cat", r.get("Age", "")))).strip()
                        members.append((name, catv, pts_s, ""))
        except Exception:
            members = []
        # if no members found in summary, try to aggregate from race sheets
        if not members:
            found = {}
            for s in xl.sheet_names:
                try:
                    df = xl.parse(s)
                except Exception:
                    continue
                # find name/club columns
                name_c = None
                club_c = None
                pts_c = None
                cat_c = None
                for c in df.columns:
                    cl = str(c).lower()
                    if cl in ("name","runner","full name","fullname") and name_c is None:
                        name_c = c
                    if cl in ("club","team") and club_c is None:
                        club_c = c
                    if cl in ("points","score","total points","total") and pts_c is None:
                        pts_c = c
                    if cl in ("category","cat","age","grade") and cat_c is None:
                        cat_c = c
                if name_c is None or club_c is None:
                    continue
                # detect gender column in this sheet
                gender_c = None
                for c in df.columns:
                    cl2 = str(c).lower()
                    if cl2 in ("gender", "sex", "m/f", "male/female"):
                        gender_c = c
                        break
                for _, r in df.iterrows():
                    try:
                        rc = str(r.get(club_c, "")).strip()
                    except Exception:
                        rc = ""
                    if rc and rc.lower() == club.lower():
                        name = str(r.get(name_c, "")).strip()
                        pts = r.get(pts_c, None) if pts_c is not None else None
                        if name:
                            ent = found.setdefault(name, {"points":0, "races":0, "cat":""})
                            ent["races"] += 1
                            try:
                                if pts is not None and pd.notna(pts) and str(pts).strip() != "":
                                    ent["points"] += int(float(pts))
                                    total_points += int(float(pts))
                            except Exception:
                                pass
                            # capture category where available
                            try:
                                if cat_c is not None:
                                    catv = str(r.get(cat_c, "")).strip()
                                    if catv and not ent.get("cat"):
                                        ent["cat"] = catv
                            except Exception:
                                pass
                            # capture gender where available
                            try:
                                if gender_c is not None:
                                    gv = r.get(gender_c, "")
                                    if pd.notna(gv) and str(gv).strip():
                                        g0 = str(gv).strip()[0].upper()
                                        if not ent.get("gender"):
                                            ent["gender"] = g0
                            except Exception:
                                pass
                            # aggregate per-race stats
                            pr = per_race.setdefault(s, {"members":0, "points":0})
                            pr["members"] += 1
                            try:
                                if pts is not None and pd.notna(pts) and str(pts).strip() != "":
                                    pr["points"] += int(float(pts))
                            except Exception:
                                pass
            members = [(n, ent.get("cat",""), str(ent.get("points","")), str(ent.get("races",""))) for n, ent in found.items()]
            # compute member gender totals from aggregated found entries
            try:
                members_male_total = 0
                members_female_total = 0
                for n, ent in found.items():
                    try:
                        ptsv = int(ent.get("points", 0))
                    except Exception:
                        try:
                            ptsv = int(float(ent.get("points", 0)))
                        except Exception:
                            ptsv = 0
                    g = ent.get("gender", "").upper() if ent.get("gender") else ""
                    if g.startswith("M"):
                        members_male_total += ptsv
                    elif g.startswith("F"):
                        members_female_total += ptsv
            except Exception:
                pass

        # populate club tree
        for idx, m in enumerate(sorted(members, key=lambda x: (-(int(x[2]) if x[2].isdigit() else 0), x[0]))):
            tag = "even" if idx % 2 else "odd"
            self._club_tree.insert("", "end", values=m, tags=(tag,))

        # Build detailed per-race team entries (Team A / Team B) from race sheets
        try:
            race_name_map = {}
            try:
                if session_config.events_path and session_config.events_path.exists():
                    schedule = load_events(session_config.events_path)
                    for ev in schedule.events:
                        m = re.search(r"(\d+)", ev.race_ref or "")
                        if m:
                            try:
                                rn = int(m.group(1))
                                race_name_map[rn] = ev.event_name
                            except Exception:
                                continue
            except Exception:
                race_name_map = {}

            race_sheets = [s for s in xl.sheet_names if s.lower().startswith("race ")]
            race_sheets = sorted_race_sheet_names(xl) if race_sheets else []

            team_data = {}
            male_total = 0
            female_total = 0
            for sheet in race_sheets:
                try:
                    df = xl.parse(sheet)
                except Exception:
                    continue
                # detect useful columns
                name_c = None
                club_c = None
                team_c = None
                pos_c = None
                pts_c = None
                for c in df.columns:
                    cl = str(c).lower()
                    if cl in ("name","runner","full name","fullname") and name_c is None:
                        name_c = c
                    if cl in ("club","team","preferred club") and club_c is None:
                        club_c = c
                    if cl == "team" and team_c is None:
                        team_c = c
                    if cl in ("overall pos","overall position","position","pos") and pos_c is None:
                        pos_c = c
                    if cl in ("points","score","total points","team points") and pts_c is None:
                        pts_c = c
                if name_c is None or club_c is None:
                    continue
                a_list = []
                b_list = []
                a_pts = 0
                b_pts = 0
                for _, r in df.iterrows():
                    try:
                        rc = str(r.get(club_c, "")).strip()
                    except Exception:
                        rc = ""
                    if not rc or rc.lower() != club.lower():
                        continue
                    name = str(r.get(name_c, "")).strip()
                    pos = r.get(pos_c, "") if pos_c is not None else ""
                    pts = r.get(pts_c, None) if pts_c is not None else None
                    pts_s = ""
                    pts_num = None
                    try:
                        if pts is not None and pd.notna(pts) and str(pts).strip() != "":
                            pts_num = int(float(pts))
                            pts_s = str(pts_num)
                    except Exception:
                        pts_s = str(pts) if pts is not None else ""
                    team_val = ""
                    try:
                        if team_c is not None:
                            team_val = str(r.get(team_c, "")).strip().upper()
                        else:
                            # try to infer team from order or empty
                            team_val = ""
                    except Exception:
                        team_val = ""
                    # detect gender column if present
                    gender_c = None
                    for c in df.columns:
                        cl2 = str(c).lower()
                        if cl2 in ("gender", "sex", "m/f", "male/female"):
                            gender_c = c
                            break
                    gender_val = ""
                    try:
                        if gender_c is not None:
                            gv = r.get(gender_c, "")
                            if pd.notna(gv) and str(gv).strip():
                                gender_val = str(gv).strip()[0].upper()
                    except Exception:
                        gender_val = ""

                    entry = (name, str(pos), pts_s, gender_val)
                    if team_val and team_val.startswith("A"):
                        a_list.append(entry)
                        try:
                            a_pts += pts_num if pts_num is not None else 0
                        except Exception:
                            pass
                    elif team_val and team_val.startswith("B"):
                        b_list.append(entry)
                        try:
                            b_pts += pts_num if pts_num is not None else 0
                        except Exception:
                            pass
                    # accumulate male/female totals across all races for the club
                    try:
                        if gender_val and gender_val.upper().startswith("M"):
                            male_total += pts_num if pts_num is not None else 0
                        elif gender_val and gender_val.upper().startswith("F"):
                            female_total += pts_num if pts_num is not None else 0
                    except Exception:
                        pass
                if a_list or b_list:
                    rn = extract_race_number(sheet) or None
                    display = race_name_map.get(rn, sheet) if rn is not None else sheet
                    team_data[sheet] = {"display": display, "A": a_list, "B": b_list, "A_pts": a_pts, "B_pts": b_pts}

            # render into Team A and Team B centres with gender-separated sections and aligned columns
            if team_data:
                for sheet in sorted(team_data.keys()):
                    rec = team_data[sheet]
                    # Team A card
                    try:
                        a_card = tk.Frame(self._team_a_center, bg="#ffffff", bd=1, relief="solid")
                        a_card.pack(fill="x", pady=6)
                        header = tk.Frame(a_card, bg="#eaf3ff")
                        header.pack(fill="x")
                        header_lbl = tk.Label(header, text=f"{sheet} — {rec['display']} — {rec['A_pts']} pts", bg="#eaf3ff", font=("Segoe UI", 10, "bold"))
                        header_lbl.pack(side="left", fill="x", expand=True, padx=6, pady=4)
                        # details frame
                        a_details = tk.Frame(a_card, bg="#ffffff")
                        a_details.pack(fill="x", padx=6, pady=(4,6))
                        # configure grid columns: name (stretch), pos (center), pts (right)
                        a_details.grid_columnconfigure(0, weight=1)
                        a_details.grid_columnconfigure(1, weight=0)
                        a_details.grid_columnconfigure(2, weight=0)

                        entries = rec.get("A") or []
                        males = [p for p in entries if p[3] and str(p[3]).upper().startswith("M")]
                        females = [p for p in entries if p[3] and str(p[3]).upper().startswith("F")]
                        other = [p for p in entries if not p[3]]
                        row_idx = 0
                        if males:
                            tk.Label(a_details, text="Male", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=row_idx, column=0, columnspan=3, sticky="w", pady=(0,4))
                            row_idx += 1
                            # header row
                            tk.Label(a_details, text="Name", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w")
                            tk.Label(a_details, text="Pos", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=1)
                            tk.Label(a_details, text="Pts", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=2, sticky="e")
                            row_idx += 1
                            for person in males:
                                tk.Label(a_details, text=person[0], bg="#ffffff").grid(row=row_idx, column=0, sticky="w")
                                tk.Label(a_details, text=person[1], bg="#ffffff").grid(row=row_idx, column=1)
                                tk.Label(a_details, text=person[2], bg="#ffffff").grid(row=row_idx, column=2, sticky="e")
                                row_idx += 1
                        if females:
                            tk.Label(a_details, text="Female", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=row_idx, column=0, columnspan=3, sticky="w", pady=(8,4))
                            row_idx += 1
                            tk.Label(a_details, text="Name", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w")
                            tk.Label(a_details, text="Pos", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=1)
                            tk.Label(a_details, text="Pts", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=2, sticky="e")
                            row_idx += 1
                            for person in females:
                                tk.Label(a_details, text=person[0], bg="#ffffff").grid(row=row_idx, column=0, sticky="w")
                                tk.Label(a_details, text=person[1], bg="#ffffff").grid(row=row_idx, column=1)
                                tk.Label(a_details, text=person[2], bg="#ffffff").grid(row=row_idx, column=2, sticky="e")
                                row_idx += 1
                        if not males and not females:
                            # fall back to showing others or a placeholder
                            if other:
                                tk.Label(a_details, text="Name", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w")
                                tk.Label(a_details, text="Pos", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=1)
                                tk.Label(a_details, text="Pts", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=2, sticky="e")
                                row_idx += 1
                                for person in other:
                                    tk.Label(a_details, text=person[0], bg="#ffffff").grid(row=row_idx, column=0, sticky="w")
                                    tk.Label(a_details, text=person[1], bg="#ffffff").grid(row=row_idx, column=1)
                                    tk.Label(a_details, text=person[2], bg="#ffffff").grid(row=row_idx, column=2, sticky="e")
                                    row_idx += 1
                            else:
                                tk.Label(a_details, text="No Team A runners for this race", bg="#ffffff", fg="#666").grid(row=row_idx, column=0, columnspan=3)
                    except Exception:
                        pass

                    # Team B card
                    try:
                        b_card = tk.Frame(self._team_b_center, bg="#ffffff", bd=1, relief="solid")
                        b_card.pack(fill="x", pady=6)
                        header = tk.Frame(b_card, bg="#e8f7ea")
                        header.pack(fill="x")
                        header_lbl = tk.Label(header, text=f"{sheet} — {rec['display']} — {rec['B_pts']} pts", bg="#e8f7ea", font=("Segoe UI", 10, "bold"))
                        header_lbl.pack(side="left", fill="x", expand=True, padx=6, pady=4)
                        b_details = tk.Frame(b_card, bg="#ffffff")
                        b_details.pack(fill="x", padx=6, pady=(4,6))
                        b_details.grid_columnconfigure(0, weight=1)
                        b_details.grid_columnconfigure(1, weight=0)
                        b_details.grid_columnconfigure(2, weight=0)

                        entries = rec.get("B") or []
                        males = [p for p in entries if p[3] and str(p[3]).upper().startswith("M")]
                        females = [p for p in entries if p[3] and str(p[3]).upper().startswith("F")]
                        other = [p for p in entries if not p[3]]
                        row_idx = 0
                        if males:
                            tk.Label(b_details, text="Male", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=row_idx, column=0, columnspan=3, sticky="w", pady=(0,4))
                            row_idx += 1
                            tk.Label(b_details, text="Name", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w")
                            tk.Label(b_details, text="Pos", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=1)
                            tk.Label(b_details, text="Pts", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=2, sticky="e")
                            row_idx += 1
                            for person in males:
                                tk.Label(b_details, text=person[0], bg="#ffffff").grid(row=row_idx, column=0, sticky="w")
                                tk.Label(b_details, text=person[1], bg="#ffffff").grid(row=row_idx, column=1)
                                tk.Label(b_details, text=person[2], bg="#ffffff").grid(row=row_idx, column=2, sticky="e")
                                row_idx += 1
                        if females:
                            tk.Label(b_details, text="Female", bg="#ffffff", font=("Segoe UI", 10, "bold")).grid(row=row_idx, column=0, columnspan=3, sticky="w", pady=(8,4))
                            row_idx += 1
                            tk.Label(b_details, text="Name", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w")
                            tk.Label(b_details, text="Pos", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=1)
                            tk.Label(b_details, text="Pts", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=2, sticky="e")
                            row_idx += 1
                            for person in females:
                                tk.Label(b_details, text=person[0], bg="#ffffff").grid(row=row_idx, column=0, sticky="w")
                                tk.Label(b_details, text=person[1], bg="#ffffff").grid(row=row_idx, column=1)
                                tk.Label(b_details, text=person[2], bg="#ffffff").grid(row=row_idx, column=2, sticky="e")
                                row_idx += 1
                        if not males and not females:
                            if other:
                                tk.Label(b_details, text="Name", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=0, sticky="w")
                                tk.Label(b_details, text="Pos", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=1)
                                tk.Label(b_details, text="Pts", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=row_idx, column=2, sticky="e")
                                row_idx += 1
                                for person in other:
                                    tk.Label(b_details, text=person[0], bg="#ffffff").grid(row=row_idx, column=0, sticky="w")
                                    tk.Label(b_details, text=person[1], bg="#ffffff").grid(row=row_idx, column=1)
                                    tk.Label(b_details, text=person[2], bg="#ffffff").grid(row=row_idx, column=2, sticky="e")
                                    row_idx += 1
                            else:
                                tk.Label(b_details, text="No Team B runners for this race", bg="#ffffff", fg="#666").grid(row=row_idx, column=0, columnspan=3)
                    except Exception:
                        pass
        except Exception:
            pass
        self._club_vars["ClubName"].set(club)
        self._club_vars["Members"].set(str(len(members)))
        self._club_vars["TotalPoints"].set(str(total_points))
        # attempt to read Div 1/Div 2 sheets to fill Team A/B totals
        team_a_str = "Team A: n/a"
        team_b_str = "Team B: n/a"
        try:
            # find division sheets
            div1 = None
            div2 = None
            for s in xl.sheet_names:
                if s.lower() == "div 1":
                    div1 = xl.parse(s)
                if s.lower() == "div 2":
                    div2 = xl.parse(s)
            def _find_team_info(df, team_disp, sheet_name):
                """Return (points:int|None, division_name:str|None, position:str|None)"""
                if df is None:
                    return (None, None, None)
                # find club column
                club_col = None
                for c in df.columns:
                    if str(c).strip().lower() == "club":
                        club_col = c
                        break
                if club_col is None:
                    for c in df.columns:
                        if str(c).strip().lower() in ("club", "team", "club name"):
                            club_col = c
                            break
                if club_col is None:
                    return (None, None, None)
                try:
                    mask = df[club_col].astype(str).fillna("").str.strip() == team_disp
                    rows = df[mask]
                    if rows.empty:
                        return (None, None, None)
                    row = rows.iloc[0]
                    pts = None
                    for candidate in ("Total Points", "TotalPoints", "Points", "Total"):
                        if candidate in df.columns:
                            val = row.get(candidate)
                            if pd.notna(val) and str(val).strip() != "":
                                try:
                                    pts = int(float(val))
                                    break
                                except Exception:
                                    pts = None
                    # position: prefer explicit column if present
                    pos = None
                    for pc in ("Position", "Pos", "Place", "Rank"):
                        if pc in df.columns:
                            pv = row.get(pc)
                            if pd.notna(pv) and str(pv).strip() != "":
                                try:
                                    pos = str(int(float(pv)))
                                except Exception:
                                    pos = str(pv)
                                break
                    if pos is None:
                        # fallback: compute position by order among non-empty club rows
                        try:
                            nonempty = df[club_col].astype(str).fillna("").str.strip()
                            idxs = nonempty[nonempty != ""].index.tolist()
                            pos = str(idxs.index(row.name) + 1) if row.name in idxs else None
                        except Exception:
                            pos = None
                    return (pts, sheet_name, pos)
                except Exception:
                    return (None, None, None)

            # team display names use 'Club -- A' format per TeamSeasonRecord.display_name
            team_a_disp = f"{club} -- A"
            team_b_disp = f"{club} -- B"
            # get team points and division/position info
            ta_pts, ta_div, ta_pos = (None, None, None)
            tb_pts, tb_div, tb_pos = (None, None, None)
            if div1 is not None:
                t = _find_team_info(div1, team_a_disp, "Div 1")
                if t[0] is not None:
                    ta_pts, ta_div, ta_pos = t
                t2 = _find_team_info(div1, team_b_disp, "Div 1")
                if t2[0] is not None:
                    tb_pts, tb_div, tb_pos = t2
            if (ta_pts is None or ta_div is None) and div2 is not None:
                t = _find_team_info(div2, team_a_disp, "Div 2")
                if t[0] is not None:
                    ta_pts, ta_div, ta_pos = t
            if (tb_pts is None or tb_div is None) and div2 is not None:
                t = _find_team_info(div2, team_b_disp, "Div 2")
                if t[0] is not None:
                    tb_pts, tb_div, tb_pos = t
            if ta_pts is not None:
                team_a_str = f"Points: {ta_pts}"
            if tb_pts is not None:
                team_b_str = f"Points: {tb_pts}"
            # helper to sum R# Male/Female columns from a division dataframe for a team
            def _sum_gender_from_div(df, team_disp, gender_letter):
                if df is None:
                    return 0
                # find club column
                club_col = None
                for c in df.columns:
                    if str(c).strip().lower() == "club":
                        club_col = c
                        break
                if club_col is None:
                    for c in df.columns:
                        if str(c).strip().lower() in ("club", "team", "club name"):
                            club_col = c
                            break
                if club_col is None:
                    return 0
                try:
                    mask = df[club_col].astype(str).fillna("").str.strip() == team_disp
                    rows = df[mask]
                    if rows.empty:
                        return 0
                    row = rows.iloc[0]
                    total = 0
                    for c in df.columns:
                        cname = str(c).lower()
                        # accept common variants: 'male' or 'men' for males; 'female' or 'women' for females
                        if gender_letter == 'M':
                            keywords = ("male", "men")
                        else:
                            keywords = ("female", "women")
                        if any(k in cname for k in keywords) and re.search(r"\d", cname):
                            try:
                                v = row.get(c)
                                if pd.notna(v) and str(v).strip() != "":
                                    total += int(float(v))
                            except Exception:
                                pass
                    return total
                except Exception:
                    return 0

            # compute male/female totals by summing R# Male/Female columns across division sheets for this club's teams
            male_total = 0
            female_total = 0
            try:
                for df in (div1, div2):
                    if df is None:
                        continue
                    male_total += _sum_gender_from_div(df, team_a_disp, 'M')
                    male_total += _sum_gender_from_div(df, team_b_disp, 'M')
                    female_total += _sum_gender_from_div(df, team_a_disp, 'F')
                    female_total += _sum_gender_from_div(df, team_b_disp, 'F')
            except Exception:
                male_total = male_total or 0
                female_total = female_total or 0
        except Exception:
            team_a_str = "Team A: n/a"
            team_b_str = "Team B: n/a"

        try:
            self._club_races_vars["ClubName"].set(club)
            self._club_races_vars["Members"].set(str(len(members)))
            self._club_races_vars["TotalPoints"].set(str(total_points))
            self._club_races_vars["TeamA"].set(team_a_str)
            self._club_races_vars["TeamB"].set(team_b_str)
            # set male/female aggregated scores computed from race sheets
            try:
                self._club_races_vars["MaleScore"].set(str(male_total))
            except Exception:
                self._club_races_vars["MaleScore"].set("")
            try:
                self._club_races_vars["FemaleScore"].set(str(female_total))
            except Exception:
                self._club_races_vars["FemaleScore"].set("")
            # Position display: show position only; show division in the Team label
            ta_pos_str = "Position: n/a"
            tb_pos_str = "Position: n/a"
            try:
                if ta_pos:
                    ta_pos_str = f"Position: #{ta_pos}"
                if tb_pos:
                    tb_pos_str = f"Position: #{tb_pos}"
            except Exception:
                pass
            self._club_races_vars["TeamA_Pos"].set(ta_pos_str)
            self._club_races_vars["TeamB_Pos"].set(tb_pos_str)
            # set Team label to include division where available
            try:
                if ta_div:
                    self._club_races_vars["TeamA_Label"].set(f"Team A ({ta_div})")
                else:
                    self._club_races_vars["TeamA_Label"].set("Team A")
                if tb_div:
                    self._club_races_vars["TeamB_Label"].set(f"Team B ({tb_div})")
                else:
                    self._club_races_vars["TeamB_Label"].set("Team B")
            except Exception:
                pass
            # Compose short list summaries for Team A and Team B showing position and points
            try:
                ta_pos_display = ta_pos_str if ta_pos is not None else "Position: n/a"
                tb_pos_display = tb_pos_str if tb_pos is not None else "Position: n/a"
                ta_points_display = team_a_str if team_a_str else "Points: n/a"
                tb_points_display = team_b_str if team_b_str else "Points: n/a"
                self._club_races_vars["TeamA_Summary"].set(f"- {ta_pos_display}\n- {ta_points_display}")
                self._club_races_vars["TeamB_Summary"].set(f"- {tb_pos_display}\n- {tb_points_display}")
            except Exception:
                try:
                    self._club_races_vars["TeamA_Summary"].set("")
                    self._club_races_vars["TeamB_Summary"].set("")
                except Exception:
                    pass
        except Exception:
            pass

        # Attempt to load user-provided header image for Club - Races
        try:
            from PIL import Image, ImageTk
            img_path = Path.cwd() / "images" / "club_races_panel.png"
            if img_path.exists():
                try:
                    img = Image.open(img_path)
                    # scale to width ~260 to fit the side panel while keeping aspect
                    w, h = img.size
                    target_w = 260
                    if w > target_w:
                        ratio = target_w / float(w)
                        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
                    self._club_races_photo = ImageTk.PhotoImage(img)
                    self._club_races_image_label.config(image=self._club_races_photo)
                except Exception:
                    # image failed to load — leave placeholder
                    pass
        except Exception:
            # PIL not available; skip image loading
            pass

    def _on_runner_selected(self):
        name = self._runner_var.get().strip()
        if not name:
            return
        results_path = self._find_results_workbook()
        if results_path is None:
            messagebox.showwarning("No workbook", "No standings workbook available.", parent=self)
            return
        xl = self._get_xl(results_path)
        if xl is None:
            messagebox.showerror("Read Failed", f"Failed to open workbook: {results_path.name}", parent=self)
            return

        # Prefer Male/Female sheets for individual standings (common layout),
        # otherwise fall back to Summary/Standings.
        summary_df = pd.DataFrame()
        runner_row = None
        try:
            # look for male/female-specific sheets first
            chosen_sheet = None
            for s in xl.sheet_names:
                n = s.lower()
                if "male" in n or "female" in n:
                    chosen_sheet = s
                    # prefer exact 'Male' or 'Female' if present
                    if n.strip() in ("male", "female"):
                        break
            if chosen_sheet:
                summary_df = xl.parse(chosen_sheet)
            else:
                # fallback to summary/standings
                chosen = None
                for s in xl.sheet_names:
                    n = s.lower()
                    if "summary" in n or "standings" in n:
                        chosen = s
                        break
                if chosen:
                    summary_df = xl.parse(chosen)
                else:
                    summary_df = xl.parse(xl.sheet_names[0])
        except Exception:
            summary_df = pd.DataFrame()

        # Find name column in the selected summary_df
        name_col = None
        for candidate in ("Name", "Runner", "Full Name", "fullname"):
            if candidate in summary_df.columns:
                name_col = candidate
                break
        if name_col is None:
            for c in summary_df.columns:
                if str(c).lower() in ("name", "runner", "full name", "fullname"):
                    name_col = c
                    break

        if name_col is not None:
            # Use raw name mapping where available to improve match chances
            raw_name = self._name_map.get(name.lower(), name)
            search_key = _norm_for_compare(raw_name)

            try:
                series = summary_df[name_col].astype(str).fillna("")
                rev_search = " ".join(reversed(search_key.split()))
                def _match_val(s):
                    ns = _norm_for_compare(s)
                    return ns == search_key or ns == rev_search
                mask = series.apply(lambda s: _match_val(s))
                matches = summary_df[mask]
            except Exception:
                matches = summary_df[summary_df[name_col].astype(str).str.strip().str.lower() == raw_name.lower()]

            if not matches.empty:
                runner_row = matches.iloc[0]

        # Fill summary card from the individual sheet row when available
        club = ""
        category = ""
        points = ""
        if runner_row is not None:
            # try common column names; prefer 'Total Points' for overall points
            club = runner_row.get("Club", runner_row.get("Team", ""))
            category = runner_row.get("Category", runner_row.get("Cat", runner_row.get("Age", "")))
            # Prefer 'Total Points' first, then Points/Total/Score
            pts_candidates = ["Total Points", "Points", "Total", "Score"]
            pts_val = None
            for c in pts_candidates:
                if c in runner_row.index:
                    pts_val = runner_row.get(c)
                    if pd.notna(pts_val) and str(pts_val).strip() != "":
                        break
            if pts_val is not None and pd.notna(pts_val) and str(pts_val).strip() != "":
                try:
                    points = str(int(float(pts_val)))
                except Exception:
                    points = str(pts_val)

        # Name and Club shown separately in the left summary
        self._summary_vars["NameClub"].set(name)
        self._summary_vars["Club"].set(club)
        self._summary_vars["Category"].set(category)
        self._summary_vars["Points"].set(points)

        # Clear details
        for i in self._tree.get_children():
            self._tree.delete(i)

        # Build race number -> official event name mapping (if events file available)
        race_name_map: dict[int, str] = {}
        try:
            if session_config.events_path and session_config.events_path.exists():
                schedule = load_events(session_config.events_path)
                for ev in schedule.events:
                    # try to extract an integer race number from the RaceRef
                    m = re.search(r"(\d+)", ev.race_ref or "")
                    if m:
                        try:
                            race_number = int(m.group(1))
                            race_name_map[race_number] = ev.event_name
                        except Exception:
                            continue
        except Exception:
            # non-fatal; proceed without event name mapping
            race_name_map = {}

        # Scan Race sheets for per-race details
        race_sheets = [s for s in xl.sheet_names if s.startswith("Race ")]
        race_sheets = sorted_race_sheet_names(xl) if race_sheets else []
        details = []
        for sheet in race_sheets:
            try:
                df = xl.parse(sheet)
            except Exception:
                continue
            # try common name columns
            candidate_cols = [c for c in df.columns if str(c).lower() in ("name","runner","full name","fullname")]
            if candidate_cols:
                col = candidate_cols[0]
                # use relaxed normalization when matching across race sheets
                raw_name = self._name_map.get(name.lower(), name)
                search_key = _norm_for_compare(raw_name)
                try:
                    series = df[col].astype(str).fillna("")
                    rev_search = " ".join(reversed(search_key.split()))
                    def _match_val(s):
                        ns = _norm_for_compare(s)
                        return ns == search_key or ns == rev_search
                    row = df[series.apply(lambda s: _match_val(s))]
                except Exception:
                    row = df[df[col].astype(str).str.strip().str.lower() == raw_name.lower()]
                if not row.empty:
                    r = row.iloc[0]
                    time = r.get("Time", "")
                    # race-specific club (the club they ran under for that race)
                    race_club = r.get("Club", r.get("Team", ""))
                    # prefer per-race points where available
                    pts = None
                    for c in ("Points", "Score", "Total Points"):
                        if c in r.index:
                            pts = r.get(c)
                            if pd.notna(pts) and str(pts).strip() != "":
                                break
                    # preserve raw time value for comparisons, then format for display
                    tval = r.get("Time", "")
                    secs = _time_to_seconds(tval)
                    # format time into H:MM:SS.0 if possible
                    time = _format_time_str(tval)
                    # format points as integer string
                    if pts is None or not pd.notna(pts):
                        pts_s = ""
                    else:
                        try:
                            pts_s = str(int(float(pts)))
                        except Exception:
                            pts_s = str(pts)
                    # Prefer official event name when available
                    rn = extract_race_number(sheet) or None
                    display_name = race_name_map.get(rn, sheet) if rn is not None else sheet
                    # derive a short distance label from the event name
                    dist_label = _extract_distance_label(display_name)
                    details.append((display_name, race_club, time, pts_s, secs, dist_label))

        # keep last details for inspection/testing
        self._last_details = details
        for idx, entry in enumerate(details):
            tag = "even" if idx % 2 else "odd"
            # tree shows Race, Club, Time, Points (first four items)
            self._tree.insert("", "end", values=entry[0:4], tags=(tag,))

        # compute best times per distance and populate the right-hand best_tree
        best: dict[str, tuple[float, str, str]] = {}
        for d in details:
            display_name, race_club, time_s, pts_s, secs, dist_label = d
            if secs is None:
                continue
            cur = best.get(dist_label)
            if cur is None or secs < cur[0]:
                try:
                    best_time_str = _format_time_str(pd.to_timedelta(secs, unit='s'))
                except Exception:
                    best_time_str = _format_time_str(secs)
                best[dist_label] = (secs, best_time_str, display_name)

        # clear and insert into best_tree
        for iid in self._best_tree.get_children():
            self._best_tree.delete(iid)
        for idx, (dist, (secs, time_s, ev)) in enumerate(sorted(best.items(), key=lambda x: x[0])):
            tag = "even" if idx % 2 else "odd"
            self._best_tree.insert("", "end", values=(dist, time_s, ev), tags=(tag,))

        self._summary_vars["Races"].set(str(len(details)))

    def select_runner(self, runner_name: str) -> bool:
        if not runner_name:
            return False
        # ensure names are loaded
        self._refresh_names()
        try:
            self._runner_var.set(runner_name)
            self._on_runner_selected()
            return True
        except Exception:
            return False

    def _copy_details(self):
        """Copy action: prefer to capture the left summary card as an image.
        Fallbacks:
        - If pywin32 is available, place image on Windows clipboard.
        - Else save PNG to a temp file and copy the file path to clipboard.
        - If imaging unavailable, fall back to TSV of rows (original behaviour).
        """
        # attempt to capture the summary frame as an image
        try:
            from PIL import ImageGrab, Image
        except Exception:
            ImageGrab = None

        if ImageGrab and hasattr(self, "_summary_frame"):
            try:
                # ensure widget updated
                self.update_idletasks()
                # prefer capturing the Runner tab frame if available,
                # else fall back to the summary frame, else the whole toplevel
                target = getattr(self, "_runner_frame", None) or getattr(self, "_summary_frame", None) or self.winfo_toplevel()
                x1 = target.winfo_rootx()
                y1 = target.winfo_rooty()
                x2 = x1 + target.winfo_width()
                y2 = y1 + target.winfo_height()
                img = ImageGrab.grab(bbox=(x1, y1, x2, y2))

                # try copying image to clipboard using win32clipboard if available
                try:
                    import win32clipboard
                    from PIL import Image
                    output = io.BytesIO()
                    # Windows expects DIB for CF_DIB: save BMP and strip header
                    img.convert("RGB").save(output, "BMP")
                    data = output.getvalue()[14:]
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                    win32clipboard.CloseClipboard()
                    messagebox.showinfo("Screenshot copied", "Runner card image copied to clipboard.", parent=self)
                    return
                except Exception:
                    # fallback: save to temp file and copy path
                    try:
                        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                        img.save(tmp.name, "PNG")
                        tmp.close()
                        self.clipboard_clear()
                        self.clipboard_append(tmp.name)
                        self.update()
                        messagebox.showinfo("Saved screenshot", f"Saved runner card image to {tmp.name} (path copied to clipboard).", parent=self)
                        return
                    except Exception as e:
                        # fall through to TSV fallback
                        print("screenshot save failed:", e)
            except Exception:
                # non-fatal: fall back to TSV below
                pass
        # TSV fallback (original behaviour)
        rows = []
        cols = ["Race", "Club", "Time", "Points"]
        for iid in self._tree.get_children():
            vals = self._tree.item(iid).get("values", [])
            rows.append([str(v) for v in vals])
        if not rows:
            messagebox.showwarning("No Data", "No runner details to copy.", parent=self)
            return
        try:
            tsv_lines = ["\t".join(cols)]
            for r in rows:
                tsv_lines.append("\t".join(r))
            tsv = "\n".join(tsv_lines)
            # copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(tsv)
            self.update()
            messagebox.showinfo("Copied", f"Copied {len(rows)} rows to clipboard.", parent=self)
        except Exception as exc:
            messagebox.showerror("Copy Failed", f"Failed to copy to clipboard: {exc}", parent=self)
