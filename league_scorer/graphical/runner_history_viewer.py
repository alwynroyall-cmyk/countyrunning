import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from pathlib import Path

import openpyxl
import pandas as pd

from ..common_files import race_discovery_exclusions
from ..manual_data_audit import log_manual_data_changes
from ..manual_edit_service import resolve_runner_field_across_files
from ..normalisation import parse_time_to_seconds
from ..race_processor import extract_race_number
from ..session_config import config as session_config
from ..source_loader import discover_race_files
from .dashboard import WRRL_GREEN, WRRL_LIGHT, WRRL_NAVY, WRRL_WHITE
from .results_workbook import find_latest_results_workbook, sorted_race_sheet_names

_FORMULA_PREFIXES = frozenset(("=", "+", "-", "@"))


def _sanitise_df_for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix formula-start characters in text cells to prevent spreadsheet injection."""
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(
                lambda v: ("'" + v) if isinstance(v, str) and v and v[0] in _FORMULA_PREFIXES else v
            )
    return df


class RunnerHistoryPanel(tk.Frame):
    """View one runner's results across all race sheets in the latest results workbook."""

    def __init__(self, parent, back_callback=None, initial_runner: str | None = None):
        super().__init__(parent, bg=WRRL_LIGHT)
        self._back_callback = back_callback
        self._initial_runner = initial_runner
        self._style = ttk.Style(self)
        self._runner_name_map = {}
        self._all_runner_names: list[str] = []
        self._club_map = {}
        self._all_clubs: list[str] = []
        self._latest_df: pd.DataFrame | None = None
        self._lookup_mode = "runner"  # "runner" or "club"
        self._wb_cache: dict = {}
        self._configure_styles()
        self._build_ui()
        self._load_runner_options()

    def _configure_styles(self):
        self._style.configure(
            "RunnerHistory.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#22313f",
            rowheight=28,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self._style.configure(
            "RunnerHistory.Treeview.Heading",
            background="#dbe5ee",
            foreground="#22313f",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        self._style.map(
            "RunnerHistory.Treeview",
            background=[("selected", "#8fb3d1")],
            foreground=[("selected", "#102030")],
        )

    def _build_ui(self):
        header = tk.Frame(self, bg=WRRL_NAVY, padx=14, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Runner Results Across Races",
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

        controls = tk.Frame(self, bg=WRRL_LIGHT, padx=14, pady=10)
        controls.pack(fill="x")

        # Removed mode selector (runner/club radio buttons)

        self._runner_var = tk.StringVar()
        self._runner_combo = ttk.Combobox(
            controls,
            textvariable=self._runner_var,
            state="normal",
            width=48,
        )
        self._runner_combo.pack(side="left")
        self._runner_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_runner_history())
        self._runner_combo.bind("<KeyRelease>", self._on_runner_typed)
        self._runner_combo.bind("<Return>", self._on_runner_enter)
        self._runner_combo_label = tk.Label(
            controls,
            text="Runner:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        )
        self._runner_combo_label.pack(side="left", padx=(0, 8), before=self._runner_combo)

        # Removed club combobox and label

        tk.Button(
            controls,
            text="Refresh",
            font=("Segoe UI", 10),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._load_runner_options,
        ).pack(side="left", padx=(8, 0))

        self._summary_var = tk.StringVar(value="")
        tk.Label(
            self,
            textvariable=self._summary_var,
            font=("Segoe UI", 9, "italic"),
            bg=WRRL_LIGHT,
            fg="#666666",
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 6))

        export_bar = tk.Frame(self, bg=WRRL_LIGHT, padx=14, pady=0)
        export_bar.pack(fill="x", padx=0, pady=(0, 8))

        tk.Button(
            export_bar,
            text="Copy Results",
            font=("Segoe UI", 9),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._copy_results,
        ).pack(side="left")

        tk.Button(
            export_bar,
            text="Export CSV",
            font=("Segoe UI", 9),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._export_results_csv,
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            export_bar,
            text="Export Excel",
            font=("Segoe UI", 9),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._export_results_excel,
        ).pack(side="left", padx=(8, 0))

        # Resolve controls (shown only when conflicting values exist)
        self._resolve_frame = tk.Frame(self, bg=WRRL_LIGHT, padx=14, pady=6)

        self._club_resolve_row = tk.Frame(self._resolve_frame, bg=WRRL_LIGHT)
        tk.Label(
            self._club_resolve_row,
            text="Resolve Club:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))
        self._club_resolve_var = tk.StringVar()
        self._club_resolve_combo = ttk.Combobox(
            self._club_resolve_row,
            textvariable=self._club_resolve_var,
            state="readonly",
            width=28,
        )
        self._club_resolve_combo.pack(side="left")
        tk.Button(
            self._club_resolve_row,
            text="Apply Club To All Inputs",
            font=("Segoe UI", 9, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
            command=self._resolve_club_across_inputs,
        ).pack(side="left", padx=(8, 0))

        self._gender_resolve_row = tk.Frame(self._resolve_frame, bg=WRRL_LIGHT)
        tk.Label(
            self._gender_resolve_row,
            text="Resolve Gender:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))
        self._gender_resolve_var = tk.StringVar()
        self._gender_resolve_combo = ttk.Combobox(
            self._gender_resolve_row,
            textvariable=self._gender_resolve_var,
            state="readonly",
            width=10,
        )
        self._gender_resolve_combo.pack(side="left")
        tk.Button(
            self._gender_resolve_row,
            text="Apply Gender To All Inputs",
            font=("Segoe UI", 9, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
            command=self._resolve_gender_across_inputs,
        ).pack(side="left", padx=(8, 0))

        self._category_resolve_row = tk.Frame(self._resolve_frame, bg=WRRL_LIGHT)
        tk.Label(
            self._category_resolve_row,
            text="Resolve Category:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))
        self._category_resolve_var = tk.StringVar()
        self._category_resolve_combo = ttk.Combobox(
            self._category_resolve_row,
            textvariable=self._category_resolve_var,
            state="readonly",
            width=28,
        )
        self._category_resolve_combo.pack(side="left")
        tk.Button(
            self._category_resolve_row,
            text="Apply Category To All Inputs",
            font=("Segoe UI", 9, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
            command=self._resolve_category_across_inputs,
        ).pack(side="left", padx=(8, 0))

        self._time_resolve_row = tk.Frame(self._resolve_frame, bg=WRRL_LIGHT)
        tk.Label(
            self._time_resolve_row,
            text="Fix QRY Time:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))
        self._time_resolve_var = tk.StringVar()
        self._time_resolve_entry = ttk.Entry(
            self._time_resolve_row,
            textvariable=self._time_resolve_var,
            width=14,
        )
        self._time_resolve_entry.pack(side="left")
        tk.Label(
            self._time_resolve_row,
            text="Use hh:mm:ss",
            font=("Segoe UI", 9, "italic"),
            bg=WRRL_LIGHT,
            fg="#666666",
        ).pack(side="left", padx=(8, 0))
        tk.Button(
            self._time_resolve_row,
            text="Apply Time To All Inputs",
            font=("Segoe UI", 9, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
            command=self._resolve_time_across_inputs,
        ).pack(side="left", padx=(10, 0))

        table_frame = tk.Frame(self, bg=WRRL_LIGHT)
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self._tree = ttk.Treeview(table_frame, show="headings", style="RunnerHistory.Treeview")
        self._tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self._tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self._tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self._tree.tag_configure("odd", background="#ffffff")
        self._tree.tag_configure("even", background="#eef3f8")

    def select_runner(self, runner_name: str) -> bool:
        """Select a runner in runner mode and load their history.

        Returns True when the runner was found (exact or prefix match), else False.
        """
        target_text = str(runner_name or "").strip()
        if not target_text:
            return False

        # Mode selector removed; just select runner and load history

        exact = next((n for n in self._all_runner_names if n.lower() == target_text.lower()), None)
        if exact is None:
            exact = next((n for n in self._all_runner_names if n.lower().startswith(target_text.lower())), None)
        if exact is None:
            return False

        self._runner_var.set(exact)
        self._load_runner_history()
        return True

    def _find_results_workbook(self):
        return find_latest_results_workbook(session_config.output_dir)

    def _ensure_workbook_cache(self, *, force: bool = False) -> bool:
        """Parse and cache all race sheet DataFrames from the latest results workbook.

        Returns True if data is available, False if no workbook was found.
        Raises RuntimeError on I/O failure.
        """
        workbook = self._find_results_workbook()
        if workbook is None:
            self._wb_cache = {}
            return False
        try:
            mtime = workbook.stat().st_mtime
        except OSError:
            self._wb_cache = {}
            return False
        if not force and (
            self._wb_cache.get("path") == workbook
            and abs(self._wb_cache.get("mtime", -1) - mtime) < 0.01
        ):
            return True  # cache is still current
        try:
            xl = pd.ExcelFile(workbook)
            try:
                race_sheets = sorted_race_sheet_names(xl)
                sheets: dict[str, pd.DataFrame] = {}
                name_map: dict[str, str] = {}
                club_map: dict[str, str] = {}
                for sheet in race_sheets:
                    df = xl.parse(sheet)
                    sheets[sheet] = df
                    if "Name" in df.columns:
                        for value in df["Name"].dropna():
                            name = str(value).strip()
                            if name:
                                name_map.setdefault(name.lower(), name)
                    if "Club" in df.columns:
                        for value in df["Club"].dropna():
                            club = str(value).strip()
                            if club:
                                club_map.setdefault(club.lower(), club)
            finally:
                xl.close()
        except Exception as exc:
            self._wb_cache = {}
            raise RuntimeError(f"Failed reading results workbook: {exc}") from exc
        self._wb_cache = {
            "path": workbook,
            "mtime": mtime,
            "race_sheets": race_sheets,
            "sheets": sheets,
            "name_map": name_map,
            "club_map": club_map,
        }
        return True

    def _load_runner_options(self):
        # Force-rebuild cache so an explicit Refresh always re-reads the file.
        self._wb_cache = {}
        try:
            available = self._ensure_workbook_cache(force=True)
        except RuntimeError as exc:
            self._runner_combo["values"] = []
            self._runner_var.set("")
            self._show_message(str(exc))
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        if not available:
            self._runner_combo["values"] = []
            self._runner_var.set("")
            self._show_message("No results workbook found in the active output folder.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        names = self._wb_cache["name_map"]
        self._runner_name_map = names
        sorted_names = sorted(names.values(), key=lambda n: n.lower())
        self._all_runner_names = sorted_names
        self._runner_combo["values"] = sorted_names

        if sorted_names:
            current = self._runner_var.get()
            # Honour initial_runner requested from external callers (e.g. Issue Review panel)
            if self._initial_runner:
                # Try exact match first, then case-insensitive via the name map
                initial = self._initial_runner
                self._initial_runner = None  # consume once
                if initial in sorted_names:
                    self._runner_var.set(initial)
                else:
                    matched = self._runner_name_map.get(initial.lower())
                    if matched:
                        self._runner_var.set(matched)
                    elif current not in sorted_names:
                        self._runner_var.set(sorted_names[0])
            elif current not in sorted_names:
                self._runner_var.set(sorted_names[0])
            if self._lookup_mode == "runner":
                self._load_runner_history()
        else:
            self._runner_var.set("")
            self._show_message("No runner names found in race sheets.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)

    def _on_mode_change(self):
        self._lookup_mode = self._mode_var.get()
        is_runner = self._lookup_mode == "runner"

        if is_runner:
            self._club_frame.pack_forget()
            self._runner_combo_label.pack(side="left", padx=(0, 8), before=self._runner_combo)
            self._runner_combo.pack(side="left")
            self._load_runner_history()
        else:
            self._runner_combo.pack_forget()
            self._runner_combo_label.pack_forget()
            self._club_frame.pack(side="left")
            self._load_club_results()

    def _load_club_results(self):
        selected = self._club_var.get().strip()
        if not selected:
            self._show_message("No club selected.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        try:
            available = self._ensure_workbook_cache()
        except RuntimeError as exc:
            self._show_message(str(exc))
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        if not available:
            self._show_message("No results workbook found.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        club_key = selected.lower()
        rows = []

        try:
            for sheet in self._wb_cache["race_sheets"]:
                race_num = extract_race_number(sheet) or ""
                df = self._wb_cache["sheets"][sheet].fillna("")
                if "Club" not in df.columns:
                    continue
                for _, row in df.iterrows():
                    club = str(row.get("Club", "")).strip()
                    if not club or club.lower() != club_key:
                        continue
                    rows.append(
                        {
                            "Race": race_num,
                            "Name": str(row.get("Name", "")).strip(),
                            "Gender": str(row.get("Gender", "")).strip(),
                            "Category": str(row.get("Category", "")).strip(),
                            "Time": str(row.get("Time", "")).strip(),
                            "Points": str(row.get("Points", "")).strip(),
                            "Team": str(row.get("Team", "")).strip(),
                        }
                    )
        except Exception as exc:
            self._show_message(f"Failed loading club results: {exc}")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        if not rows:
            self._show_message("Club not found in race sheets.")
            self._summary_var.set("No race rows for selected club.")
            self._refresh_resolve_controls(None)
            return

        df = pd.DataFrame(rows)
        self._latest_df = df.copy()
        runner_count = df["Name"].nunique() if "Name" in df.columns else 0
        self._summary_var.set(
            f"{selected}: {len(df)} race row(s), {runner_count} runner(s)."
        )
        self._populate_table(df)
        self._refresh_resolve_controls(None)

    def _load_runner_history(self):
        selected = self._runner_var.get().strip()
        if not selected:
            self._show_message("No runner selected.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        try:
            available = self._ensure_workbook_cache()
        except RuntimeError as exc:
            self._show_message(str(exc))
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        if not available:
            self._show_message("No results workbook found.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        runner_key = selected.lower()
        rows = []

        try:
            for sheet in self._wb_cache["race_sheets"]:
                race_num = extract_race_number(sheet) or ""
                df = self._wb_cache["sheets"][sheet].fillna("")
                if "Name" not in df.columns:
                    continue
                for _, row in df.iterrows():
                    name = str(row.get("Name", "")).strip()
                    if not name or name.lower() != runner_key:
                        continue
                    rows.append(
                        {
                            "Race": race_num,
                            "Club": str(row.get("Club", "")).strip(),
                            "Gender": str(row.get("Gender", "")).strip(),
                            "Category": str(row.get("Category", "")).strip(),
                            "Time": str(row.get("Time", "")).strip(),
                            "Points": str(row.get("Points", "")).strip(),
                            "Team": str(row.get("Team", "")).strip(),
                        }
                    )
        except Exception as exc:
            self._show_message(f"Failed loading runner history: {exc}")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        if not rows:
            self._show_message("Runner not found in race sheets.")
            self._summary_var.set("No race rows for selected runner.")
            self._refresh_resolve_controls(None)
            return

        df = pd.DataFrame(rows)
        self._latest_df = df.copy()
        try:
            points_total = int(pd.to_numeric(df["Points"], errors="coerce").fillna(0).sum())
        except Exception:
            points_total = 0

        self._summary_var.set(
            f"{selected}: {len(df)} race row(s), total points {points_total}."
        )
        self._populate_table(df)
        self._refresh_resolve_controls(df)

    def _on_runner_typed(self, _event=None):
        typed = self._runner_var.get().strip()
        if not self._all_runner_names:
            return

        if not typed:
            filtered = self._all_runner_names
        else:
            needle = typed.lower()
            starts = [n for n in self._all_runner_names if n.lower().startswith(needle)]
            starts_lower = {s.lower() for s in starts}
            contains = [n for n in self._all_runner_names if needle in n.lower() and n.lower() not in starts_lower]
            filtered = starts + contains

        self._runner_combo["values"] = filtered[:300]

    def _on_runner_enter(self, _event=None):
        typed = self._runner_var.get().strip()
        if not typed:
            return

        # Prefer exact match, then first prefix match.
        exact = next((n for n in self._all_runner_names if n.lower() == typed.lower()), None)
        if exact:
            self._runner_var.set(exact)
            self._load_runner_history()
            return

        prefix = next((n for n in self._all_runner_names if n.lower().startswith(typed.lower())), None)
        if prefix:
            self._runner_var.set(prefix)
            self._load_runner_history()

    def _on_club_typed(self, _event=None):
        typed = self._club_var.get().strip()
        if not self._all_clubs:
            return

        if not typed:
            filtered = self._all_clubs
        else:
            needle = typed.lower()
            starts = [c for c in self._all_clubs if c.lower().startswith(needle)]
            starts_lower = {s.lower() for s in starts}
            contains = [c for c in self._all_clubs if needle in c.lower() and c.lower() not in starts_lower]
            filtered = starts + contains

        self._club_combo["values"] = filtered[:300]

    def _on_club_enter(self, _event=None):
        typed = self._club_var.get().strip()
        if not typed:
            return

        exact = next((c for c in self._all_clubs if c.lower() == typed.lower()), None)
        if exact:
            self._club_var.set(exact)
            self._load_club_results()
            return

        prefix = next((c for c in self._all_clubs if c.lower().startswith(typed.lower())), None)
        if prefix:
            self._club_var.set(prefix)
            self._load_club_results()

    def _refresh_resolve_controls(self, df: pd.DataFrame | None) -> None:
        """Show/hide club/category resolve controls based on conflicting values."""
        self._resolve_frame.pack_forget()
        self._club_resolve_row.pack_forget()
        self._gender_resolve_row.pack_forget()
        self._category_resolve_row.pack_forget()
        self._time_resolve_row.pack_forget()

        if df is None or df.empty:
            return

        show_any = False

        def _norm(value) -> str:
            text = str(value).strip() if value is not None else ""
            return "" if text.lower() in {"nan", "none"} else text

        club_values_all = [_norm(v) for v in df.get("Club", pd.Series(dtype=str)).tolist()]
        club_values_unique = set(club_values_all)
        club_choices = sorted({v for v in club_values_all if v}, key=lambda s: s.lower())
        if len(club_values_unique) > 1 and club_choices:
            self._club_resolve_combo["values"] = club_choices
            if self._club_resolve_var.get() not in club_choices:
                self._club_resolve_var.set(club_choices[0])
            self._club_resolve_row.pack(fill="x", pady=(0, 6))
            show_any = True

        gender_values_all = [_norm(v).upper() for v in df.get("Gender", pd.Series(dtype=str)).tolist()]
        gender_nonblank = {v for v in gender_values_all if v in {"F", "M"}}
        has_blank_gender = any(not v for v in gender_values_all)
        if has_blank_gender or len(gender_nonblank) > 1:
            gender_choices = [value for value in ("F", "M") if value in gender_nonblank]
            for value in ("F", "M"):
                if value not in gender_choices:
                    gender_choices.append(value)
            self._gender_resolve_combo["values"] = gender_choices
            if self._gender_resolve_var.get() not in gender_choices:
                self._gender_resolve_var.set(next(iter(gender_nonblank), gender_choices[0]))
            self._gender_resolve_row.pack(fill="x", pady=(0, 6))
            show_any = True

        category_values_all = [_norm(v) for v in df.get("Category", pd.Series(dtype=str)).tolist()]
        category_values_unique = set(category_values_all)
        category_choices = sorted({v for v in category_values_all if v}, key=lambda s: s.lower())
        if len(category_values_unique) > 1 and category_choices:
            self._category_resolve_combo["values"] = category_choices
            if self._category_resolve_var.get() not in category_choices:
                self._category_resolve_var.set(category_choices[0])
            self._category_resolve_row.pack(fill="x")
            show_any = True

        time_values_all = [_norm(v) for v in df.get("Time", pd.Series(dtype=str)).tolist()]
        has_qry_time = any(v.upper() == "QRY" for v in time_values_all if v)
        if has_qry_time:
            self._time_resolve_row.pack(fill="x", pady=(6, 0))
            show_any = True

        if show_any:
            self._resolve_frame.pack(fill="x", padx=14, pady=(0, 6))

    def _resolve_club_across_inputs(self) -> None:
        self._resolve_field_across_inputs(field_type="club", target_value=self._club_resolve_var.get().strip())

    def _resolve_gender_across_inputs(self) -> None:
        self._resolve_field_across_inputs(field_type="gender", target_value=self._gender_resolve_var.get().strip().upper())

    def _resolve_category_across_inputs(self) -> None:
        self._resolve_field_across_inputs(field_type="category", target_value=self._category_resolve_var.get().strip())

    def _resolve_time_across_inputs(self) -> None:
        entered = self._time_resolve_var.get().strip()
        if not entered:
            messagebox.showwarning("No Time Entered", "Enter a replacement time value first.", parent=self)
            return

        seconds = parse_time_to_seconds(entered)
        if seconds is None or seconds <= 0:
            messagebox.showerror("Invalid Time", "Enter time as hh:mm:ss (for example 00:42:17).", parent=self)
            return

        total = int(round(seconds))
        hours = total // 3600
        minutes = (total % 3600) // 60
        secs = total % 60
        normalised = f"{hours:02d}:{minutes:02d}:{secs:02d}"
        self._resolve_field_across_inputs(field_type="time", target_value=normalised)

    def _copy_results(self) -> None:
        if self._latest_df is None or self._latest_df.empty:
            messagebox.showwarning("No Data", "No enquiry results to copy.", parent=self)
            return

        try:
            payload = self._latest_df.to_csv(sep="\t", index=False)
            self.clipboard_clear()
            self.clipboard_append(payload)
            self.update()
            messagebox.showinfo("Copied", f"Copied {len(self._latest_df)} row(s) to the clipboard.", parent=self)
        except Exception as exc:
            messagebox.showerror("Copy Failed", f"Could not copy results: {exc}", parent=self)

    def _export_results_csv(self) -> None:
        self._export_results_file("csv")

    def _export_results_excel(self) -> None:
        self._export_results_file("xlsx")

    def _export_results_file(self, file_type: str) -> None:
        if self._latest_df is None or self._latest_df.empty:
            messagebox.showwarning("No Data", "No enquiry results to export.", parent=self)
            return

        mode = self._lookup_mode
        selected = self._runner_var.get().strip() if mode == "runner" else self._club_var.get().strip()
        slug = "_".join(part for part in selected.replace("/", " ").split() if part) or mode
        default_name = f"{mode}_enquiry_{slug}.{file_type}"
        filetypes = [("CSV files", "*.csv")] if file_type == "csv" else [("Excel files", "*.xlsx")]

        filename = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=f".{file_type}",
            filetypes=filetypes + [("All files", "*.*")],
            initialfile=default_name,
        )
        if not filename:
            return

        try:
            safe_df = _sanitise_df_for_export(self._latest_df)
            if file_type == "csv":
                safe_df.to_csv(filename, index=False)
            else:
                safe_df.to_excel(filename, sheet_name="Enquiry", index=False)
            messagebox.showinfo("Exported", f"Saved {len(self._latest_df)} row(s) to {Path(filename).name}.", parent=self)
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Could not export results: {exc}", parent=self)

    def _resolve_field_across_inputs(self, field_type: str, target_value: str) -> None:
        selected_runner = self._runner_var.get().strip()
        if not selected_runner:
            messagebox.showwarning("No Runner Selected", "Select a runner first.", parent=self)
            return
        if not target_value:
            messagebox.showwarning("No Value Selected", f"Select a {field_type} value first.", parent=self)
            return

        raw_data_dir = session_config.raw_data_dir
        if not raw_data_dir or not raw_data_dir.exists():
            messagebox.showerror("Input Not Found", "Active input directory is not available.", parent=self)
            return

        race_files = discover_race_files(
            raw_data_dir,
            excluded_names=race_discovery_exclusions(),
        )
        if not race_files:
            messagebox.showinfo("No Race Files", "No race files found in the active input directory.", parent=self)
            return

        if not messagebox.askyesno(
            "Apply Across Inputs",
            f"Set {field_type} = '{target_value}' for runner '{selected_runner}' across all race input files?",
            parent=self,
        ):
            return

        updated_rows, touched_files, audit_changes, failed_files = resolve_runner_field_across_files(
            race_files,
            selected_runner=selected_runner,
            field_type=field_type,
            target_value=target_value,
        )

        log_error = log_manual_data_changes(
            audit_changes,
            source="Runner History",
            action=f"Resolve {field_type}",
        )
        if log_error:
            failed_files.append(f"Manual_Data_Audit: {log_error}")

        self._load_runner_history()

        if failed_files:
            messagebox.showwarning(
                "Applied With Issues",
                f"Updated {updated_rows} row(s) across {touched_files} file(s).\n\n"
                "Some files failed:\n" + "\n".join(failed_files),
                parent=self,
            )
        else:
            messagebox.showinfo(
                "Apply Complete",
                f"Updated {updated_rows} row(s) across {touched_files} file(s).",
                parent=self,
            )

    def _populate_table(self, df: pd.DataFrame):
        self._tree.delete(*self._tree.get_children())

        columns = list(df.columns)
        self._tree["columns"] = columns

        header_font = tkfont.Font(font=("Segoe UI", 10, "bold"))
        body_font = tkfont.Font(font=("Segoe UI", 10))

        for col in columns:
            anchor = "e" if col in {"Race", "Points"} else "w"
            self._tree.heading(col, text=col, anchor=anchor)
            width = self._measure_column_width(df, col, header_font, body_font)
            self._tree.column(col, width=width, minwidth=width, stretch=False, anchor=anchor)

        for index, (_, row) in enumerate(df.iterrows()):
            tag = "even" if index % 2 else "odd"
            self._tree.insert("", "end", values=list(row), tags=(tag,))

    def _show_message(self, msg: str):
        self._latest_df = None
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = ["Message"]
        self._tree.heading("Message", text="Message", anchor="w")
        self._tree.column("Message", width=800, minwidth=500, stretch=True, anchor="w")
        self._tree.insert("", "end", values=[msg], tags=("odd",))

    def _measure_column_width(self, df: pd.DataFrame, col: str, header_font, body_font) -> int:
        max_px = header_font.measure(str(col)) + 28
        for value in df[col].astype(str).tolist():
            max_px = max(max_px, body_font.measure(value) + 22)
        return min(max(max_px, 80), 420)
