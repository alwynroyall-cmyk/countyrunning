import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

import pandas as pd

from ..race_processor import extract_race_number
from ..session_config import config as session_config
from .dashboard import WRRL_GREEN, WRRL_LIGHT, WRRL_NAVY, WRRL_WHITE
from .results_workbook import find_latest_results_workbook, sorted_race_sheet_names


class ClubHistoryPanel(tk.Frame):
    """View one club's results across all race sheets in the latest results workbook."""

    def __init__(self, parent, back_callback=None):
        super().__init__(parent, bg=WRRL_LIGHT)
        self._back_callback = back_callback
        self._style = ttk.Style(self)
        self._configure_styles()
        self._build_ui()
        self._load_club_options()

    def _configure_styles(self):
        self._style.configure(
            "ClubHistory.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#22313f",
            rowheight=27,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self._style.configure(
            "ClubHistory.Treeview.Heading",
            background="#dbe5ee",
            foreground="#22313f",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )

    def _build_ui(self):
        header = tk.Frame(self, bg=WRRL_NAVY, padx=14, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Club Results Across Races",
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

        controls = tk.Frame(self, bg=WRRL_LIGHT, padx=14, pady=10)
        controls.pack(fill="x")

        tk.Label(
            controls,
            text="Club:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))

        self._club_var = tk.StringVar()
        self._club_combo = ttk.Combobox(
            controls,
            textvariable=self._club_var,
            state="readonly",
            width=42,
        )
        self._club_combo.pack(side="left")
        self._club_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_club_history())

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
            command=self._load_club_options,
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

        body = tk.Frame(self, bg=WRRL_LIGHT)
        body.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        body.rowconfigure(1, weight=1)
        body.columnconfigure(0, weight=1)

        summary_box = tk.LabelFrame(
            body,
            text="Race Summary",
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
            font=("Segoe UI", 10, "bold"),
            padx=8,
            pady=8,
        )
        summary_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        summary_box.columnconfigure(0, weight=1)

        self._summary_tree = ttk.Treeview(summary_box, show="headings", height=6, style="ClubHistory.Treeview")
        self._summary_tree.grid(row=0, column=0, sticky="nsew")
        sum_scroll = ttk.Scrollbar(summary_box, orient="vertical", command=self._summary_tree.yview)
        sum_scroll.grid(row=0, column=1, sticky="ns")
        self._summary_tree.configure(yscrollcommand=sum_scroll.set)

        details_box = tk.LabelFrame(
            body,
            text="Runner Rows",
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
            font=("Segoe UI", 10, "bold"),
            padx=8,
            pady=8,
        )
        details_box.grid(row=1, column=0, sticky="nsew")
        details_box.rowconfigure(0, weight=1)
        details_box.columnconfigure(0, weight=1)

        self._details_tree = ttk.Treeview(details_box, show="headings", style="ClubHistory.Treeview")
        self._details_tree.grid(row=0, column=0, sticky="nsew")

        detail_y = ttk.Scrollbar(details_box, orient="vertical", command=self._details_tree.yview)
        detail_y.grid(row=0, column=1, sticky="ns")
        detail_x = ttk.Scrollbar(details_box, orient="horizontal", command=self._details_tree.xview)
        detail_x.grid(row=1, column=0, sticky="ew")
        self._details_tree.configure(yscrollcommand=detail_y.set, xscrollcommand=detail_x.set)

    def _find_results_workbook(self):
        return find_latest_results_workbook(session_config.output_dir)

    def _load_club_options(self):
        workbook = self._find_results_workbook()
        if workbook is None:
            self._club_combo["values"] = []
            self._club_var.set("")
            self._show_summary_message("No results workbook found in the active output folder.")
            self._show_detail_message("No results workbook found in the active output folder.")
            self._summary_var.set("")
            return

        clubs = {}
        try:
            xl = pd.ExcelFile(workbook)
            race_sheets = sorted_race_sheet_names(xl)
            for sheet in race_sheets:
                df = xl.parse(sheet)
                if "Club" not in df.columns:
                    continue
                for value in df["Club"].dropna().tolist():
                    club = str(value).strip()
                    if club:
                        key = club.lower()
                        if key not in clubs:
                            clubs[key] = club
        except Exception as exc:
            self._club_combo["values"] = []
            self._club_var.set("")
            self._show_summary_message(f"Failed reading results workbook: {exc}")
            self._show_detail_message(f"Failed reading results workbook: {exc}")
            self._summary_var.set("")
            return

        sorted_clubs = sorted(clubs.values(), key=lambda c: c.lower())
        self._club_combo["values"] = sorted_clubs

        if sorted_clubs:
            current = self._club_var.get()
            if current not in sorted_clubs:
                self._club_var.set(sorted_clubs[0])
            self._load_club_history()
        else:
            self._club_var.set("")
            self._show_summary_message("No clubs found in race sheets.")
            self._show_detail_message("No clubs found in race sheets.")
            self._summary_var.set("")

    def _load_club_history(self):
        selected = self._club_var.get().strip()
        if not selected:
            self._show_summary_message("No club selected.")
            self._show_detail_message("No club selected.")
            self._summary_var.set("")
            return

        workbook = self._find_results_workbook()
        if workbook is None:
            self._show_summary_message("No results workbook found.")
            self._show_detail_message("No results workbook found.")
            self._summary_var.set("")
            return

        club_key = selected.lower()
        detail_rows = []

        try:
            xl = pd.ExcelFile(workbook)
            race_sheets = sorted_race_sheet_names(xl)
            for sheet in race_sheets:
                race_num = extract_race_number(sheet) or ""
                df = xl.parse(sheet).fillna("")
                if "Club" not in df.columns:
                    continue
                for _, row in df.iterrows():
                    club = str(row.get("Club", "")).strip()
                    if not club or club.lower() != club_key:
                        continue
                    detail_rows.append(
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
            self._show_summary_message(f"Failed loading club history: {exc}")
            self._show_detail_message(f"Failed loading club history: {exc}")
            self._summary_var.set("")
            return

        if not detail_rows:
            self._show_summary_message("Club not found in race sheets.")
            self._show_detail_message("Club not found in race sheets.")
            self._summary_var.set("No race rows for selected club.")
            return

        detail_df = pd.DataFrame(detail_rows)
        points_numeric = pd.to_numeric(detail_df["Points"], errors="coerce").fillna(0)
        race_summary = (
            detail_df.assign(_points=points_numeric)
            .groupby("Race", as_index=False)
            .agg(
                Runners=("Name", "count"),
                Total_Points=("_points", "sum"),
                Top_Scorer=("Name", "first"),
            )
            .sort_values("Race")
        )
        race_summary["Total_Points"] = race_summary["Total_Points"].astype(int)
        race_summary.rename(
            columns={
                "Total_Points": "Total Points",
                "Top_Scorer": "Top Scorer",
            },
            inplace=True,
        )

        total_points = int(points_numeric.sum())
        self._summary_var.set(
            f"{selected}: {len(detail_df)} runner row(s) across {len(race_summary)} race(s), total points {total_points}."
        )

        self._populate_summary_table(race_summary)
        self._populate_detail_table(detail_df.sort_values(["Race", "Name"]))

    def _populate_summary_table(self, df: pd.DataFrame):
        self._summary_tree.delete(*self._summary_tree.get_children())
        cols = list(df.columns)
        self._summary_tree["columns"] = cols

        header_font = tkfont.Font(font=("Segoe UI", 10, "bold"))
        body_font = tkfont.Font(font=("Segoe UI", 10))

        for col in cols:
            anchor = "e" if col in {"Race", "Runners", "Total Points"} else "w"
            self._summary_tree.heading(col, text=col, anchor=anchor)
            width = self._measure_column_width(df, col, header_font, body_font)
            self._summary_tree.column(col, width=width, minwidth=width, stretch=False, anchor=anchor)

        for _, row in df.iterrows():
            self._summary_tree.insert("", "end", values=list(row))

    def _populate_detail_table(self, df: pd.DataFrame):
        self._details_tree.delete(*self._details_tree.get_children())
        cols = list(df.columns)
        self._details_tree["columns"] = cols

        header_font = tkfont.Font(font=("Segoe UI", 10, "bold"))
        body_font = tkfont.Font(font=("Segoe UI", 10))

        for col in cols:
            anchor = "e" if col in {"Race", "Points"} else "w"
            self._details_tree.heading(col, text=col, anchor=anchor)
            width = self._measure_column_width(df, col, header_font, body_font)
            self._details_tree.column(col, width=width, minwidth=width, stretch=False, anchor=anchor)

        for _, row in df.iterrows():
            self._details_tree.insert("", "end", values=list(row))

    def _show_summary_message(self, msg: str):
        self._summary_tree.delete(*self._summary_tree.get_children())
        self._summary_tree["columns"] = ["Message"]
        self._summary_tree.heading("Message", text="Message", anchor="w")
        self._summary_tree.column("Message", width=900, minwidth=400, stretch=True, anchor="w")
        self._summary_tree.insert("", "end", values=[msg])

    def _show_detail_message(self, msg: str):
        self._details_tree.delete(*self._details_tree.get_children())
        self._details_tree["columns"] = ["Message"]
        self._details_tree.heading("Message", text="Message", anchor="w")
        self._details_tree.column("Message", width=900, minwidth=400, stretch=True, anchor="w")
        self._details_tree.insert("", "end", values=[msg])

    def _measure_column_width(self, df: pd.DataFrame, col: str, header_font, body_font) -> int:
        max_px = header_font.measure(str(col)) + 28
        for value in df[col].astype(str).tolist():
            max_px = max(max_px, body_font.measure(value) + 22)
        return min(max(max_px, 80), 420)
