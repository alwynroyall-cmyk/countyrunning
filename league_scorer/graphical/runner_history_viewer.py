import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk

import openpyxl
import pandas as pd

from ..manual_data_audit import log_manual_data_changes
from ..race_processor import extract_race_number
from ..session_config import config as session_config
from ..source_loader import discover_race_files
from .dashboard import WRRL_GREEN, WRRL_LIGHT, WRRL_NAVY, WRRL_WHITE


class RunnerHistoryPanel(tk.Frame):
    """View one runner's results across all race sheets in the latest results workbook."""

    def __init__(self, parent, back_callback=None):
        super().__init__(parent, bg=WRRL_LIGHT)
        self._back_callback = back_callback
        self._style = ttk.Style(self)
        self._runner_name_map = {}
        self._all_runner_names: list[str] = []
        self._latest_df: pd.DataFrame | None = None
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
            text="Runner:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))

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

    def _find_results_workbook(self):
        out_dir = session_config.output_dir
        if not out_dir or not out_dir.exists():
            return None

        candidates = []
        for path in out_dir.glob("*.xlsx"):
            if not path.name.lower().endswith("-- results.xlsx"):
                continue
            race_number = extract_race_number(path.stem)
            if race_number is None:
                continue
            candidates.append((race_number, path))

        if not candidates:
            return None

        return max(candidates, key=lambda item: item[0])[1]

    def _load_runner_options(self):
        workbook = self._find_results_workbook()
        if workbook is None:
            self._runner_combo["values"] = []
            self._runner_var.set("")
            self._show_message("No results workbook found in the active output folder.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        try:
            xl = pd.ExcelFile(workbook)
            race_sheets = sorted(
                [name for name in xl.sheet_names if name.startswith("Race ")],
                key=lambda s: extract_race_number(s) or 0,
            )
            names = {}
            for sheet in race_sheets:
                df = xl.parse(sheet)
                if "Name" not in df.columns:
                    continue
                for value in df["Name"].dropna().tolist():
                    name = str(value).strip()
                    if name:
                        key = name.lower()
                        if key not in names:
                            names[key] = name
        except Exception as exc:
            self._runner_combo["values"] = []
            self._runner_var.set("")
            self._show_message(f"Failed reading results workbook: {exc}")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        self._runner_name_map = names
        sorted_names = sorted(names.values(), key=lambda n: n.lower())
        self._all_runner_names = sorted_names
        self._runner_combo["values"] = sorted_names

        if sorted_names:
            current = self._runner_var.get()
            if current not in sorted_names:
                self._runner_var.set(sorted_names[0])
            self._load_runner_history()
        else:
            self._runner_var.set("")
            self._show_message("No runner names found in race sheets.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)

    def _load_runner_history(self):
        selected = self._runner_var.get().strip()
        if not selected:
            self._show_message("No runner selected.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        workbook = self._find_results_workbook()
        if workbook is None:
            self._show_message("No results workbook found.")
            self._summary_var.set("")
            self._refresh_resolve_controls(None)
            return

        runner_key = selected.lower()
        rows = []

        try:
            xl = pd.ExcelFile(workbook)
            race_sheets = sorted(
                [name for name in xl.sheet_names if name.startswith("Race ")],
                key=lambda s: extract_race_number(s) or 0,
            )
            for sheet in race_sheets:
                race_num = extract_race_number(sheet) or ""
                df = xl.parse(sheet)
                if "Name" not in df.columns:
                    continue
                df = df.fillna("")
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
            contains = [n for n in self._all_runner_names if needle in n.lower() and n.lower() not in {s.lower() for s in starts}]
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

    def _refresh_resolve_controls(self, df: pd.DataFrame | None) -> None:
        """Show/hide club/category resolve controls based on conflicting values."""
        self._resolve_frame.pack_forget()
        self._club_resolve_row.pack_forget()
        self._category_resolve_row.pack_forget()

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

        category_values_all = [_norm(v) for v in df.get("Category", pd.Series(dtype=str)).tolist()]
        category_values_unique = set(category_values_all)
        category_choices = sorted({v for v in category_values_all if v}, key=lambda s: s.lower())
        if len(category_values_unique) > 1 and category_choices:
            self._category_resolve_combo["values"] = category_choices
            if self._category_resolve_var.get() not in category_choices:
                self._category_resolve_var.set(category_choices[0])
            self._category_resolve_row.pack(fill="x")
            show_any = True

        if show_any:
            self._resolve_frame.pack(fill="x", padx=14, pady=(0, 6))

    def _resolve_club_across_inputs(self) -> None:
        self._resolve_field_across_inputs(field_type="club", target_value=self._club_resolve_var.get().strip())

    def _resolve_category_across_inputs(self) -> None:
        self._resolve_field_across_inputs(field_type="category", target_value=self._category_resolve_var.get().strip())

    def _resolve_field_across_inputs(self, field_type: str, target_value: str) -> None:
        selected_runner = self._runner_var.get().strip()
        if not selected_runner:
            messagebox.showwarning("No Runner Selected", "Select a runner first.", parent=self)
            return
        if not target_value:
            messagebox.showwarning("No Value Selected", f"Select a {field_type} value first.", parent=self)
            return

        input_dir = session_config.input_dir
        if not input_dir or not input_dir.exists():
            messagebox.showerror("Input Not Found", "Active input directory is not available.", parent=self)
            return

        race_files = discover_race_files(
            input_dir,
            excluded_names=("clubs.xlsx", "name_corrections.xlsx", "wrrl_events.xlsx"),
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

        runner_key = selected_runner.lower()
        updated_rows = 0
        touched_files = 0
        audit_changes: list[dict] = []
        failed_files: list[str] = []

        for _, path in race_files.items():
            try:
                wb = openpyxl.load_workbook(path)
            except Exception as exc:
                failed_files.append(f"{path.name}: {exc}")
                continue

            try:
                ws = wb.active
                name_col, field_col = self._find_columns(ws, field_type)
                if name_col is None or field_col is None:
                    wb.close()
                    continue

                file_changed = False
                for row_idx in range(2, ws.max_row + 1):
                    row_name = self._row_name_value(ws, row_idx, name_col)
                    if row_name.lower() != runner_key:
                        continue

                    current = ws.cell(row=row_idx, column=field_col).value
                    current_text = "" if current is None else str(current).strip()
                    if current_text == target_value:
                        continue

                    ws.cell(row=row_idx, column=field_col).value = target_value
                    updated_rows += 1
                    file_changed = True
                    audit_changes.append(
                        {
                            "runner": selected_runner,
                            "field": field_type,
                            "old_value": current_text,
                            "new_value": target_value,
                            "file_path": path,
                            "row_idx": row_idx,
                        }
                    )

                if file_changed:
                    wb.save(path)
                    touched_files += 1
            except Exception as exc:
                failed_files.append(f"{path.name}: {exc}")
            finally:
                wb.close()

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

    def _find_columns(self, ws, field_type: str):
        headers = [
            str(c.value).strip().lower() if c.value is not None else ""
            for c in next(ws.iter_rows(min_row=1, max_row=1))
        ]

        name_col = next(
            (i + 1 for i, h in enumerate(headers) if "name" in h and "first" not in h and "last" not in h),
            None,
        )
        if name_col is None:
            first_col = next((i + 1 for i, h in enumerate(headers) if "first" in h), None)
            last_col = next((i + 1 for i, h in enumerate(headers) if "last" in h), None)
            if first_col is not None and last_col is not None:
                name_col = (first_col, last_col)

        if field_type == "club":
            field_col = next((i + 1 for i, h in enumerate(headers) if "club" in h), None)
        else:
            field_col = next((i + 1 for i, h in enumerate(headers) if "category" in h or h == "cat"), None)

        return name_col, field_col

    def _row_name_value(self, ws, row_idx: int, name_col) -> str:
        if isinstance(name_col, tuple):
            first_val = ws.cell(row=row_idx, column=name_col[0]).value
            last_val = ws.cell(row=row_idx, column=name_col[1]).value
            first = "" if first_val is None else str(first_val).strip()
            last = "" if last_val is None else str(last_val).strip()
            return f"{first} {last}".strip()
        val = ws.cell(row=row_idx, column=name_col).value
        return "" if val is None else str(val).strip()

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
