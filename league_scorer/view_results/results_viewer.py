import re
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from pathlib import Path
from ..session_config import config as session_config
from ..graphical.results_workbook import find_latest_results_workbook


class ResultsViewerPanel(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#f5f5f5")
        self._style = ttk.Style(self)
        self._current_df = None
        self._current_option = "overall"
        self._current_results_path = None
        self._status_var = tk.StringVar(value="")
        self._xl_cache: tuple[pd.ExcelFile | None, Path | None, float] = (None, None, 0.0)
        self._configure_styles()
        self._build_ui()

    def _get_xl(self, results_path: Path) -> pd.ExcelFile | None:
        """Return a cached pd.ExcelFile for *results_path*, re-opening only when the file changes."""
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

    def _configure_styles(self):
        self._style.configure(
            "Results.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#22313f",
            rowheight=28,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self._style.configure(
            "Results.Treeview.Heading",
            background="#dbe5ee",
            foreground="#22313f",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        self._style.map(
            "Results.Treeview",
            background=[("selected", "#8fb3d1")],
            foreground=[("selected", "#102030")],
        )

    def _build_ui(self):
        title = tk.Label(self, text="View League Results", font=("Segoe UI", 16, "bold"), bg="#f5f5f5", fg="#3a4658")
        title.pack(pady=(10, 8))

        utility_frame = tk.Frame(self, bg="#f5f5f5")
        utility_frame.pack(fill="x", padx=10, pady=(0, 6))

        tk.Button(
            utility_frame,
            text="Refresh Workbook",
            command=self._refresh_results,
            font=("Segoe UI", 9),
            bg="#dbe1e8",
            fg="#22313f",
            relief="flat",
            padx=10,
            pady=4,
        ).pack(side="left")

        tk.Button(
            utility_frame,
            text="Open Workbook",
            command=self._open_results_workbook,
            font=("Segoe UI", 9),
            bg="#dbe1e8",
            fg="#22313f",
            relief="flat",
            padx=10,
            pady=4,
        ).pack(side="left", padx=(8, 0))

        tk.Label(
            utility_frame,
            textvariable=self._status_var,
            font=("Segoe UI", 9),
            bg="#f5f5f5",
            fg="#66707c",
            anchor="e",
        ).pack(side="right")

        # Export buttons
        export_frame = tk.Frame(self, bg="#f5f5f5")
        export_frame.pack(pady=(0, 10))
        
        tk.Button(
            export_frame,
            text="📋 Copy to Clipboard",
            command=self._export_to_clipboard,
            font=("Segoe UI", 9),
            bg="#e8f0f7",
            fg="#22313f",
            relief="flat",
            padx=10,
            pady=4,
        ).pack(side="left", padx=4)
        
    
        tk.Button(
            export_frame,
            text="📊 Export Excel",
            command=self._export_to_excel,
            font=("Segoe UI", 9),
            bg="#e8f0f7",
            fg="#22313f",
            relief="flat",
            padx=10,
            pady=4,
        ).pack(side="left", padx=4)

        # Dropdowns for race/overall/individual
        options = [
            ("Division 1 Teams", "div1"),
            ("Division 2 Teams", "div2"),
            ("Top 20 Male Individuals", "male"),
            ("Top 20 Female Individuals", "female"),
            ("Race Results", "race"),
        ]
        self._option_var = tk.StringVar(value="div1")
        opt_frame = tk.Frame(self, bg="#f5f5f5")
        opt_frame.pack(pady=(0, 10))
        for text, val in options:
            ttk.Radiobutton(opt_frame, text=text, variable=self._option_var, value=val, command=self._on_option_change, style="TRadiobutton").pack(side="left", padx=8)

        # Race selector
        self._race_var = tk.StringVar()
        self._race_dropdown = ttk.Combobox(self, textvariable=self._race_var, state="readonly")
        self._race_dropdown.pack(pady=(0, 10))
        self._race_dropdown.bind("<<ComboboxSelected>>", lambda e: self._load_results())

        # Table
        table_frame = tk.Frame(self, bg="#f5f5f5")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self._tree = ttk.Treeview(
            table_frame,
            show="headings",
            style="Results.Treeview",
        )
        self._tree.grid(row=0, column=0, sticky="nsew")

        self._tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
        self._tree_scroll.grid(row=0, column=1, sticky="ns")

        self._tree_x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self._tree.xview)
        self._tree_x_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self._tree.configure(yscrollcommand=self._tree_scroll.set, xscrollcommand=self._tree_x_scroll.set)
        self._tree.tag_configure("odd", background="#ffffff")
        self._tree.tag_configure("even", background="#eef3f8")

        self._load_race_options()
        self._update_race_selector_state()
        self._load_results()

    def _on_option_change(self):
        self._update_race_selector_state()
        self._load_results()

    def _update_race_selector_state(self):
        state = "readonly" if self._option_var.get() == "race" else "disabled"
        self._race_dropdown.configure(state=state)

    def _load_race_options(self):
        results_path = self._find_results_workbook()
        self._current_results_path = results_path
        if results_path is None:
            self._status_var.set("No standings workbook found")
            self._race_dropdown["values"] = []
            self._race_var.set("")
            return
        xl = self._get_xl(results_path)
        if xl is None:
            self._status_var.set(f"Failed to read workbook: {results_path.name}")
            self._race_dropdown["values"] = []
            self._race_var.set("")
            return
        self._status_var.set(f"Workbook: {results_path.name}")
        races = [s for s in xl.sheet_names if s.startswith("Race ")]
        self._race_dropdown["values"] = races
        if races and self._race_var.get() not in races:
            self._race_var.set(races[0])
        elif not races:
            self._race_var.set("")

    def _load_results(self):
        out_dir = session_config.output_dir
        if not out_dir or not out_dir.exists():
            self._show_message("No output directory set.")
            return
        results_path = self._find_results_workbook()
        self._current_results_path = results_path
        if results_path is None:
            self._show_message("No standings workbook found in outputs/publish/xlsx/standings.")
            return
        xl = self._get_xl(results_path)
        if xl is None:
            self._show_message(f"Failed to open workbook: {results_path.name}")
            return
        try:
            self._status_var.set(f"Workbook: {results_path.name}")
            opt = self._option_var.get()
            self._current_option = opt
            if opt == "overall":
                df = xl.parse("Summary")
            elif opt == "div1":
                df = xl.parse("Div 1")
            elif opt == "div2":
                df = xl.parse("Div 2")
            elif opt == "male":
                df = xl.parse("Male").head(20)
            elif opt == "female":
                df = xl.parse("Female").head(20)
            elif opt == "race":
                race = self._race_var.get()
                if race:
                    df = xl.parse(race)
                else:
                    self._show_message("No race selected.")
                    return
            else:
                self._show_message("Unknown results view selected.")
                return
            self._current_df = df
            self._populate_table(df)
        except Exception as exc:
            self._show_message(f"Failed to load results: {exc}")

    def _refresh_results(self):
        self._load_race_options()
        self._update_race_selector_state()
        self._load_results()

    def _open_results_workbook(self):
        if self._current_results_path is None or not self._current_results_path.exists():
            messagebox.showwarning("Workbook Missing", "No standings workbook is available to open.", parent=self)
            return
        try:
            import os
            if sys.platform == "win32":
                os.startfile(str(self._current_results_path))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(self._current_results_path)], check=False)
            else:
                subprocess.run(["xdg-open", str(self._current_results_path)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", f"Could not open workbook: {exc}", parent=self)

    def _find_results_workbook(self):
        return find_latest_results_workbook(session_config.output_dir)

    def _populate_table(self, df):
        self._tree.delete(*self._tree.get_children())
        display_df, numeric_columns = self._format_dataframe(df)

        # Rename race-related columns for Div 1 / Div 2 views
        if self._current_option in ("div1", "div2"):
            col_rename = {}
            for col in list(display_df.columns):
                new = col
                # Race N Men/Female Score -> R{N} Men / R{N} Women
                m = re.match(r"Race\s*(\d+)\s*(Men|Female|Women)(?:\s*Score)?$", col, re.IGNORECASE)
                if m:
                    n = m.group(1)
                    gender = m.group(2)
                    gender_norm = "M" if gender.lower().startswith("men") else "F"
                    new = f"R{n} {gender_norm}"
                else:
                    # Race N aggregate -> R{N} Score
                    m2 = re.match(r"Race\s*(\d+)\s*aggregate$", col, re.IGNORECASE)
                    if m2:
                        n = m2.group(1)
                        new = f"R{n} Sc"
                    else:
                        # Race N Team Points -> R{N} Pts
                        m3 = re.match(r"Race\s*(\d+)\s*Team\s*Points$", col, re.IGNORECASE)
                        if m3:
                            n = m3.group(1)
                            new = f"R{n} Pts"
                if new != col:
                    col_rename[col] = new
            if col_rename:
                display_df = display_df.rename(columns=col_rename)
                # update numeric columns names to the renamed variants
                numeric_columns = {col_rename.get(c, c) for c in numeric_columns}
        self._tree["columns"] = list(display_df.columns)

        header_font = tkfont.Font(font=("Segoe UI", 10, "bold"))
        body_font = tkfont.Font(font=("Segoe UI", 10))

        for col in display_df.columns:
            # centre the position column in all views
            if col.lower() == "position":
                anchor = "center"
            else:
                anchor = "e" if col in numeric_columns else "w"
            width = self._measure_column_width(display_df, col, header_font, body_font)
            self._tree.heading(col, text=col, anchor=anchor)
            self._tree.column(col, width=width, minwidth=width, stretch=False, anchor=anchor)

        for index, (_, row) in enumerate(display_df.iterrows()):
            tag = "even" if index % 2 else "odd"
            self._tree.insert("", "end", values=list(row), tags=(tag,))

    def _show_message(self, msg):
        self._current_df = None
        if not self._status_var.get():
            self._status_var.set("No workbook loaded")
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = ["Message"]
        self._tree.heading("Message", text="Message", anchor="w")
        self._tree.column("Message", width=480, minwidth=480, stretch=False, anchor="w")
        self._tree.insert("", "end", values=[msg], tags=("odd",))

    def _format_dataframe(self, df):
        display_df = df.copy()
        numeric_columns = set()

        for col in display_df.columns:
            series = display_df[col]
            numeric_series = pd.to_numeric(series, errors="coerce")
            non_null = series.dropna()
            numeric_non_null = numeric_series[series.notna()]
            is_numeric = not non_null.empty and numeric_non_null.notna().all()

            if is_numeric:
                numeric_columns.add(col)

            formatted_values = []
            for original, numeric_value in zip(series.tolist(), numeric_series.tolist()):
                if pd.isna(original):
                    formatted_values.append("")
                elif is_numeric and pd.notna(numeric_value):
                    formatted_values.append(str(int(round(float(numeric_value)))))
                else:
                    formatted_values.append(str(original))

            display_df[col] = formatted_values

        return display_df, numeric_columns

    def _measure_column_width(self, df, column, header_font, body_font):
        # Determine a uniform target width derived from any 'R# M' / 'R# F' column
        mf_col = None
        for colname in df.columns:
            if re.match(r"^R\d+\s+[MF]$", str(colname)):
                mf_col = colname
                break

        if mf_col is not None:
            header_width = header_font.measure(str(mf_col))
            max_text_width = header_width
            for v in df[mf_col].tolist():
                max_text_width = max(max_text_width, body_font.measure(str(v)))
            base = min(max(max_text_width + 24, 72), 110)
            # make uniform width slightly smaller (10% reduction) but never smaller
            # than the header width plus padding
            raw = int(base * 1.2 * 0.9)
            target = max(raw, header_width + 12, 36)
            return target

        # Fallback: compute width for this column if no M/F columns were found
        max_text_width = header_font.measure(str(column))
        for value in df[column].tolist():
            max_text_width = max(max_text_width, body_font.measure(str(value)))

        if column.lower() in {"name", "club", "item", "notes"}:
            return min(max(max_text_width + 28, 140), 260)

        if "race" in column.lower() or "points" in column.lower() or "score" in column.lower() or column.lower() == "position":
            fallback = min(max_text_width + 24, 110)
            return max(int(fallback * 0.9), header_font.measure(str(column)) + 12)
        fallback = min(max_text_width + 24, 180)
        return max(int(fallback * 0.9), header_font.measure(str(column)) + 12)

    def _export_to_clipboard(self):
        """Copy the current table to clipboard as tab-separated values."""
        if self._current_df is None or self._current_df.empty:
            messagebox.showwarning("No Data", "No results to copy.", parent=self)
            return
        
        try:
            # Convert to tab-separated values
            tsv_data = self._current_df.to_csv(sep='\t', index=False)
            # Copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(tsv_data)
            self.update()
            messagebox.showinfo("Success", f"Copied {len(self._current_df)} rows to clipboard.", parent=self)
        except Exception as exc:
            messagebox.showerror("Export Error", f"Failed to copy to clipboard: {exc}", parent=self)


    def _export_to_excel(self):
        """Export the current table to Excel file."""
        if self._current_df is None or self._current_df.empty:
            messagebox.showwarning("No Data", "No results to export.", parent=self)
            return
        
        filename = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"results_{self._current_option}.xlsx"
        )
        
        if not filename:
            return
        
        try:
            self._current_df.to_excel(filename, sheet_name="Results", index=False)
            messagebox.showinfo("Success", f"Exported {len(self._current_df)} rows to {Path(filename).name}", parent=self)
        except Exception as exc:
            messagebox.showerror("Export Error", f"Failed to export Excel: {exc}", parent=self)
