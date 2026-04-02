import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
import pandas as pd
from ..session_config import config as session_config
from .results_workbook import find_latest_results_workbook


class ResultsViewerPanel(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#f5f5f5")
        self._style = ttk.Style(self)
        self._configure_styles()
        self._build_ui()

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

        # Dropdowns for race/overall/individual
        options = [
            ("Overall League Table", "overall"),
            ("Division 1 Teams", "div1"),
            ("Division 2 Teams", "div2"),
            ("Top 20 Male Individuals", "male"),
            ("Top 20 Female Individuals", "female"),
            ("Race Results", "race"),
        ]
        self._option_var = tk.StringVar(value="overall")
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
        if results_path is None:
            self._race_dropdown["values"] = []
            self._race_var.set("")
            return
        try:
            xl = pd.ExcelFile(results_path)
            races = [s for s in xl.sheet_names if s.startswith("Race ")]
            self._race_dropdown["values"] = races
            if races and self._race_var.get() not in races:
                self._race_var.set(races[0])
            elif not races:
                self._race_var.set("")
        except Exception:
            self._race_dropdown["values"] = []
            self._race_var.set("")

    def _load_results(self):
        out_dir = session_config.output_dir
        if not out_dir or not out_dir.exists():
            self._show_message("No output directory set.")
            return
        results_path = self._find_results_workbook()
        if results_path is None:
            self._show_message("No 'Race N -- Results.xlsx' file found in output directory.")
            return
        try:
            xl = pd.ExcelFile(results_path)
            opt = self._option_var.get()
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
            self._populate_table(df)
        except Exception as exc:
            self._show_message(f"Failed to load results: {exc}")

    def _find_results_workbook(self):
        return find_latest_results_workbook(session_config.output_dir)

    def _populate_table(self, df):
        self._tree.delete(*self._tree.get_children())
        display_df, numeric_columns = self._format_dataframe(df)
        self._tree["columns"] = list(display_df.columns)

        header_font = tkfont.Font(font=("Segoe UI", 10, "bold"))
        body_font = tkfont.Font(font=("Segoe UI", 10))

        for col in display_df.columns:
            anchor = "e" if col in numeric_columns else "w"
            width = self._measure_column_width(display_df, col, header_font, body_font)
            self._tree.heading(col, text=col, anchor=anchor)
            self._tree.column(col, width=width, minwidth=width, stretch=False, anchor=anchor)

        for index, (_, row) in enumerate(display_df.iterrows()):
            tag = "even" if index % 2 else "odd"
            self._tree.insert("", "end", values=list(row), tags=(tag,))

    def _show_message(self, msg):
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
        max_text_width = header_font.measure(str(column))
        for value in df[column].tolist():
            max_text_width = max(max_text_width, body_font.measure(str(value)))

        if column.lower() in {"name", "club", "item", "notes"}:
            return min(max(max_text_width + 28, 140), 260)
        if "race" in column.lower() or "points" in column.lower() or "score" in column.lower() or column.lower() == "position":
            return min(max(max_text_width + 24, 72), 110)
        return min(max(max_text_width + 24, 90), 180)
