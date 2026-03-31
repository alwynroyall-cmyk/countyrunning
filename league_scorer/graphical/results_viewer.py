import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import pandas as pd
from ..session_config import config as session_config

class ResultsViewerPanel(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#f5f5f5")
        self._build_ui()

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
        ]
        self._option_var = tk.StringVar(value="overall")
        opt_frame = tk.Frame(self, bg="#f5f5f5")
        opt_frame.pack(pady=(0, 10))
        for text, val in options:
            ttk.Radiobutton(opt_frame, text=text, variable=self._option_var, value=val, command=self._load_results, style="TRadiobutton").pack(side="left", padx=8)

        # Race selector
        self._race_var = tk.StringVar()
        self._race_dropdown = ttk.Combobox(self, textvariable=self._race_var, state="readonly")
        self._race_dropdown.pack(pady=(0, 10))
        self._race_dropdown.bind("<<ComboboxSelected>>", lambda e: self._load_results())

        # Table
        self._tree = ttk.Treeview(self, show="headings")
        self._tree.pack(fill="both", expand=True, padx=10, pady=10)
        self._tree_scroll = ttk.Scrollbar(self, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=self._tree_scroll.set)
        self._tree_scroll.pack(side="right", fill="y")

        self._load_race_options()
        self._load_results()

    def _load_race_options(self):
        # Try to find Results.xlsx in output dir
        out_dir = session_config.output_dir
        if not out_dir or not out_dir.exists():
            self._race_dropdown["values"] = []
            self._race_var.set("")
            return
        results_path = out_dir / "Results.xlsx"
        if not results_path.exists():
            self._race_dropdown["values"] = []
            self._race_var.set("")
            return
        try:
            xl = pd.ExcelFile(results_path)
            races = [s for s in xl.sheet_names if s.startswith("Race ")]
            self._race_dropdown["values"] = races
            if races:
                self._race_var.set(races[0])
            else:
                self._race_var.set("")
        except Exception:
            self._race_dropdown["values"] = []
            self._race_var.set("")

    def _load_results(self):
        out_dir = session_config.output_dir
        if not out_dir or not out_dir.exists():
            self._show_message("No output directory set.")
            return
        results_path = out_dir / "Results.xlsx"
        if not results_path.exists():
            self._show_message("No Results.xlsx found in output directory.")
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
            else:
                race = self._race_var.get()
                if race:
                    df = xl.parse(race)
                else:
                    self._show_message("No race selected.")
                    return
            self._populate_table(df)
        except Exception as exc:
            self._show_message(f"Failed to load results: {exc}")

    def _populate_table(self, df):
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = list(df.columns)
        for col in df.columns:
            self._tree.heading(col, text=col)
            self._tree.column(col, width=120, anchor="center")
        for _, row in df.iterrows():
            self._tree.insert("", "end", values=list(row))

    def _show_message(self, msg):
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = ["Message"]
        self._tree.heading("Message", text="Message")
        self._tree.column("Message", width=400, anchor="center")
        self._tree.insert("", "end", values=[msg])
