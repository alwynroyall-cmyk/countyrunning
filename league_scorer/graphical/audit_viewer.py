import tkinter as tk
from pathlib import Path
from tkinter import ttk

import pandas as pd

from ..audit_data_service import list_audit_workbooks


class AuditViewerPanel(tk.Frame):
    def __init__(self, parent, preferred_workbook: Path | None = None):
        super().__init__(parent, bg="#f5f5f5")
        self._style = ttk.Style(self)
        self._workbooks: dict[str, Path] = {}
        self._preferred_workbook = preferred_workbook
        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        self._style.configure(
            "Audit.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#22313f",
            rowheight=28,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self._style.configure(
            "Audit.Treeview.Heading",
            background="#dbe5ee",
            foreground="#22313f",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )

    def _build_ui(self):
        title = tk.Label(self, text="View Audit Output", font=("Segoe UI", 16, "bold"), bg="#f5f5f5", fg="#3a4658")
        title.pack(pady=(10, 8))

        selector = tk.Frame(self, bg="#f5f5f5")
        selector.pack(pady=(0, 10), fill="x", padx=10)

        tk.Label(selector, text="Workbook", font=("Segoe UI", 10, "bold"), bg="#f5f5f5").pack(side="left", padx=(0, 8))
        self._file_var = tk.StringVar()
        self._file_dropdown = ttk.Combobox(selector, textvariable=self._file_var, state="readonly", width=36)
        self._file_dropdown.pack(side="left", padx=(0, 16))
        self._file_dropdown.bind("<<ComboboxSelected>>", lambda e: self._load_sheet_options())

        tk.Label(selector, text="Sheet", font=("Segoe UI", 10, "bold"), bg="#f5f5f5").pack(side="left", padx=(0, 8))
        self._sheet_var = tk.StringVar()
        self._sheet_dropdown = ttk.Combobox(selector, textvariable=self._sheet_var, state="readonly", width=28)
        self._sheet_dropdown.pack(side="left")
        self._sheet_dropdown.bind("<<ComboboxSelected>>", lambda e: self._load_results())

        table_frame = tk.Frame(self, bg="#f5f5f5")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self._tree = ttk.Treeview(table_frame, show="headings", style="Audit.Treeview")
        self._tree.grid(row=0, column=0, sticky="nsew")

        self._tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
        self._tree_scroll.grid(row=0, column=1, sticky="ns")

        self._tree_x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self._tree.xview)
        self._tree_x_scroll.grid(row=1, column=0, sticky="ew")

        self._tree.configure(yscrollcommand=self._tree_scroll.set, xscrollcommand=self._tree_x_scroll.set)
        self._tree.tag_configure("odd", background="#ffffff")
        self._tree.tag_configure("even", background="#eef3f8")

        self._load_workbooks()
        self._load_sheet_options()
        self._load_results()

    def _load_workbooks(self):
        self._workbooks = list_audit_workbooks()

        if not self._workbooks:
            self._file_dropdown["values"] = []
            self._file_var.set("")
            return

        names = list(self._workbooks.keys())
        self._file_dropdown["values"] = names
        preferred_name = next(
            (name for name, path in self._workbooks.items() if self._preferred_workbook and path == self._preferred_workbook),
            None,
        )
        if preferred_name is not None:
            self._file_var.set(preferred_name)
        elif names and self._file_var.get() not in names:
            self._file_var.set(names[0])

    def _load_sheet_options(self):
        workbook = self._selected_workbook()
        if workbook is None:
            self._sheet_dropdown["values"] = []
            self._sheet_var.set("")
            return
        try:
            xl = pd.ExcelFile(workbook)
            self._sheet_dropdown["values"] = xl.sheet_names
            if xl.sheet_names and self._sheet_var.get() not in xl.sheet_names:
                self._sheet_var.set(xl.sheet_names[0])
        except Exception:
            self._sheet_dropdown["values"] = []
            self._sheet_var.set("")

    def _selected_workbook(self) -> Path | None:
        name = self._file_var.get()
        if not name:
            return None
        path = self._workbooks.get(name)
        return path if path and path.exists() else None

    def _load_results(self):
        workbook = self._selected_workbook()
        if workbook is None:
            self._show_message("No audit workbook found.")
            return
        sheet = self._sheet_var.get()
        if not sheet:
            self._show_message("No audit sheet selected.")
            return
        try:
            df = pd.read_excel(workbook, sheet_name=sheet, engine="openpyxl")
            self._populate_table(df)
        except Exception as exc:
            self._show_message(f"Failed to load audit sheet: {exc}")

    def _populate_table(self, df: pd.DataFrame):
        self._tree.delete(*self._tree.get_children())
        if df.empty:
            self._show_message("No data in selected audit sheet.")
            return

        df = df.fillna("")
        columns = list(df.columns)
        self._tree["columns"] = columns
        for col in columns:
            self._tree.heading(col, text=col)
            width = min(max(len(str(col)) * 8, 100), 260)
            self._tree.column(col, width=width, anchor="w", stretch=True)

        for idx, (_, row) in enumerate(df.iterrows()):
            tag = "even" if idx % 2 else "odd"
            values = [self._format_value(value) for value in row.tolist()]
            self._tree.insert("", "end", values=values, tags=(tag,))

    def _show_message(self, message: str):
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = ["Message"]
        self._tree.heading("Message", text="Message")
        self._tree.column("Message", width=700, anchor="w")
        self._tree.insert("", "end", values=(message,))

    @staticmethod
    def _format_value(value):
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value