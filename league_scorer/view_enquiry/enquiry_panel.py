import tkinter as tk
from tkinter import ttk
from typing import Optional

from ..session_config import config as session_config
from ..settings import settings

# Lightweight placeholder for the new Enquiry panel.
class RunnerClubEnquiryPanel(tk.Frame):
    """Panel to search published results by runner or club.

    This initial implementation is a minimal scaffold that mirrors the
    constructor signature used by the dashboard (back_callback, initial_runner).
    It provides a simple search field and a `select_runner(name)` method so
    callers (Issue Review / Dashboard) can pre-load a runner selection.
    """

    def __init__(self, parent, back_callback=None, initial_runner: Optional[str] = None):
        super().__init__(parent, bg="#f7f9fb")
        self._back_callback = back_callback
        self._initial_runner = initial_runner
        self._build_ui()
        if initial_runner:
            # best-effort pre-select
            self.select_runner(initial_runner)

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

        search_bar = tk.Frame(self, bg="#f7f9fb", padx=12, pady=12)
        search_bar.pack(fill="x")

        tk.Label(search_bar, text="Search (runner or club):", bg="#f7f9fb").pack(side="left")
        self._search_var = tk.StringVar()
        self._search_entry = ttk.Entry(search_bar, textvariable=self._search_var, width=48)
        self._search_entry.pack(side="left", padx=(8, 8))
        ttk.Button(search_bar, text="Search", command=self._on_search).pack(side="left")

        self._result_frame = tk.Frame(self, bg="#ffffff", padx=12, pady=8)
        self._result_frame.pack(fill="both", expand=True, padx=12, pady=(0,12))

        self._results_list = tk.Listbox(self._result_frame)
        self._results_list.pack(fill="both", expand=True)

    def _on_search(self):
        query = self._search_var.get().strip()
        # Placeholder behaviour: show the query as a single result. Real
        # implementation will query latest Results workbook and show matches.
        self._results_list.delete(0, "end")
        if query:
            self._results_list.insert("end", f"Search for: {query}")

    def select_runner(self, runner_name: str) -> bool:
        """Pre-select a runner name; returns True if selection was applied.

        Current placeholder implementation simply inserts the name into
        the results list and focuses it.
        """
        if not runner_name:
            return False
        self._results_list.delete(0, "end")
        self._results_list.insert("end", f"Runner: {runner_name}")
        try:
            self._results_list.selection_set(0)
            self._results_list.see(0)
        except Exception:
            pass
        return True
