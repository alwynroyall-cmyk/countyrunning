"""Structured log viewer panel for dashboard settings."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from ..structured_logging import read_structured_events, structured_log_path

WRRL_NAVY = "#3a4658"
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"


class LogViewerPanel(tk.Frame):
    def __init__(self, parent: tk.Misc, back_callback=None, dashboard_callback=None) -> None:
        super().__init__(parent, bg=WRRL_LIGHT)
        self._back_callback = back_callback
        self._dashboard_callback = dashboard_callback
        self._level_var = tk.StringVar(value="ALL")
        self._search_var = tk.StringVar(value="")
        self._status_var = tk.StringVar(value="")
        self._build_ui()
        self.refresh()

    def _build_ui(self) -> None:
        title_row = tk.Frame(self, bg=WRRL_NAVY, padx=16, pady=10)
        title_row.pack(fill="x")
        tk.Label(
            title_row,
            text="Structured Log Viewer",
            font=("Segoe UI", 16, "bold"),
            bg=WRRL_NAVY,
            fg=WRRL_WHITE,
        ).pack(side="left")

        if self._back_callback:
            tk.Button(
                title_row,
                text="\u25c4 Settings",
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

        if self._dashboard_callback:
            tk.Button(
                title_row,
                text="\u25c4 Dashboard",
                font=("Segoe UI", 10, "bold"),
                bg="#dbe1e8",
                fg=WRRL_NAVY,
                relief="flat",
                padx=10,
                pady=4,
                cursor="hand2",
                command=self._dashboard_callback,
            ).pack(side="right", padx=(0, 8))

        controls = tk.Frame(self, bg=WRRL_LIGHT, padx=16, pady=8)
        controls.pack(fill="x")

        tk.Label(controls, text="Level:", bg=WRRL_LIGHT, fg=WRRL_NAVY, font=("Segoe UI", 10, "bold")).pack(side="left")
        level_combo = ttk.Combobox(
            controls,
            textvariable=self._level_var,
            state="readonly",
            width=10,
            values=["ALL", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        )
        level_combo.pack(side="left", padx=(6, 12))
        level_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh())

        tk.Label(controls, text="Search:", bg=WRRL_LIGHT, fg=WRRL_NAVY, font=("Segoe UI", 10, "bold")).pack(side="left")
        search_entry = tk.Entry(controls, textvariable=self._search_var, width=36, font=("Segoe UI", 10))
        search_entry.pack(side="left", padx=(6, 12))
        self._search_var.trace_add("write", lambda *_: self.refresh())

        tk.Button(
            controls,
            text="Refresh",
            command=self.refresh,
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
        ).pack(side="left")

        tk.Button(
            controls,
            text="Clear Log",
            command=self._clear_log,
            bg="#f2dede",
            fg="#7a1f1f",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
        ).pack(side="left", padx=(8, 0))

        table_wrap = tk.Frame(self, bg=WRRL_LIGHT, padx=16, pady=0)
        table_wrap.pack(fill="both", expand=True, pady=(0, 8))

        cols = ("ts", "level", "event", "details")
        self._tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=18)
        self._tree.heading("ts", text="Timestamp (UTC)")
        self._tree.heading("level", text="Level")
        self._tree.heading("event", text="Event")
        self._tree.heading("details", text="Details")

        self._tree.column("ts", width=210, anchor="w")
        self._tree.column("level", width=80, anchor="center")
        self._tree.column("event", width=230, anchor="w")
        self._tree.column("details", width=600, anchor="w")

        ybar = ttk.Scrollbar(table_wrap, orient="vertical", command=self._tree.yview)
        xbar = ttk.Scrollbar(table_wrap, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        self._tree.pack(side="top", fill="both", expand=True)
        ybar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")

        footer = tk.Frame(self, bg=WRRL_NAVY, padx=16, pady=6)
        footer.pack(fill="x")
        tk.Label(
            footer,
            textvariable=self._status_var,
            bg=WRRL_NAVY,
            fg="#a0b0c0",
            font=("Segoe UI", 9, "italic"),
            anchor="w",
        ).pack(fill="x")

    def _clear_log(self) -> None:
        path = structured_log_path()
        if not path.exists():
            self.refresh()
            return
        if not messagebox.askyesno("Clear Structured Log", f"Delete all entries from:\n{path}", parent=self):
            return
        try:
            path.unlink()
        except OSError as exc:
            messagebox.showerror("Clear Failed", str(exc), parent=self)
        self.refresh()

    def refresh(self) -> None:
        for item in self._tree.get_children():
            self._tree.delete(item)

        rows = read_structured_events(limit=1000)
        level_filter = self._level_var.get().strip().upper()
        needle = self._search_var.get().strip().lower()

        visible = 0
        for row in rows:
            level = str(row.get("level", "")).upper()
            event = str(row.get("event", ""))
            ts = str(row.get("ts", ""))

            details_map = {k: v for k, v in row.items() if k not in {"ts", "level", "event"}}
            details = " | ".join(f"{k}={details_map[k]}" for k in sorted(details_map))

            if level_filter and level_filter != "ALL" and level != level_filter:
                continue

            joined = f"{ts} {level} {event} {details}".lower()
            if needle and needle not in joined:
                continue

            self._tree.insert("", "end", values=(ts, level, event, details))
            visible += 1

        if visible == 0:
            self._tree.insert(
                "",
                "end",
                values=("", "", "No events yet", "Run scorer/import/edit actions, then click Refresh."),
            )

        self._status_var.set(f"Showing {visible} event(s). Log file: {structured_log_path()}")
