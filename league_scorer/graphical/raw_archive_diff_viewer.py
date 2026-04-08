from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from ..raw_archive_diff_service import (
    DiffRow,
    build_side_by_side_diff,
    list_comparable_file_pairs,
    load_comparable_lines,
)
from ..session_config import config as session_config
from .dashboard import WRRL_GREEN, WRRL_LIGHT, WRRL_NAVY, WRRL_WHITE


class RawArchiveDiffPanel(tk.Frame):
    def __init__(self, parent, back_callback=None):
        super().__init__(parent, bg=WRRL_LIGHT)
        self._back_callback = back_callback
        self._pairs_by_name = {}
        self._pair_names: list[str] = []
        self._status_var = tk.StringVar(value="Select a file to compare.")
        self._file_var = tk.StringVar()
        self._build_ui()
        self._load_file_pairs()

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=WRRL_NAVY, padx=14, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Raw Data vs Archive Diff",
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

        tk.Label(
            controls,
            text="File:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        ).pack(side="left", padx=(0, 8))

        self._file_combo = ttk.Combobox(
            controls,
            textvariable=self._file_var,
            state="readonly",
            width=64,
        )
        self._file_combo.pack(side="left")
        self._file_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_selected_diff())

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
            command=self._load_file_pairs,
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            controls,
            text="Open Raw",
            font=("Segoe UI", 10),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=lambda: self._open_selected_file(kind="raw"),
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            controls,
            text="Open Archive",
            font=("Segoe UI", 10),
            bg="#dbe1e8",
            fg=WRRL_NAVY,
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=lambda: self._open_selected_file(kind="archive"),
        ).pack(side="left", padx=(8, 0))

        tk.Label(
            self,
            textvariable=self._status_var,
            font=("Segoe UI", 9, "italic"),
            bg=WRRL_LIGHT,
            fg="#5f6d7b",
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 8))

        panes = tk.PanedWindow(self, orient="horizontal", sashrelief="flat", bg=WRRL_LIGHT)
        panes.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        left_frame = tk.Frame(panes, bg=WRRL_LIGHT)
        right_frame = tk.Frame(panes, bg=WRRL_LIGHT)
        panes.add(left_frame, stretch="always")
        panes.add(right_frame, stretch="always")

        tk.Label(left_frame, text="Raw Data", font=("Segoe UI", 11, "bold"), bg=WRRL_LIGHT, fg=WRRL_NAVY).pack(anchor="w", pady=(0, 6))
        tk.Label(right_frame, text="Archive", font=("Segoe UI", 11, "bold"), bg=WRRL_LIGHT, fg=WRRL_NAVY).pack(anchor="w", pady=(0, 6))

        left_box = tk.Frame(left_frame, bg=WRRL_LIGHT)
        left_box.pack(fill="both", expand=True)
        right_box = tk.Frame(right_frame, bg=WRRL_LIGHT)
        right_box.pack(fill="both", expand=True)

        self._left_text = tk.Text(left_box, wrap="none", font=("Consolas", 10), bg="#ffffff", fg="#22313f")
        self._right_text = tk.Text(right_box, wrap="none", font=("Consolas", 10), bg="#ffffff", fg="#22313f")
        self._left_text.grid(row=0, column=0, sticky="nsew")
        self._right_text.grid(row=0, column=0, sticky="nsew")

        left_scroll_y = ttk.Scrollbar(left_box, orient="vertical", command=self._sync_yview)
        right_scroll_y = ttk.Scrollbar(right_box, orient="vertical", command=self._sync_yview)
        left_scroll_y.grid(row=0, column=1, sticky="ns")
        right_scroll_y.grid(row=0, column=1, sticky="ns")
        left_scroll_x = ttk.Scrollbar(left_box, orient="horizontal", command=self._left_text.xview)
        right_scroll_x = ttk.Scrollbar(right_box, orient="horizontal", command=self._right_text.xview)
        left_scroll_x.grid(row=1, column=0, sticky="ew")
        right_scroll_x.grid(row=1, column=0, sticky="ew")

        left_box.grid_rowconfigure(0, weight=1)
        left_box.grid_columnconfigure(0, weight=1)
        right_box.grid_rowconfigure(0, weight=1)
        right_box.grid_columnconfigure(0, weight=1)

        self._left_text.configure(yscrollcommand=left_scroll_y.set, xscrollcommand=left_scroll_x.set)
        self._right_text.configure(yscrollcommand=right_scroll_y.set, xscrollcommand=right_scroll_x.set)

        self._configure_tags(self._left_text)
        self._configure_tags(self._right_text)

        self._left_text.bind("<MouseWheel>", self._on_mousewheel)
        self._right_text.bind("<MouseWheel>", self._on_mousewheel)

    def _configure_tags(self, widget: tk.Text) -> None:
        widget.tag_configure("same", background="#ffffff")
        widget.tag_configure("replace", background="#fff0c9")
        widget.tag_configure("delete", background="#fde4e4")
        widget.tag_configure("insert", background="#e4f6e8")

    def _load_file_pairs(self) -> None:
        input_dir = session_config.input_dir
        if input_dir is None:
            self._pairs_by_name = {}
            self._pair_names = []
            self._file_combo["values"] = []
            self._file_var.set("")
            self._status_var.set("Inputs are not configured.")
            self._render_message("Set the season data root before comparing raw and archive files.")
            return

        pairs = list_comparable_file_pairs(input_dir)
        self._pairs_by_name = {pair.filename: pair for pair in pairs}
        self._pair_names = [pair.filename for pair in pairs]
        self._file_combo["values"] = self._pair_names

        if not self._pair_names:
            self._file_var.set("")
            self._status_var.set("No matching raw/archive file pairs found.")
            self._render_message("No files with the same name exist in both raw_data and raw_data_archive.")
            return

        current = self._file_var.get()
        if current not in self._pairs_by_name:
            self._file_var.set(self._pair_names[0])
        self._load_selected_diff()

    def _load_selected_diff(self) -> None:
        pair = self._selected_pair()
        if pair is None:
            self._render_message("Choose a file to compare.")
            return

        self._status_var.set(f"Loading {pair.filename}\u2026")
        self._render_message("Loading\u2026")

        def _worker():
            try:
                raw_lines = load_comparable_lines(pair.raw_path)
                archive_lines = load_comparable_lines(pair.archive_path)
                diff_rows = build_side_by_side_diff(raw_lines, archive_lines)
            except Exception as exc:
                self.after(0, lambda: (
                    self._status_var.set(f"Failed to compare {pair.filename}"),
                    self._render_message(f"Could not compare the selected file: {exc}"),
                ))
                return

            differing_rows = sum(1 for row in diff_rows if row.status != "same")

            def _done() -> None:
                self._status_var.set(
                    f"{pair.filename}: {differing_rows} differing line(s), "
                    f"raw lines {len(raw_lines)}, archive lines {len(archive_lines)}"
                )
                self._render_diff_rows(diff_rows)

            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    def _selected_pair(self):
        filename = self._file_var.get().strip()
        if not filename:
            return None
        return self._pairs_by_name.get(filename)

    def _render_message(self, message: str) -> None:
        self._left_text.config(state="normal")
        self._right_text.config(state="normal")
        self._left_text.delete("1.0", "end")
        self._right_text.delete("1.0", "end")
        self._left_text.insert("1.0", message)
        self._right_text.insert("1.0", message)
        self._left_text.config(state="disabled")
        self._right_text.config(state="disabled")

    def _render_diff_rows(self, rows: list[DiffRow]) -> None:
        self._left_text.config(state="normal")
        self._right_text.config(state="normal")
        self._left_text.delete("1.0", "end")
        self._right_text.delete("1.0", "end")

        for row in rows:
            left_no = "" if row.left_line_no is None else f"{row.left_line_no:>5}"
            right_no = "" if row.right_line_no is None else f"{row.right_line_no:>5}"
            left_line = f"{left_no} | {row.left_text}\n"
            right_line = f"{right_no} | {row.right_text}\n"

            left_start = self._left_text.index("end-1c")
            right_start = self._right_text.index("end-1c")
            self._left_text.insert("end", left_line)
            self._right_text.insert("end", right_line)
            self._left_text.tag_add(row.status, left_start, f"{left_start} lineend")
            self._right_text.tag_add(row.status, right_start, f"{right_start} lineend")

        self._left_text.config(state="disabled")
        self._right_text.config(state="disabled")
        self._left_text.yview_moveto(0)
        self._right_text.yview_moveto(0)

    def _sync_yview(self, *args) -> None:
        self._left_text.yview(*args)
        self._right_text.yview(*args)

    def _on_mousewheel(self, event) -> str:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        self._left_text.yview_scroll(delta, "units")
        self._right_text.yview_scroll(delta, "units")
        return "break"

    def _open_selected_file(self, *, kind: str) -> None:
        pair = self._selected_pair()
        if pair is None:
            messagebox.showwarning("No File Selected", "Choose a file pair first.", parent=self)
            return

        path = pair.raw_path if kind == "raw" else pair.archive_path
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except OSError as exc:
            messagebox.showerror("Open Failed", f"Could not open file: {exc}", parent=self)