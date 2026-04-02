import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from ..audit_cleanser import create_cleansed_race_file
from ..audit import LeagueAuditor
from ..club_loader import load_clubs
from ..common_files import race_discovery_exclusions
from ..exceptions import FatalError
from ..race_processor import extract_race_number
from ..series_consolidation import consolidate_series_files
from ..source_loader import discover_race_files
from .gui import GREEN, GREEN_H, LIGHT, NAVY, PANEL, WHITE, LeagueScorerApp


class LeagueAuditApp(LeagueScorerApp):
    """Runner audit panel supporting single-file and all-file audit runs."""

    def __init__(self, *args, **kwargs) -> None:
        self._completion_callback = kwargs.pop("completion_callback", None)
        # _race_dir: folder the user browses for race files (may differ from session inputs).
        # _input_dir (set by parent): session inputs folder — used for clubs.xlsx / LeagueAuditor.
        self._race_dir: Path | None = None
        self._race_dir_var = tk.StringVar(value="Same as session input folder (default)")
        self._selected_race_file: Path | None = None
        self._selected_race_var = tk.StringVar(value="No race file selected.")
        self._audit_mode_var = tk.StringVar(value="single")
        self._replace_existing_var = tk.BooleanVar(value=True)
        super().__init__(*args, **kwargs)

    def _discover_race_files(self) -> dict:
        """Discover race files from _race_dir (if set) or the session input folder."""
        scan_dir = self._race_dir if self._race_dir and self._race_dir.is_dir() else self._input_dir
        if not scan_dir or not scan_dir.is_dir():
            return {}
        discovered = discover_race_files(scan_dir, excluded_names=race_discovery_exclusions())
        return {
            race_num: path
            for race_num, path in discovered.items()
            if not path.name.lower().endswith(" (audited).xlsx")
        }

    def _build_header(self) -> None:
        bar = tk.Frame(self, bg=NAVY, height=52)
        bar.pack(side="top", fill="x")
        bar.pack_propagate(False)

        if self._back_callback:
            tk.Button(
                bar,
                text="\u25c4  Dashboard",
                font=("Segoe UI", 10, "bold"),
                bg=GREEN,
                fg=WHITE,
                relief="flat",
                bd=0,
                padx=12,
                pady=0,
                cursor="hand2",
                command=self._on_back,
                activebackground=GREEN_H,
                activeforeground=WHITE,
            ).pack(side="right", padx=(0, 8), pady=10)

        tk.Label(
            bar,
            text="Run Runner Audit",
            font=("Segoe UI", 14, "bold"),
            bg=NAVY,
            fg=WHITE,
        ).pack(side="left", padx=16, pady=10)

        if self._year is not None:
            tk.Label(
                bar,
                text=f"Season {self._year}",
                font=("Segoe UI", 10),
                bg=NAVY,
                fg="#a0b8d0",
            ).pack(side="left", padx=4)

    def _build_action_bar(self) -> None:
        super()._build_action_bar()
        self._run_btn.config(text="▶   Run Audit")

    def _build_race_selector(self) -> None:
        for child in self._race_selector_host.winfo_children():
            child.destroy()
        outer = tk.Frame(self._race_selector_host, bg=LIGHT)
        outer.pack(fill="x", padx=14, pady=(0, 4))

        tk.Label(
            outer,
            text="Race results folder:",
            font=("Segoe UI", 9, "bold"),
            bg=LIGHT,
            fg=NAVY,
        ).pack(anchor="w")

        folder_row = tk.Frame(outer, bg=PANEL, highlightthickness=1, highlightbackground="#c0c8d8")
        folder_row.pack(fill="x", pady=(4, 0))

        tk.Label(
            folder_row,
            textvariable=self._race_dir_var,
            font=("Segoe UI", 9),
            bg=PANEL,
            fg="#22313f",
            anchor="w",
            justify="left",
            padx=10,
            pady=10,
        ).pack(side="left", fill="x", expand=True)

        tk.Button(
            folder_row,
            text="Browse Folder…",
            command=self._browse_input_folder,
            font=("Segoe UI", 9, "bold"),
            bg=GREEN,
            fg=WHITE,
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            activebackground=GREEN_H,
            activeforeground=WHITE,
        ).pack(side="right", padx=8, pady=8)

        tk.Label(
            outer,
            text="Race file to audit (single-file mode):",
            font=("Segoe UI", 9, "bold"),
            bg=LIGHT,
            fg=NAVY,
        ).pack(anchor="w", pady=(8, 0))

        row = tk.Frame(outer, bg=PANEL, highlightthickness=1, highlightbackground="#c0c8d8")
        row.pack(fill="x", pady=(4, 0))

        tk.Label(
            row,
            textvariable=self._selected_race_var,
            font=("Segoe UI", 9),
            bg=PANEL,
            fg="#22313f",
            anchor="w",
            justify="left",
            padx=10,
            pady=10,
        ).pack(side="left", fill="x", expand=True)

        tk.Button(
            row,
            text="Browse File…",
            command=self._browse_race_file,
            font=("Segoe UI", 9, "bold"),
            bg=GREEN,
            fg=WHITE,
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            activebackground=GREEN_H,
            activeforeground=WHITE,
        ).pack(side="right", padx=8, pady=8)

        tk.Label(
            outer,
            text="Choose a folder first. In All mode, all discovered race files in that folder are audited.",
            font=("Segoe UI", 8),
            bg=LIGHT,
            fg="#59687a",
        ).pack(anchor="w", pady=(4, 0))

        mode_row = tk.Frame(outer, bg=LIGHT)
        mode_row.pack(fill="x", pady=(8, 0))

        tk.Label(
            mode_row,
            text="Audit scope:",
            font=("Segoe UI", 9, "bold"),
            bg=LIGHT,
            fg=NAVY,
        ).pack(side="left")

        tk.Radiobutton(
            mode_row,
            text="Selected file",
            variable=self._audit_mode_var,
            value="single",
            font=("Segoe UI", 9),
            bg=LIGHT,
            fg="#22313f",
            selectcolor=WHITE,
            activebackground=LIGHT,
            activeforeground="#22313f",
        ).pack(side="left", padx=(8, 0))

        tk.Radiobutton(
            mode_row,
            text="All discovered race files",
            variable=self._audit_mode_var,
            value="all",
            font=("Segoe UI", 9),
            bg=LIGHT,
            fg="#22313f",
            selectcolor=WHITE,
            activebackground=LIGHT,
            activeforeground="#22313f",
        ).pack(side="left", padx=(8, 0))

        tk.Checkbutton(
            mode_row,
            text="Replace existing audited files",
            variable=self._replace_existing_var,
            font=("Segoe UI", 9),
            bg=LIGHT,
            fg="#22313f",
            selectcolor=WHITE,
            activebackground=LIGHT,
            activeforeground="#22313f",
        ).pack(side="left", padx=(16, 0))

        consolidate_row = tk.Frame(outer, bg=LIGHT)
        consolidate_row.pack(fill="x", pady=(10, 0))

        tk.Label(
            consolidate_row,
            text="Series consolidation:",
            font=("Segoe UI", 9, "bold"),
            bg=LIGHT,
            fg=NAVY,
        ).pack(side="left")

        tk.Button(
            consolidate_row,
            text="Consolidate Series…",
            command=self._consolidate_series,
            font=("Segoe UI", 9, "bold"),
            bg=GREEN,
            fg=WHITE,
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            activebackground=GREEN_H,
            activeforeground=WHITE,
        ).pack(side="left", padx=(10, 0))

        tk.Label(
            outer,
            text="Use this for multi-leg files such as Westbury 5k Series #1/#2/#3. The selected files are merged into one Consolidated workbook and the originals are moved into a series folder.",
            font=("Segoe UI", 8),
            bg=LIGHT,
            fg="#59687a",
            wraplength=920,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

    def _browse_input_folder(self) -> None:
        initial_dir = str(self._race_dir or self._input_dir) if (self._race_dir or self._input_dir) else None
        selected_dir = filedialog.askdirectory(
            parent=self,
            title="Select Race Results Folder",
            initialdir=initial_dir,
            mustexist=True,
        )
        if not selected_dir:
            return

        chosen = Path(selected_dir)
        if not chosen.is_dir():
            messagebox.showerror("Folder not found", f"Selected folder does not exist:\n{chosen}", parent=self)
            return

        self._race_dir = chosen
        self._race_dir_var.set(str(chosen))

        self._selected_race_file = None
        self._selected_race_var.set("No race file selected.")
        self._refresh_race_discovery()

    def _browse_race_file(self) -> None:
        if not self._input_dir or not self._input_dir.is_dir():
            messagebox.showerror("Input not found", f"Input folder does not exist:\n{self._input_dir}", parent=self)
            return

        path_str = filedialog.askopenfilename(
            parent=self,
            title="Select Race File To Audit",
            initialdir=str(self._input_dir),
            filetypes=[("Race workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not path_str:
            return

        path = Path(path_str)
        try:
            path.relative_to(self._input_dir)
        except ValueError:
            messagebox.showerror(
                "Wrong folder",
                f"Select a race file from the active input folder only:\n{self._input_dir}",
                parent=self,
            )
            return

        if path.name.lower() in {"clubs.xlsx", "wrrl_events.xlsx", "name_corrections.xlsx"}:
            messagebox.showerror("Invalid file", "Select a race result file, not a lookup or events workbook.", parent=self)
            return

        race_num = extract_race_number(path.stem)
        if race_num is None:
            messagebox.showerror(
                "Invalid race file",
                "The selected file name does not contain a valid race number.",
                parent=self,
            )
            return

        self._selected_race_file = path
        self._selected_race_var.set(path.name)

    def _consolidate_series(self) -> None:
        if not self._input_dir or not self._input_dir.is_dir():
            messagebox.showerror("Input not found", f"Input folder does not exist:\n{self._input_dir}", parent=self)
            return

        selected_paths = filedialog.askopenfilenames(
            parent=self,
            title="Select Series Files To Consolidate",
            initialdir=str(self._input_dir),
            filetypes=[("Race workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not selected_paths:
            return

        try:
            result = consolidate_series_files(
                [Path(path_str) for path_str in selected_paths],
                self._input_dir,
            )
        except Exception as exc:
            messagebox.showerror("Series consolidation failed", str(exc), parent=self)
            return

        self._selected_race_file = result.consolidated_path
        self._selected_race_var.set(result.consolidated_path.name)
        self._append_log(
            f"INFO      Consolidated → {result.consolidated_path.name}",
            tag="INFO",
        )
        self._append_log(
            f"INFO      Moved source files into → {result.archive_dir.name}",
            tag="INFO",
        )
        messagebox.showinfo(
            "Series Consolidated",
            "\n\n".join(
                [
                    f"Consolidated workbook created:\n{result.consolidated_path}",
                    f"Source files moved into:\n{result.archive_dir}",
                    "The consolidated workbook is now selected for audit.",
                ]
            ),
            parent=self,
        )

    def _on_run(self) -> None:
        if not self._input_dir or not self._input_dir.is_dir():
            messagebox.showerror("Input not found", f"Input folder does not exist:\n{self._input_dir}", parent=self)
            return
        if not self._output_dir:
            messagebox.showerror("No output folder", "Output folder is not configured.", parent=self)
            return
        (self._output_dir / "audit").mkdir(parents=True, exist_ok=True)

        selected: dict[int, Path] = {}
        mode = self._audit_mode_var.get().strip().lower()
        if mode == "all":
            selected = dict(self._race_files)
            if not selected:
                messagebox.showwarning(
                    "No race files found",
                    "No discoverable race files were found in the active input folder.",
                    parent=self,
                )
                return
        else:
            if self._selected_race_file is None:
                messagebox.showwarning("No race file selected", "Browse to a single race file before running audit.", parent=self)
                return

            race_num = extract_race_number(self._selected_race_file.stem)
            if race_num is None:
                messagebox.showerror("Invalid race file", "The selected race file no longer has a valid race number.", parent=self)
                return

            selected = {race_num: self._selected_race_file}

        self._run_btn.config(state="disabled", bg="#557766")
        self._start_progress()

        self._append_log("─" * 64, tag="DIVIDER")
        self._append_log(f"INFO      Input   → {self._input_dir}", tag="INFO")
        self._append_log(f"INFO      Output  → {self._output_dir / 'audit'}", tag="INFO")
        if mode == "all":
            self._append_log(f"INFO      Scope   → All files ({len(selected)})", tag="INFO")
        else:
            self._append_log(f"INFO      Scope   → Single file", tag="INFO")
            self._append_log(f"INFO      Race    → {self._selected_race_file.name}", tag="INFO")
        self._append_log(
            f"INFO      Replace → {'Yes' if self._replace_existing_var.get() else 'No'}",
            tag="INFO",
        )
        self._append_log("INFO      Audit starting…", tag="INFO")

        self._worker = threading.Thread(
            target=self._run_pipeline,
            args=(self._input_dir, self._output_dir, self._year, selected, self._replace_existing_var.get()),
            # _input_dir is always the session inputs folder (clubs.xlsx / LeagueAuditor);
            # race files in `selected` may come from a different folder.
            daemon=True,
        )
        self._worker.start()

    def _run_pipeline(
        self,
        input_dir: Path,
        output_dir: Path,
        year: int | None,
        race_files: dict,
        overwrite_existing: bool,
    ) -> None:
        try:
            if year is None:
                raise FatalError("Season year is not configured.")
            raw_to_preferred, club_info = load_clubs(input_dir / "clubs.xlsx")
            preferred_clubs = sorted(club_info)

            audited_paths = []
            audited_race_files = {}
            for filepath in race_files.values():
                audited_path = create_cleansed_race_file(
                    filepath,
                    raw_to_preferred,
                    preferred_clubs,
                    overwrite_existing=overwrite_existing,
                )
                self._log_queue.put(("log", "INFO", f"INFO      Audited → {audited_path.name}"))
                audited_paths.append(audited_path)
                race_num = extract_race_number(audited_path.stem)
                if race_num is None:
                    raise FatalError(f"Audited file does not contain a valid race number: {audited_path.name}")
                audited_race_files[race_num] = audited_path

            workbook_path = LeagueAuditor(input_dir, output_dir, year).run(race_files=audited_race_files)
            self._log_queue.put((self._SENTINEL_OK, workbook_path, audited_paths))
        except FatalError as exc:
            self._log_queue.put((self._SENTINEL_FATAL, str(exc)))
        except Exception as exc:
            self._log_queue.put((self._SENTINEL_FATAL, f"Unexpected error: {exc}"))

    def _poll_queue(self) -> None:
        while True:
            try:
                item = self._log_queue.get_nowait()
            except queue.Empty:
                break

            if isinstance(item, tuple) and item[0] == self._SENTINEL_OK:
                workbook_path = item[1]
                audited_paths = item[2] if len(item) > 2 else []
                self._append_log(f"✔  Audit completed successfully: {workbook_path.name}", tag="SUCCESS")
                self._on_done()
                details = [f"Audit workbook written to:\n{workbook_path}"]
                if audited_paths:
                    details.append("Audited race file(s):")
                    details.extend(str(path) for path in audited_paths)
                messagebox.showinfo(
                    "Audit Complete",
                    "\n\n".join([details[0], "\n".join(details[1:])]) if len(details) > 1 else details[0],
                    parent=self,
                )
                if self._completion_callback is not None:
                    preferred_workbook = audited_paths[0] if audited_paths else workbook_path
                    self._completion_callback(preferred_workbook)

            elif isinstance(item, tuple) and item[0] == self._SENTINEL_FATAL:
                self._append_log(f"✘  FATAL — {item[1]}", tag="ERROR")
                self._on_done()

            elif isinstance(item, tuple) and item[0] == "log":
                _, level, message = item
                tag = level if level in ("ERROR", "WARNING", "DEBUG") else "INFO"
                self._append_log(message, tag=tag)

        if self._alive:
            try:
                self.after(100, self._poll_queue)
            except tk.TclError:
                pass