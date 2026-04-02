"""
gui.py - League Management scorer window.

Opens as a Toplevel child of the dashboard.  Input/output paths are taken
directly from session_config so no folder selection is needed here.
Displays a checklist of discovered race files so the user can choose
which races to include before running the pipeline.
"""

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext, simpledialog

from ..exceptions import FatalError
from ..main import LeagueScorer
from ..raceroster_import import (
    SporthiveRaceNotDirectlyImportableError,
    import_raceroster_results,
    import_sporthive_manual_pages,
)
from ..source_loader import discover_race_files

# ---------------------------------------------------------------------------
# Brand colours (match dashboard)
# ---------------------------------------------------------------------------
NAVY    = "#3a4658"
GREEN   = "#2d7a4a"
GREEN_H = "#1f5632"   # hover / active
AMBER   = "#e6a817"
WHITE   = "#ffffff"
LIGHT   = "#f5f5f5"
PANEL   = "#eef0f4"

# Log-level colours (dark console background)
_LOG_COLOURS = {
    "ERROR":   "#f44747",
    "WARNING": "#dca765",
    "INFO":    "#9cdcfe",
    "DEBUG":   "#858585",
    "SUCCESS": "#6fb06f",
    "DIVIDER": "#444455",
}


# ---------------------------------------------------------------------------
# Logging bridge
# ---------------------------------------------------------------------------

class _QueueHandler(logging.Handler):
    def __init__(self, q: queue.Queue) -> None:
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        self.q.put(("log", record.levelname, self.format(record)))


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class LeagueScorerApp(tk.Frame):
    """League scorer panel — embeds inside the dashboard."""

    _SENTINEL_OK    = "done"
    _SENTINEL_FATAL = "fatal"

    def __init__(self, parent: tk.Misc,
                 input_dir: "Path | None" = None,
                 output_dir: "Path | None" = None,
                 year: int | None = None,
                 back_callback=None) -> None:
        super().__init__(parent, bg=LIGHT)

        self._back_callback = back_callback
        self._alive = True

        self._input_dir  = Path(input_dir)  if input_dir  else None
        self._output_dir = Path(output_dir) if output_dir else None
        self._year = year

        self._log_queue: queue.Queue = queue.Queue()
        self._worker: threading.Thread | None = None
        self._handler: _QueueHandler | None = None

        # Discover available race files and build checkbox state
        self._race_files: dict[int, Path] = self._discover_race_files()
        self._race_vars:  dict[int, tk.BooleanVar] = {
            n: tk.BooleanVar(value=True) for n in self._race_files
        }

        self._attach_log_handler()
        self._build_ui()
        self._poll_queue()

    # -------------------------------------------------------------------------
    # Logging
    # -------------------------------------------------------------------------

    def _attach_log_handler(self) -> None:
        self._handler = _QueueHandler(self._log_queue)
        self._handler.setFormatter(logging.Formatter("%(levelname)-8s  %(message)s"))
        root_log = logging.getLogger()
        root_log.addHandler(self._handler)
        root_log.setLevel(logging.INFO)

    def _detach_log_handler(self) -> None:
        if self._handler:
            logging.getLogger().removeHandler(self._handler)
            self._handler = None

    # -------------------------------------------------------------------------
    # Race file discovery
    # -------------------------------------------------------------------------

    def _discover_race_files(self) -> dict:
        """Scan input_dir and return {race_num: Path} for all valid race files."""
        if not self._input_dir or not self._input_dir.is_dir():
            return {}
        return discover_race_files(
            self._input_dir,
            excluded_names=("clubs.xlsx", "wrrl_events.xlsx"),
        )

    # -------------------------------------------------------------------------
    # UI construction
    # -------------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        self._build_header()
        self._build_paths_panel()
        self._race_selector_host = tk.Frame(self, bg=LIGHT)
        self._race_selector_host.pack(side="top", fill="x")
        self._build_race_selector()
        self._build_action_bar()
        self._build_log_frame()

    def _build_header(self) -> None:
        bar = tk.Frame(self, bg=NAVY, height=52)
        bar.pack(side="top", fill="x")
        bar.pack_propagate(False)

        # Back button (left side)
        if self._back_callback:
            tk.Button(
                bar,
                text="\u25c4  Dashboard",
                font=("Segoe UI", 10, "bold"),
                bg=GREEN, fg=WHITE,
                relief="flat", bd=0,
                padx=12, pady=0,
                cursor="hand2",
                command=self._on_back,
                activebackground=GREEN_H, activeforeground=WHITE,
            ).pack(side="right", padx=(0, 8), pady=10)

        tk.Label(
            bar,
            text="Run League Scorer",
            font=("Segoe UI", 14, "bold"),
            bg=NAVY, fg=WHITE,
        ).pack(side="left", padx=16, pady=10)

        # Season badge
        if self._year is not None:
            tk.Label(
                bar,
                text=f"Season {self._year}",
                font=("Segoe UI", 10),
                bg=NAVY, fg="#a0b8d0",
            ).pack(side="left", padx=4)

    def _build_paths_panel(self) -> None:
        panel = tk.Frame(self, bg=PANEL, bd=0)
        panel.pack(side="top", fill="x", padx=14, pady=(10, 4))

        for label, path in [
            ("Input folder:",  self._input_dir),
            ("Output folder:", self._output_dir),
        ]:
            row = tk.Frame(panel, bg=PANEL)
            row.pack(fill="x", padx=10, pady=3)
            tk.Label(row, text=label, width=14, anchor="w",
                     font=("Segoe UI", 9, "bold"),
                     bg=PANEL, fg=NAVY).pack(side="left")
            tk.Label(row, text=str(path) if path else "Not set",
                     font=("Segoe UI", 9),
                     bg=PANEL, fg="#334455" if path else AMBER,
                     anchor="w").pack(side="left", padx=(4, 0))

    def _build_race_selector(self) -> None:
        """Scrollable checklist of discovered race files."""
        for child in self._race_selector_host.winfo_children():
            child.destroy()

        outer = tk.Frame(self._race_selector_host, bg=LIGHT)
        outer.pack(side="top", fill="x", padx=14, pady=(0, 4))
        self._race_selector_frame = outer

        # Header row with Select All / None buttons
        hdr = tk.Frame(outer, bg=LIGHT)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Race files to process:",
                 font=("Segoe UI", 9, "bold"),
                 bg=LIGHT, fg=NAVY).pack(side="left")
        tk.Button(hdr, text="All", font=("Segoe UI", 8),
                  bg=PANEL, fg=NAVY, relief="flat", padx=6, pady=1,
                  cursor="hand2", command=self._select_all,
                  activebackground="#d0d4de").pack(side="right", padx=(2, 0))
        tk.Button(hdr, text="None", font=("Segoe UI", 8),
                  bg=PANEL, fg=NAVY, relief="flat", padx=6, pady=1,
                  cursor="hand2", command=self._select_none,
                  activebackground="#d0d4de").pack(side="right", padx=(0, 2))
        tk.Button(hdr, text="Import URL…", font=("Segoe UI", 8, "bold"),
              bg=GREEN, fg=WHITE, relief="flat", padx=8, pady=1,
              cursor="hand2", command=self._on_import_raceroster,
              activebackground=GREEN_H, activeforeground=WHITE).pack(side="right", padx=(0, 8))

        if not self._race_files:
            tk.Label(outer, text="  No race files found in input folder.",
                     font=("Segoe UI", 9), bg=LIGHT, fg=AMBER).pack(anchor="w")
            return

        # Scrollable canvas for the checkboxes
        canvas_h = min(len(self._race_files) * 24 + 6, 120)
        canvas = tk.Canvas(outer, bg=PANEL, height=canvas_h,
                           highlightthickness=1, highlightbackground="#c0c8d8")
        vsb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="x", expand=True)

        inner = tk.Frame(canvas, bg=PANEL)
        canvas.create_window((0, 0), window=inner, anchor="nw")

        # Two-column grid of checkboxes
        cols = 2
        for idx, (race_num, fp) in enumerate(self._race_files.items()):
            var = self._race_vars[race_num]
            cb  = tk.Checkbutton(
                inner,
                text=fp.name,
                variable=var,
                font=("Segoe UI", 9),
                bg=PANEL, fg="#222233",
                activebackground=PANEL,
                selectcolor=WHITE,
                anchor="w",
            )
            cb.grid(row=idx // cols, column=idx % cols, sticky="w", padx=10, pady=1)

        inner.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _select_all(self) -> None:
        for var in self._race_vars.values():
            var.set(True)

    def _select_none(self) -> None:
        for var in self._race_vars.values():
            var.set(False)

    def _refresh_race_discovery(self) -> None:
        self._race_files = self._discover_race_files()
        self._race_vars = {n: tk.BooleanVar(value=True) for n in self._race_files}
        self._build_race_selector()

    def _prompt_raceroster_import(self) -> Path | None:
        if not self._input_dir or not self._input_dir.is_dir():
            messagebox.showerror(
                "Input not found",
                f"Input folder does not exist:\n{self._input_dir}",
                parent=self,
            )
            return None

        race_url = simpledialog.askstring(
            "Import From Race Roster",
            "Paste the Race Roster race URL:",
            parent=self,
        )
        if not race_url:
            return None

        race_number = simpledialog.askinteger(
            "League Race Number",
            "Enter the league race number for this file (for example 4):",
            parent=self,
            minvalue=1,
            maxvalue=99,
        )
        if race_number is None:
            return None

        race_name = simpledialog.askstring(
            "Race Name",
            "Optional race name for the file title (for example Broad Town 5):",
            parent=self,
        )

        sporthive_race_hint = None
        if "sporthive.com/events/s/" in race_url.lower() and "/race/" not in race_url.lower():
            sporthive_race_hint = simpledialog.askinteger(
                "Sporthive Race ID",
                "This Sporthive link is an event summary. Enter the Race ID from the 'View results' URL (the number after /race/).",
                parent=self,
                minvalue=1,
            )
            if sporthive_race_hint is None:
                return None

        try:
            output_path, count, history_path = import_raceroster_results(
                race_url=race_url,
                input_dir=self._input_dir,
                league_race_number=race_number,
                race_name_override=race_name,
                sporthive_race_id_hint=sporthive_race_hint,
            )
        except SporthiveRaceNotDirectlyImportableError:
            use_manual = messagebox.askyesno(
                "Sporthive Manual Import",
                "This Sporthive race cannot be imported directly via API.\n\n"
                "Switch to manual page-paste mode for this race?",
                parent=self,
            )
            if not use_manual:
                return None
            return self._prompt_manual_sporthive_import(race_url, race_number, race_name)
        except Exception as exc:
            messagebox.showerror("Race Roster import failed", str(exc), parent=self)
            return None

        self._append_log(f"INFO      Imported → {output_path.name} ({count} rows)", tag="INFO")
        self._append_log(f"INFO      History  → {history_path.name}", tag="INFO")
        self._last_import_history_path = history_path
        return output_path

    def _ask_multiline_page_text(self, title: str, prompt: str) -> str | None:
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.geometry("760x520")
        dialog.configure(bg=LIGHT)

        tk.Label(dialog, text=prompt, bg=LIGHT, fg=NAVY, justify="left", anchor="w").pack(
            fill="x", padx=12, pady=(12, 6)
        )

        text_box = scrolledtext.ScrolledText(dialog, wrap="word", font=("Consolas", 10), height=22)
        text_box.pack(fill="both", expand=True, padx=12, pady=6)

        result = {"value": None}

        button_row = tk.Frame(dialog, bg=LIGHT)
        button_row.pack(fill="x", padx=12, pady=(0, 12))

        def _ok() -> None:
            result["value"] = text_box.get("1.0", "end").strip()
            dialog.destroy()

        def _cancel() -> None:
            result["value"] = None
            dialog.destroy()

        tk.Button(button_row, text="Cancel", command=_cancel, bg=PANEL, fg=NAVY, relief="flat", padx=10, pady=4).pack(side="right")
        tk.Button(button_row, text="Use This Page", command=_ok, bg=GREEN, fg=WHITE, relief="flat", padx=12, pady=4).pack(side="right", padx=(0, 8))

        dialog.wait_window()
        return result["value"]

    def _prompt_manual_sporthive_import(self, race_url: str, race_number: int, race_name: str | None) -> Path | None:
        pages: list[str] = []
        page_no = 1

        while True:
            text = self._ask_multiline_page_text(
                title=f"Sporthive Page {page_no}",
                prompt=(
                    f"Paste results table text for Sporthive page {page_no}.\n"
                    "Copy the rows shown on screen (including pipe-delimited lines) and click 'Use This Page'."
                ),
            )
            if text is None:
                return None
            if not text.strip():
                messagebox.showwarning("No content", "No rows detected. Paste the page content and try again.", parent=self)
                continue
            pages.append(text)

            more = messagebox.askyesno(
                "Add Another Page?",
                "Add another Sporthive results page?\n\nChoose No when all pages are pasted.",
                parent=self,
            )
            if not more:
                break
            page_no += 1

        try:
            output_path, count, history_path = import_sporthive_manual_pages(
                race_url=race_url,
                pages_text=pages,
                input_dir=self._input_dir,
                league_race_number=race_number,
                race_name_override=race_name,
            )
        except Exception as exc:
            messagebox.showerror("Sporthive manual import failed", str(exc), parent=self)
            return None

        self._append_log(f"INFO      Imported (manual) → {output_path.name} ({count} rows)", tag="INFO")
        self._append_log(f"INFO      History           → {history_path.name}", tag="INFO")
        self._last_import_history_path = history_path
        return output_path

    def _on_import_raceroster(self) -> None:
        imported = self._prompt_raceroster_import()
        if imported is None:
            return
        self._refresh_race_discovery()
        messagebox.showinfo(
            "Import Complete",
            f"Imported race file:\n{imported}\n\n"
            f"Import history:\n{getattr(self, '_last_import_history_path', '')}",
            parent=self,
        )

    def _build_log_frame(self) -> None:
        frame = tk.Frame(self, bg=LIGHT)
        frame.pack(side="top", fill="both", expand=True, padx=14, pady=(4, 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self._log_box = scrolledtext.ScrolledText(
            frame,
            state="disabled",
            wrap="word",
            font=("Courier New", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground=WHITE,
            relief="flat",
        )
        self._log_box.pack(fill="both", expand=True)

        for tag, colour in _LOG_COLOURS.items():
            self._log_box.tag_config(tag, foreground=colour)

    def _build_action_bar(self) -> None:
        bar = tk.Frame(self, bg=LIGHT, pady=10)
        bar.pack(side="bottom", fill="x", padx=14)

        # Clear log (left side)
        tk.Button(
            bar,
            text="Clear log",
            font=("Segoe UI", 9),
            bg=PANEL, fg=NAVY,
            relief="flat", bd=0,
            padx=10, pady=4,
            cursor="hand2",
            command=self._clear_log,
            activebackground="#d0d4de",
            activeforeground=NAVY,
        ).pack(side="left")

        # Progress bar (centre-right)
        self._progress_var = tk.IntVar(value=0)
        self._progress_frame = tk.Frame(bar, bg=LIGHT)
        self._progress_frame.pack(side="right", padx=(0, 10))
        self._progress_canvas = tk.Canvas(
            self._progress_frame, width=180, height=8,
            bg="#ccccdd", highlightthickness=0, relief="flat",
        )
        self._progress_canvas.pack()
        self._progress_bar_id = self._progress_canvas.create_rectangle(
            0, 0, 0, 8, fill=GREEN, width=0
        )
        self._anim_offset = 0

        # Run button (right side, large green)
        self._run_btn = tk.Button(
            bar,
            text="▶   Run Scorer",
            font=("Segoe UI", 11, "bold"),
            bg=GREEN, fg=WHITE,
            relief="flat", bd=0,
            padx=20, pady=8,
            cursor="hand2",
            command=self._on_run,
            activebackground=GREEN_H,
            activeforeground=WHITE,
        )
        self._run_btn.pack(side="right", padx=(0, 4))

    # -------------------------------------------------------------------------
    # Animated indeterminate progress bar
    # -------------------------------------------------------------------------

    def _start_progress(self) -> None:
        self._anim_offset = 0
        self._anim_running = True
        self._animate_progress()

    def _stop_progress(self) -> None:
        self._anim_running = False
        self._progress_canvas.coords(self._progress_bar_id, 0, 0, 0, 8)

    def _animate_progress(self) -> None:
        if not self._anim_running:
            return
        w = 180
        bar_w = 60
        x0 = self._anim_offset % (w + bar_w) - bar_w
        x1 = x0 + bar_w
        self._progress_canvas.coords(self._progress_bar_id, x0, 0, x1, 8)
        self._anim_offset += 4
        self.after(30, self._animate_progress)

    # -------------------------------------------------------------------------
    # Run logic
    # -------------------------------------------------------------------------

    def _on_run(self) -> None:
        if not self._input_dir or not self._input_dir.is_dir():
            messagebox.showerror(
                "Input not found",
                f"Input folder does not exist:\n{self._input_dir}",
                parent=self,
            )
            return
        if not self._output_dir:
            messagebox.showerror("No output folder",
                                 "Output folder is not configured.", parent=self)
            return
        self._output_dir.mkdir(parents=True, exist_ok=True)

        # Build the selected subset of race files
        selected = {
            n: fp for n, fp in self._race_files.items()
            if self._race_vars[n].get()
        }
        if not selected:
            messagebox.showwarning(
                "No races selected",
                "Please select at least one race file to process.",
                parent=self,
            )
            return

        self._run_btn.config(state="disabled", bg="#557766")
        self._start_progress()

        names = ", ".join(fp.name for fp in selected.values())
        self._append_log("─" * 64, tag="DIVIDER")
        self._append_log(f"INFO      Input   → {self._input_dir}",  tag="INFO")
        self._append_log(f"INFO      Output  → {self._output_dir}", tag="INFO")
        self._append_log(f"INFO      Races   → {names}",            tag="INFO")
        self._append_log("INFO      Pipeline starting…",            tag="INFO")

        self._worker = threading.Thread(
            target=self._run_pipeline,
            args=(self._input_dir, self._output_dir, self._year, selected),
            daemon=True,
        )
        self._worker.start()

    def _run_pipeline(self, input_dir: Path, output_dir: Path,
                      year: int | None,
                      race_files: dict) -> None:
        try:
            if year is None:
                raise FatalError("Season year is not configured.")
            warnings = LeagueScorer(input_dir, output_dir, year).run(race_files=race_files)
            self._log_queue.put((self._SENTINEL_OK, warnings))
        except FatalError as exc:
            self._log_queue.put((self._SENTINEL_FATAL, str(exc)))
        except Exception as exc:
            self._log_queue.put((self._SENTINEL_FATAL, f"Unexpected error: {exc}"))

    # -------------------------------------------------------------------------
    # Queue polling
    # -------------------------------------------------------------------------

    def _poll_queue(self) -> None:
        while True:
            try:
                item = self._log_queue.get_nowait()
            except queue.Empty:
                break

            if isinstance(item, tuple) and item[0] == self._SENTINEL_OK:
                warnings = item[1]
                self._append_log("✔  Pipeline completed successfully.", tag="SUCCESS")
                self._on_done()
                if warnings:
                    self._show_completion_warnings(warnings)

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
                pass  # widget was destroyed before this reschedule

    def _on_done(self) -> None:
        self._stop_progress()
        self._run_btn.config(state="normal", bg=GREEN)

    def _show_completion_warnings(self, warnings: list[str]) -> None:
        self._append_log("!  Run completed with warnings.", tag="WARNING")
        self._append_log("!  DOCX reports were created, but some PDF exports failed.", tag="WARNING")
        for warning in warnings:
            self._append_log(f"!  {warning}", tag="WARNING")

        summary = "Run completed with warnings.\n\nDOCX reports were created, but PDF export failed for:\n"
        messagebox.showwarning(
            "PDF Export Warnings",
            summary + "\n".join(f"- {warning}" for warning in warnings),
            parent=self,
        )

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------

    def _log_is_at_bottom(self) -> bool:
        _, bottom = self._log_box.yview()
        return bottom >= 0.999

    def _append_log(self, message: str, tag: str = "INFO") -> None:
        follow_output = self._log_is_at_bottom()
        self._log_box.config(state="normal")
        if message == getattr(self, "_last_log_msg", None) and tag == getattr(self, "_last_log_tag", None):
            # Same as previous line — update repeat count in-place
            self._last_log_count += 1
            # Replace the last line with the message + repeat annotation
            self._log_box.delete("end-2l", "end-1c")
            self._log_box.insert("end-1c",
                f"{message}  \u00d7{self._last_log_count}\n", tag)
        else:
            self._last_log_msg   = message
            self._last_log_tag   = tag
            self._last_log_count = 1
            self._log_box.insert("end", message + "\n", tag)
        if follow_output:
            self._log_box.see("end")
        self._log_box.config(state="disabled")

    def _clear_log(self) -> None:
        self._log_box.config(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.config(state="disabled")
        self._last_log_msg   = None
        self._last_log_tag   = None
        self._last_log_count = 1

    def _on_back(self) -> None:
        self._alive = False
        self._detach_log_handler()
        if self._back_callback:
            self._back_callback()


# ---------------------------------------------------------------------------
# Standalone entry point (legacy / IDLE use)
# ---------------------------------------------------------------------------

def launch() -> None:
    """Standalone launcher (opens its own root window)."""
    root = tk.Tk()
    root.title("League Management — Run Scorer")
    root.geometry("860x640")
    root.configure(bg=LIGHT)
    app = LeagueScorerApp(root)
    app.pack(fill="both", expand=True)
    root.mainloop()
