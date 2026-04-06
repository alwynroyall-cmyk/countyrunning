"""
events_viewer.py - Tkinter window for viewing the WRRL Championship Events schedule.

Displays the loaded events in a sortable table with colour-coded status rows.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from PIL import ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

from ..events_loader import (
    EventsSchedule,
    EventEntry,
    STATUS_CONFIRMED,
    STATUS_PROVISIONAL,
    STATUS_TBC,
)
from .timeline_generator import generate_timeline


# ---------------------------------------------------------------------------
# Brand colours (shared with dashboard)
# ---------------------------------------------------------------------------

WRRL_NAVY  = "#3a4658"
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"

# Status row colours
_STATUS_BG: dict[str, str] = {
    STATUS_CONFIRMED.lower():   "#e8f5e9",   # soft green
    STATUS_PROVISIONAL.lower(): "#e3f2fd",   # soft blue
    STATUS_TBC.lower():         "#fff8e1",   # soft amber
}
_STATUS_FG: dict[str, str] = {
    STATUS_CONFIRMED.lower():   "#1b5e20",
    STATUS_PROVISIONAL.lower(): "#0d47a1",
    STATUS_TBC.lower():         "#e65100",
}

# Treeview columns: (id, heading, width, anchor)
_COLUMNS = [
    ("race_ref",        "Ref",         70,  "center"),
    ("event_name",      "Event",      200,  "w"),
    ("distance",        "Distance",    75,  "center"),
    ("location",        "Location",   120,  "w"),
    ("organiser",       "Organiser",  130,  "w"),
    ("date_type",       "Date Type",   80,  "center"),
    ("scheduled_dates", "Dates",      160,  "w"),
    ("entry_fee",       "Fee",         60,  "center"),
    ("scoring_basis",   "Scoring",     85,  "center"),
    ("status",          "Status",      90,  "center"),
]



class EventsViewerPanel(tk.Frame):
    """Embedded panel for viewing the Championship Events schedule."""

    def __init__(self, parent: tk.Misc, schedule: EventsSchedule,
                 year: int = 2026, images_dir: Path | None = None,
                 output_dir: Path | None = None) -> None:
        super().__init__(parent, bg=WRRL_LIGHT)
        self._schedule   = schedule
        self._year       = year
        self._images_dir = images_dir
        self._output_dir = output_dir
        self._build_ui()
        self._populate(schedule.events)

    # -------------------------------------------------------------------------
    # UI construction
    # -------------------------------------------------------------------------

    def _build_ui(self) -> None:
        self._build_toolbar()
        self._build_summary_bar()
        self._build_table()
        self._build_status_bar()

    def _build_toolbar(self) -> None:
        bar = tk.Frame(self, bg=WRRL_NAVY, height=48)
        bar.pack(side="top", fill="x")
        bar.pack_propagate(False)

        tk.Label(
            bar,
            text="Championship Events Schedule",
            font=("Segoe UI", 14, "bold"),
            bg=WRRL_NAVY,
            fg=WRRL_WHITE,
        ).pack(side="left", padx=16, pady=8)

        # Generate Timeline button
        tl_btn = tk.Button(
            bar,
            text="\U0001f4c5  Generate Timeline",
            font=("Segoe UI", 9, "bold"),
            bg=WRRL_GREEN,
            fg=WRRL_WHITE,
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._on_generate_timeline,
            activebackground="#1f5632",
            activeforeground=WRRL_WHITE,
        )
        tl_btn.pack(side="right", padx=4, pady=8)

        # Source path label
        if self._schedule.source_path:
            src = self._schedule.source_path.name
            tk.Label(
                bar,
                text=f"Source: {src}",
                font=("Segoe UI", 9),
                bg=WRRL_NAVY,
                fg="#a0b0c0",
            ).pack(side="right", padx=4, pady=8)

    def _build_summary_bar(self) -> None:
        s = self._schedule
        total       = len(s.events)
        confirmed   = len(s.confirmed)
        provisional = len(s.provisional)
        tbc         = len(s.tbc)

        bar = tk.Frame(self, bg="#e8eaf0", height=32)
        bar.pack(side="top", fill="x")
        bar.pack_propagate(False)

        summary = (
            f"  Total: {total}    "
            f"Confirmed: {confirmed}    "
            f"Provisional: {provisional}    "
            f"TBC: {tbc}"
        )
        tk.Label(
            bar,
            text=summary,
            font=("Segoe UI", 9),
            bg="#e8eaf0",
            fg=WRRL_NAVY,
            anchor="w",
        ).pack(side="left", padx=12, pady=4)

    def _build_table(self) -> None:
        frame = tk.Frame(self, bg=WRRL_LIGHT)
        frame.pack(side="top", fill="both", expand=True, padx=10, pady=8)

        col_ids = [c[0] for c in _COLUMNS]
        self._tree = ttk.Treeview(
            frame,
            columns=col_ids,
            show="headings",
            selectmode="browse",
        )

        for col_id, heading, width, anchor in _COLUMNS:
            self._tree.heading(col_id, text=heading,
                               command=lambda c=col_id: self._sort_by(c))
            self._tree.column(col_id, width=width, anchor=anchor, minwidth=30)

        # Configure status colour tags
        for status_key, bg in _STATUS_BG.items():
            self._tree.tag_configure(
                status_key,
                background=bg,
                foreground=_STATUS_FG[status_key],
            )

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_status_bar(self) -> None:
        bar = tk.Frame(self, bg="#dde0e8", height=24)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        self._status_var = tk.StringVar(value="Click a column header to sort.")
        tk.Label(
            bar,
            textvariable=self._status_var,
            font=("Segoe UI", 8),
            bg="#dde0e8",
            fg="#555566",
            anchor="w",
        ).pack(side="left", padx=8)

    # -------------------------------------------------------------------------
    # Data population
    # -------------------------------------------------------------------------

    def _populate(self, events: list[EventEntry]) -> None:
        """Clear and re-populate the treeview."""
        for item in self._tree.get_children():
            self._tree.delete(item)

        for ev in events:
            tag = ev.status.lower()
            self._tree.insert(
                "",
                "end",
                values=(
                    ev.race_ref,
                    ev.event_name,
                    ev.distance,
                    ev.location,
                    ev.organiser,
                    ev.date_type,
                    ev.scheduled_dates,
                    ev.entry_fee,
                    ev.scoring_basis,
                    ev.status,
                ),
                tags=(tag,),
            )

    # -------------------------------------------------------------------------
    # Sorting
    # -------------------------------------------------------------------------

    _sort_reverse: dict[str, bool] = {}

    def _sort_by(self, col_id: str) -> None:
        reverse = self._sort_reverse.get(col_id, False)
        items = [
            (self._tree.set(k, col_id), k)
            for k in self._tree.get_children("")
        ]
        try:
            items.sort(
                key=lambda t: int(t[0]) if t[0].isdigit() else t[0].lower(),
                reverse=reverse,
            )
        except Exception:
            items.sort(key=lambda t: t[0].lower(), reverse=reverse)

        for idx, (_, k) in enumerate(items):
            self._tree.move(k, "", idx)

        self._sort_reverse[col_id] = not reverse
        self._status_var.set(f"Sorted by '{col_id}' {'desc' if reverse else 'asc'}.")

    # -------------------------------------------------------------------------
    # Generate Timeline
    # -------------------------------------------------------------------------

    def _on_generate_timeline(self) -> None:
        """Generate the season timeline PNG and show a preview window."""
        if not self._schedule.events:
            messagebox.showwarning("No Events", "No events are loaded.", parent=self)
            return

        # Determine save path
        default_name = f"WRRL_season_timeline_{self._year}.png"
        if self._output_dir:
            default_path = self._output_dir / default_name
        else:
            default_path = None

        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Save Timeline As",
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialdir=str(self._output_dir) if self._output_dir else None,
            initialfile=default_name,
        )
        if not save_path:
            return  # user cancelled

        try:
            img = generate_timeline(
                self._schedule,
                year=self._year,
                output_path=Path(save_path),
                images_dir=self._images_dir,
            )
        except Exception as exc:
            messagebox.showerror("Timeline Error",
                                 f"Could not generate timeline:\n{exc}", parent=self)
            return

        # Show preview
        self._show_timeline_preview(img, Path(save_path))

    def _show_timeline_preview(self, img, saved_path: Path) -> None:
        """Open a simple Toplevel window displaying the generated timeline image."""
        win = tk.Toplevel(self)
        win.title(f"Season Timeline - {saved_path.name}")
        win.configure(bg=WRRL_LIGHT)
        win.resizable(True, True)

        # Toolbar
        bar = tk.Frame(win, bg=WRRL_NAVY, height=40)
        bar.pack(side="top", fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text=f"Saved: {saved_path}",
                 font=("Segoe UI", 9), bg=WRRL_NAVY, fg="#a0b0c0").pack(
                     side="left", padx=12)

        # Scrollable canvas
        frame = tk.Frame(win, bg=WRRL_LIGHT)
        frame.pack(fill="both", expand=True)
        hsb = tk.Scrollbar(frame, orient="horizontal")
        vsb = tk.Scrollbar(frame, orient="vertical")
        canvas = tk.Canvas(frame, bg="#e8e8e8",
                           xscrollcommand=hsb.set,
                           yscrollcommand=vsb.set)
        hsb.config(command=canvas.xview)
        vsb.config(command=canvas.yview)
        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right",  fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        if _PIL_AVAILABLE:
            photo = ImageTk.PhotoImage(img)
            canvas.create_image(0, 0, anchor="nw", image=photo)
            canvas.config(scrollregion=(0, 0, img.width, img.height))
            canvas._photo_ref = photo  # prevent GC
            win.geometry(f"{min(img.width + 20, 1450)}x{min(img.height + 80, 800)}")
        else:
            canvas.create_text(200, 100,
                text="Install Pillow to preview images.\nFile saved to: " + str(saved_path),
                fill=WRRL_NAVY, font=("Segoe UI", 12))
            win.geometry("500x200")


