from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Dict, List

import pandas as pd

from ..name_lookup import append_name_corrections, read_name_lookup_state
from .club_match_dialog import (
    _append_club_conversions,
    _build_write_result_text,
    _read_club_lookup_state,
    _summarise_selection,
    _set_widget_tree_bg,
    ClubMatchCandidate,
)

_HIGHLIGHT_BG = "#fff1c7"
_DEFAULT_BG = "#f5f5f5"


@dataclass
class NameReviewCandidate:
    current_name: str
    proposed_name: str
    message: str
    confidence: str
    occurrences: str


class ManualReviewDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        club_df: pd.DataFrame | None = None,
        clubs_path: Path | None = None,
        name_df: pd.DataFrame | None = None,
        names_path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.title("Manual Review")
        self.geometry("1080x620")
        self.minsize(920, 460)
        self.configure(bg=_DEFAULT_BG)
        self.transient(parent.winfo_toplevel())

        self._clubs_path = clubs_path
        self._names_path = names_path
        self._club_candidates = _build_club_candidates(club_df)
        self._name_candidates = _build_name_candidates(name_df)
        self._club_vars: Dict[str, tk.BooleanVar] = {}
        self._name_vars: Dict[str, tk.BooleanVar] = {}
        self._name_value_vars: Dict[str, tk.StringVar] = {}
        self._club_row_frames: Dict[str, tk.Frame] = {}
        self._name_row_frames: Dict[str, tk.Frame] = {}

        self._build_ui()
        self.grab_set()

    def _build_ui(self) -> None:
        title = tk.Label(
            self,
            text="Manual Review",
            font=("Segoe UI", 15, "bold"),
            bg=_DEFAULT_BG,
            fg="#22313f",
        )
        title.pack(anchor="w", padx=16, pady=(14, 4))

        subtitle = tk.Label(
            self,
            text="Review club and name suggestions, tick the rows to approve, then confirm before writing lookup updates.",
            font=("Segoe UI", 10),
            bg=_DEFAULT_BG,
            fg="#4a5868",
        )
        subtitle.pack(anchor="w", padx=16, pady=(0, 10))

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        added_tabs = 0

        if self._club_candidates:
            club_tab = tk.Frame(notebook, bg=_DEFAULT_BG)
            notebook.add(club_tab, text=f"Clubs ({len(self._club_candidates)})")
            self._build_club_tab(club_tab)
            added_tabs += 1

        if self._name_candidates:
            name_tab = tk.Frame(notebook, bg=_DEFAULT_BG)
            notebook.add(name_tab, text=f"Names ({len(self._name_candidates)})")
            self._build_name_tab(name_tab)
            added_tabs += 1

        if not added_tabs:
            notebook.destroy()
            tk.Label(
                self,
                text="No manual review suggestions are available in this workbook.",
                font=("Segoe UI", 10),
                bg=_DEFAULT_BG,
                fg="#4a5868",
            ).pack(anchor="w", padx=16, pady=(0, 12))

        action_bar = tk.Frame(self, bg=_DEFAULT_BG)
        action_bar.pack(fill="x", padx=16, pady=(0, 16))

        tk.Button(
            action_bar,
            text="Close",
            command=self.destroy,
            font=("Segoe UI", 10),
            bg="#d0d4de",
            fg="#22313f",
            relief="flat",
            padx=12,
            pady=6,
        ).pack(side="left")

        tk.Button(
            action_bar,
            text="Apply Selected Updates",
            command=self._open_confirmation,
            font=("Segoe UI", 10, "bold"),
            bg="#2d7a4a",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=7,
        ).pack(side="right")

    def _build_club_tab(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg="#dbe5ee")
        header.pack(fill="x")
        self._make_header_label(header, "Tick", 0, 8)
        self._make_header_label(header, "Current Club", 1, 24)
        self._make_header_label(header, "Proposed Club", 2, 24)
        self._make_header_label(header, "Message", 3, 48)

        rows_frame = self._build_scroll_frame(parent)
        for candidate in self._club_candidates:
            self._add_club_row(rows_frame, candidate)

    def _build_name_tab(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg="#dbe5ee")
        header.pack(fill="x")
        self._make_header_label(header, "Tick", 0, 8)
        self._make_header_label(header, "Current Name", 1, 24)
        self._make_header_label(header, "Proposed Name", 2, 28)
        self._make_header_label(header, "Message", 3, 44)

        rows_frame = self._build_scroll_frame(parent)
        for candidate in self._name_candidates:
            self._add_name_row(rows_frame, candidate)

    def _build_scroll_frame(self, parent: tk.Misc) -> tk.Frame:
        outer = tk.Frame(parent, bg=_DEFAULT_BG)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg=_DEFAULT_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        rows_frame = tk.Frame(canvas, bg=_DEFAULT_BG)
        rows_frame.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return rows_frame

    def _make_header_label(self, parent: tk.Misc, text: str, column: int, width: int) -> None:
        label = tk.Label(
            parent,
            text=text,
            font=("Segoe UI", 10, "bold"),
            bg="#dbe5ee",
            fg="#22313f",
            anchor="w",
            width=width,
            padx=6,
            pady=6,
        )
        label.grid(row=0, column=column, sticky="ew")

    def _add_club_row(self, parent: tk.Misc, candidate: ClubMatchCandidate) -> None:
        var = tk.BooleanVar(value=False)
        self._club_vars[candidate.current_club] = var
        frame = tk.Frame(parent, bg=_DEFAULT_BG, bd=1, relief="flat")
        frame.pack(fill="x", pady=1)
        self._club_row_frames[candidate.current_club] = frame

        tk.Checkbutton(frame, variable=var, bg=_DEFAULT_BG, activebackground=_DEFAULT_BG, highlightthickness=0).grid(row=0, column=0, sticky="w", padx=(6, 0), pady=6)
        self._make_value_label(frame, candidate.current_club, 1, 24)
        self._make_value_label(frame, candidate.proposed_club, 2, 24)
        self._make_value_label(frame, candidate.message, 3, 72)
        frame.grid_columnconfigure(3, weight=1)

    def _add_name_row(self, parent: tk.Misc, candidate: NameReviewCandidate) -> None:
        var = tk.BooleanVar(value=False)
        value_var = tk.StringVar(value=candidate.proposed_name)
        self._name_vars[candidate.current_name] = var
        self._name_value_vars[candidate.current_name] = value_var

        frame = tk.Frame(parent, bg=_DEFAULT_BG, bd=1, relief="flat")
        frame.pack(fill="x", pady=1)
        self._name_row_frames[candidate.current_name] = frame

        tk.Checkbutton(frame, variable=var, bg=_DEFAULT_BG, activebackground=_DEFAULT_BG, highlightthickness=0).grid(row=0, column=0, sticky="w", padx=(6, 0), pady=6)
        self._make_value_label(frame, candidate.current_name, 1, 24)
        entry = tk.Entry(frame, textvariable=value_var, font=("Segoe UI", 10), bg="#ffffff", fg="#22313f", relief="solid", bd=1)
        entry.grid(row=0, column=2, sticky="ew", padx=6, pady=6)
        self._make_value_label(frame, candidate.message, 3, 64)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_columnconfigure(3, weight=1)

    def _make_value_label(self, parent: tk.Misc, text: str, column: int, width: int) -> None:
        label = tk.Label(
            parent,
            text=text,
            font=("Segoe UI", 10),
            bg=_DEFAULT_BG,
            fg="#22313f",
            anchor="w",
            justify="left",
            width=width,
            wraplength=520 if column >= 3 else 0,
            padx=6,
            pady=6,
        )
        label.grid(row=0, column=column, sticky="ew")

    def _open_confirmation(self) -> None:
        selected_clubs = [candidate for candidate in self._club_candidates if self._club_vars[candidate.current_club].get()]
        selected_names = [
            {
                "current_name": candidate.current_name,
                "proposed_name": self._name_value_vars[candidate.current_name].get().strip(),
                "message": candidate.message,
            }
            for candidate in self._name_candidates
            if self._name_vars[candidate.current_name].get()
        ]

        if not selected_clubs and not selected_names:
            messagebox.showwarning(
                "No updates selected",
                "Tick at least one club or name correction before continuing.",
                parent=self,
            )
            return

        club_summary = None
        if selected_clubs and self._clubs_path is not None:
            club_summary = _summarise_selection(selected_clubs, _read_club_lookup_state(self._clubs_path))

        name_summary = None
        if selected_names and self._names_path is not None:
            name_summary = _summarise_name_selection(selected_names, read_name_lookup_state(self._names_path))

        ManualReviewConfirmationDialog(self, club_summary, name_summary)

    def highlight_club_conflicts(self, conflict_clubs: List[str]) -> None:
        conflicts = set(conflict_clubs)
        for current_club, frame in self._club_row_frames.items():
            bg = _HIGHLIGHT_BG if current_club in conflicts else _DEFAULT_BG
            _set_widget_tree_bg(frame, bg)

    def highlight_name_conflicts(self, conflict_names: List[str]) -> None:
        conflicts = set(conflict_names)
        for current_name, frame in self._name_row_frames.items():
            bg = _HIGHLIGHT_BG if current_name in conflicts else _DEFAULT_BG
            _set_widget_tree_bg(frame, bg)


class ManualReviewConfirmationDialog(tk.Toplevel):
    def __init__(self, approval_dialog: ManualReviewDialog, club_summary: dict | None, name_summary: dict | None) -> None:
        super().__init__(approval_dialog)
        self.title("Confirm Manual Review Updates")
        self.geometry("860x520")
        self.configure(bg=_DEFAULT_BG)
        self.transient(approval_dialog)
        self._approval_dialog = approval_dialog
        self._club_summary = club_summary
        self._name_summary = name_summary

        self._build_ui()
        self.grab_set()

    def _build_ui(self) -> None:
        tk.Label(
            self,
            text="Confirm Manual Review Updates",
            font=("Segoe UI", 15, "bold"),
            bg=_DEFAULT_BG,
            fg="#22313f",
        ).pack(anchor="w", padx=16, pady=(14, 8))

        text = tk.Text(self, wrap="word", font=("Segoe UI", 10), bg="#ffffff", fg="#22313f")
        text.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        text.insert("1.0", _build_confirmation_text(self._club_summary, self._name_summary))
        text.config(state="disabled")

        action_bar = tk.Frame(self, bg=_DEFAULT_BG)
        action_bar.pack(fill="x", padx=16, pady=(0, 16))

        tk.Button(
            action_bar,
            text="Back To Selection",
            command=self._back_to_selection,
            font=("Segoe UI", 10),
            bg="#d0d4de",
            fg="#22313f",
            relief="flat",
            padx=12,
            pady=6,
        ).pack(side="left")

        has_conflicts = bool((self._club_summary and self._club_summary["conflicts"]) or (self._name_summary and self._name_summary["conflicts"]))
        tk.Button(
            action_bar,
            text="Confirm Write",
            command=self._confirm_write,
            state="disabled" if has_conflicts else "normal",
            font=("Segoe UI", 10, "bold"),
            bg="#2d7a4a",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=7,
        ).pack(side="right")

    def _back_to_selection(self) -> None:
        if self._club_summary:
            self._approval_dialog.highlight_club_conflicts([item.current_club for item in self._club_summary["conflicts"]])
        if self._name_summary:
            self._approval_dialog.highlight_name_conflicts([item["current_name"] for item in self._name_summary["conflicts"]])
        self.destroy()

    def _confirm_write(self) -> None:
        messages = []
        if self._club_summary and self._approval_dialog._clubs_path is not None:
            club_result = _append_club_conversions(self._approval_dialog._clubs_path, self._club_summary["new_rows"])
            messages.append(_build_write_result_text(club_result))
        if self._name_summary and self._approval_dialog._names_path is not None:
            name_result = append_name_corrections(self._approval_dialog._names_path, self._name_summary["new_rows"])
            messages.append(_build_name_write_result_text(name_result))

        self.destroy()
        self._approval_dialog.destroy()
        messagebox.showinfo("Manual Review Updated", "\n\n".join(messages), parent=self.master)


def _build_club_candidates(summary_df: pd.DataFrame | None) -> List[ClubMatchCandidate]:
    if summary_df is None or summary_df.empty:
        return []
    from .club_match_dialog import _build_candidates
    return _build_candidates(summary_df)


def _build_name_candidates(summary_df: pd.DataFrame | None) -> List[NameReviewCandidate]:
    if summary_df is None or summary_df.empty:
        return []

    candidates: List[NameReviewCandidate] = []
    df = summary_df.fillna("")
    for _, row in df.iterrows():
        current_name = str(row.get("Raw Name", "")).strip()
        proposed_name = str(row.get("Suggested Name", "")).strip()
        message = str(row.get("Message", "")).strip()
        confidence = str(row.get("Confidence", "")).strip()
        occurrences = str(row.get("Occurrences", "")).strip()
        if not current_name:
            continue
        if not message:
            parts = []
            if confidence:
                parts.append(f"Confidence {confidence}%")
            if occurrences:
                parts.append(f"Occurrences {occurrences}")
            message = " | ".join(parts)
        candidates.append(
            NameReviewCandidate(
                current_name=current_name,
                proposed_name=proposed_name,
                message=message,
                confidence=confidence,
                occurrences=occurrences,
            )
        )
    return candidates


def _summarise_name_selection(selected: List[dict], lookup_state: dict) -> dict:
    exact_existing: List[dict] = []
    conflicts: List[dict] = []
    new_rows: List[dict] = []
    alias_to_preferred = lookup_state["alias_to_preferred"]

    for item in selected:
        current_name = item["current_name"]
        proposed_name = item["proposed_name"]
        if not proposed_name:
            conflicts.append(item)
            continue
        existing = alias_to_preferred.get(current_name.lower())
        if existing:
            if existing == proposed_name:
                exact_existing.append(item)
            else:
                conflicts.append(item)
            continue
        new_rows.append(item)

    return {
        "selected": selected,
        "exact_existing": exact_existing,
        "conflicts": conflicts,
        "new_rows": new_rows,
    }


def _build_confirmation_text(club_summary: dict | None, name_summary: dict | None) -> str:
    lines: List[str] = []
    if club_summary:
        lines.append("Club updates")
        lines.append(f"Selected conversions: {len(club_summary['selected'])}")
        lines.append(f"New conversions ready to write: {len(club_summary['new_rows'])}")
        lines.append(f"Exact existing matches: {len(club_summary['exact_existing'])}")
        lines.append(f"Conflicting existing matches: {len(club_summary['conflicts'])}")
        lines.append("")
        if club_summary["new_rows"]:
            lines.append("Ready to write:")
            for item in club_summary["new_rows"]:
                lines.append(f"- {item.current_club} -> {item.proposed_club}")
            lines.append("")
        if club_summary["conflicts"]:
            lines.append("Club conflicts:")
            for item in club_summary["conflicts"]:
                lines.append(f"- {item.current_club} conflicts with proposed {item.proposed_club}")
            lines.append("")

    if name_summary:
        lines.append("Name updates")
        lines.append(f"Selected corrections: {len(name_summary['selected'])}")
        lines.append(f"New corrections ready to write: {len(name_summary['new_rows'])}")
        lines.append(f"Exact existing matches: {len(name_summary['exact_existing'])}")
        lines.append(f"Conflicting or incomplete rows: {len(name_summary['conflicts'])}")
        lines.append("")
        if name_summary["new_rows"]:
            lines.append("Ready to write:")
            for item in name_summary["new_rows"]:
                lines.append(f"- {item['current_name']} -> {item['proposed_name']}")
            lines.append("")
        if name_summary["conflicts"]:
            lines.append("Name conflicts or incomplete rows:")
            for item in name_summary["conflicts"]:
                target = item.get("proposed_name", "") or "<blank>"
                lines.append(f"- {item['current_name']} -> {target}")
            lines.append("")

    if (club_summary and club_summary["conflicts"]) or (name_summary and name_summary["conflicts"]):
        lines.append("Return to selection and untick or fix conflicting rows before writing.")
    else:
        lines.append("Confirm Write will append only new lookup rows. Exact existing matches will be skipped.")
    return "\n".join(lines)


def _build_name_write_result_text(result: dict) -> str:
    lines = [f"Name corrections updated: {result['path']}", ""]
    lines.append(f"Corrections written: {result['written']}")
    lines.append(f"Exact existing matches skipped: {result['skipped_existing']}")
    lines.append(f"Conflicting selections skipped: {result['skipped_conflicts']}")
    lines.append("")
    lines.append("Re-run the audited workbook to refresh names against the updated lookup.")
    return "\n".join(lines)