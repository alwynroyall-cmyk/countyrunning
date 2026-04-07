from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import pandas as pd
import tkinter as tk
from tkinter import ttk

from ..session_config import config as session_config
from .results_workbook import find_latest_results_workbook, sorted_race_sheet_names

WRRL_NAVY = "#3a4658"
WRRL_GREEN = "#2d7a4a"
WRRL_LIGHT = "#f5f5f5"
WRRL_WHITE = "#ffffff"

_HIGHLIGHT_BG = "#fff1c7"
_DEFAULT_BG = WRRL_LIGHT


# -----------------------------
# Text / formatting helpers
# -----------------------------

def proper_case(s: str) -> str:
    import re
    return re.sub(
        r"([A-Za-z])([A-Za-z']*)",
        lambda m: m.group(1).upper() + m.group(2).lower(),
        s,
    )


# -----------------------------
# UI helpers
# -----------------------------

def make_row_frame(parent: tk.Misc) -> tk.Frame:
    return tk.Frame(parent, bg=_DEFAULT_BG, bd=1, relief="flat")


def make_header_label(parent: tk.Misc, text: str, column: int, width: int) -> None:
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


def make_value_label(
    parent: tk.Misc,
    text: str,
    column: int,
    width: int,
    wrap: bool = False,
) -> tk.Label:
    label = tk.Label(
        parent,
        text=text,
        font=("Segoe UI", 10),
        bg=_DEFAULT_BG,
        fg="#22313f",
        anchor="w",
        justify="left",
        width=width,
        wraplength=520 if wrap else 0,
        padx=6,
        pady=6,
    )
    label.grid(row=0, column=column, sticky="ew")
    return label


def build_scroll_frame(parent: tk.Misc) -> tk.Frame:
    outer = tk.Frame(parent, bg=_DEFAULT_BG)
    outer.pack(fill="both", expand=True)
    canvas = tk.Canvas(outer, bg=_DEFAULT_BG, highlightthickness=0)
    scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    rows_frame = tk.Frame(canvas, bg=_DEFAULT_BG)
    rows_frame.bind(
        "<Configure>",
        lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=rows_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    return rows_frame


def set_widget_tree_bg(widget: tk.Widget, bg: str) -> None:
    try:
        widget.configure(bg=bg)
    except Exception:
        pass
    for child in widget.winfo_children():
        set_widget_tree_bg(child, bg)


# -----------------------------
# Anomaly scanning helpers
# -----------------------------

def scan_workbook_for_runner_state() -> Dict[str, Dict[str, set]]:
    if session_config.output_dir is None:
        return {}

    workbook = find_latest_results_workbook(session_config.output_dir)
    if workbook is None:
        return {}

    runner_state: Dict[str, Dict[str, set]] = {}
    try:
        xl = pd.ExcelFile(workbook)
        for sheet in sorted_race_sheet_names(xl):
            df = xl.parse(sheet).fillna("")
            if "Name" not in df.columns:
                continue
            for _, row in df.iterrows():
                name = str(row.get("Name", "")).strip()
                if not name:
                    continue
                norm = name.lower()
                state = runner_state.setdefault(
                    norm,
                    {
                        "club": set(),
                        "gender": set(),
                        "category": set(),
                        "raw_names": set(),
                    },
                )
                state["raw_names"].add(name)

                club = str(row.get("Club", "")).strip()
                if club:
                    state["club"].add(club)

                gender = str(row.get("Gender", "")).strip().upper()
                if gender:
                    state["gender"].add(gender)

                category = str(row.get("Category", "")).strip()
                if category:
                    state["category"].add(category)
    except Exception:
        return {}

    return runner_state


def detect_runner_anomalies(
    runner_state: Dict[str, Dict[str, set]]
) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for norm, state in runner_state.items():
        flags: List[str] = []
        detail_parts: List[str] = []

        if len(state["club"]) > 1:
            flags.append("Club")
            detail_parts.append("clubs=" + ", ".join(sorted(state["club"])))

        if len(state["gender"]) > 1:
            flags.append("Gender")
            detail_parts.append("gender=" + ", ".join(sorted(state["gender"])))

        if len(state["category"]) > 1:
            flags.append("Category")
            preview = sorted(state["category"])
            detail_parts.append(
                "categories="
                + ", ".join(preview[:6])
                + ("..." if len(preview) > 6 else "")
            )

        if not flags:
            continue

        # Use any raw name, but display in Proper Case
        raw_name = next(iter(state["raw_names"]))
        rows.append(
            {
                "runner": proper_case(raw_name),
                "anomalies": "/".join(flags),
                "details": " | ".join(detail_parts),
            }
        )

    rows.sort(key=lambda item: item["runner"].lower())
    return rows


# -----------------------------
# Row builders for club/name tabs
# -----------------------------

def add_club_row(
    parent: tk.Misc,
    candidate,
    var: tk.BooleanVar,
    row_frames: Dict[str, tk.Frame],
) -> None:
    frame = make_row_frame(parent)
    frame.pack(fill="x", pady=1)
    row_frames[candidate.current_club] = frame

    tk.Checkbutton(
        frame,
        variable=var,
        bg=_DEFAULT_BG,
        activebackground=_DEFAULT_BG,
        highlightthickness=0,
    ).grid(row=0, column=0, sticky="w", padx=(6, 0), pady=6)

    make_value_label(frame, candidate.current_club, 1, 24)
    make_value_label(frame, candidate.proposed_club, 2, 24)
    make_value_label(frame, candidate.message, 3, 72, wrap=True)
    frame.grid_columnconfigure(3, weight=1)


def add_name_row(
    parent: tk.Misc,
    candidate,
    var: tk.BooleanVar,
    value_var: tk.StringVar,
    row_frames: Dict[str, tk.Frame],
) -> None:
    frame = make_row_frame(parent)
    frame.pack(fill="x", pady=1)
    row_frames[candidate.current_name] = frame

    tk.Checkbutton(
        frame,
        variable=var,
        bg=_DEFAULT_BG,
        activebackground=_DEFAULT_BG,
        highlightthickness=0,
    ).grid(row=0, column=0, sticky="w", padx=(6, 0), pady=6)

    make_value_label(frame, candidate.current_name, 1, 24)

    entry = tk.Entry(
        frame,
        textvariable=value_var,
        font=("Segoe UI", 10),
        bg="#ffffff",
        fg="#22313f",
        relief="solid",
        bd=1,
    )
    entry.grid(row=0, column=2, sticky="ew", padx=6, pady=6)

    make_value_label(frame, candidate.message, 3, 64, wrap=True)
    frame.grid_columnconfigure(2, weight=1)
    frame.grid_columnconfigure(3, weight=1)


__all__ = [
    "WRRL_NAVY",
    "WRRL_GREEN",
    "WRRL_LIGHT",
    "WRRL_WHITE",
    "_HIGHLIGHT_BG",
    "_DEFAULT_BG",
    "proper_case",
    "make_row_frame",
    "make_header_label",
    "make_value_label",
    "build_scroll_frame",
    "set_widget_tree_bg",
    "scan_workbook_for_runner_state",
    "detect_runner_anomalies",
    "add_club_row",
    "add_name_row",
]