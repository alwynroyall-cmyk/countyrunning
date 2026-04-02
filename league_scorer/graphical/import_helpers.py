"""Shared UI helpers for Race Roster and Sporthive imports."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog

from ..raceroster_import import import_sporthive_manual_pages


@dataclass
class RaceImportRequest:
    race_url: str
    race_number: int
    race_name: str | None
    sporthive_race_hint: int | None


def prompt_race_import_request(parent: tk.Misc) -> RaceImportRequest | None:
    """Prompt for common race-import fields used by both scorer and dashboard UIs."""
    race_url = simpledialog.askstring(
        "Import From Race Roster",
        "Paste the Race Roster race URL:",
        parent=parent,
    )
    if not race_url:
        return None

    race_number = simpledialog.askinteger(
        "League Race Number",
        "Enter the league race number for this file (for example 4):",
        parent=parent,
        minvalue=1,
        maxvalue=99,
    )
    if race_number is None:
        return None

    race_name = simpledialog.askstring(
        "Race Name",
        "Optional race name for the file title (for example Broad Town 5):",
        parent=parent,
    )

    sporthive_race_hint = None
    if is_sporthive_event_summary_url(race_url):
        sporthive_race_hint = simpledialog.askinteger(
            "Sporthive Race ID",
            "This Sporthive link is an event summary. Enter the Race ID from the 'View results' URL (the number after /race/).",
            parent=parent,
            minvalue=1,
        )
        if sporthive_race_hint is None:
            return None

    return RaceImportRequest(
        race_url=race_url,
        race_number=race_number,
        race_name=race_name,
        sporthive_race_hint=sporthive_race_hint,
    )


def is_sporthive_event_summary_url(race_url: str) -> bool:
    lower_url = race_url.lower()
    return "sporthive.com/events/s/" in lower_url and "/race/" not in lower_url


def ask_multiline_page_text(
    parent: tk.Misc,
    title: str,
    prompt: str,
    *,
    bg: str,
    fg: str,
    panel_bg: str,
    accent_bg: str,
    accent_fg: str,
) -> str | None:
    """Show a reusable multiline paste dialog used for manual Sporthive paging."""
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.transient(parent.winfo_toplevel())
    dialog.grab_set()
    dialog.geometry("760x520")
    dialog.configure(bg=bg)

    tk.Label(dialog, text=prompt, bg=bg, fg=fg, justify="left", anchor="w").pack(
        fill="x", padx=12, pady=(12, 6)
    )

    text_box = scrolledtext.ScrolledText(dialog, wrap="word", font=("Consolas", 10), height=22)
    text_box.pack(fill="both", expand=True, padx=12, pady=6)

    result = {"value": None}
    button_row = tk.Frame(dialog, bg=bg)
    button_row.pack(fill="x", padx=12, pady=(0, 12))

    def _ok() -> None:
        result["value"] = text_box.get("1.0", "end").strip()
        dialog.destroy()

    def _cancel() -> None:
        result["value"] = None
        dialog.destroy()

    tk.Button(
        button_row,
        text="Cancel",
        command=_cancel,
        bg=panel_bg,
        fg=fg,
        relief="flat",
        padx=10,
        pady=4,
    ).pack(side="right")
    tk.Button(
        button_row,
        text="Use This Page",
        command=_ok,
        bg=accent_bg,
        fg=accent_fg,
        relief="flat",
        padx=12,
        pady=4,
    ).pack(side="right", padx=(0, 8))

    dialog.wait_window()
    return result["value"]


def run_manual_sporthive_import(
    parent: tk.Misc,
    *,
    input_dir: Path,
    request: RaceImportRequest,
    ask_page_text,
) -> tuple[Path, int, Path] | None:
    """Drive manual page capture and import for Sporthive races."""
    pages: list[str] = []
    page_no = 1

    while True:
        text = ask_page_text(page_no)
        if text is None:
            return None
        if not text.strip():
            messagebox.showwarning(
                "No content",
                "No rows detected. Paste the page content and try again.",
                parent=parent,
            )
            continue
        pages.append(text)

        more = messagebox.askyesno(
            "Add Another Page?",
            "Add another Sporthive results page?\n\nChoose No when all pages are pasted.",
            parent=parent,
        )
        if not more:
            break
        page_no += 1

    try:
        return import_sporthive_manual_pages(
            race_url=request.race_url,
            pages_text=pages,
            input_dir=input_dir,
            league_race_number=request.race_number,
            race_name_override=request.race_name,
        )
    except Exception as exc:
        messagebox.showerror("Sporthive manual import failed", str(exc), parent=parent)
        return None
