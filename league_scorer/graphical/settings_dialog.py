import getpass
import platform
import socket
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk, messagebox
from ..settings import settings, DEFAULT_SETTINGS
from ..graphical.dashboard import WRRL_NAVY, WRRL_LIGHT


SETTING_METADATA = {
    "BEST_N": {
        "title": "Highest number of races to score (Best n from League)",
        "description": "This applies to Club League and Individual League.",
    },
    "MAX_RACES": {
        "title": "Maximum number of races (events) in a league season",
        "description": "Used for league calculations.",
    },
    "TEAM_SIZE": {
        "title": "Team Size",
        "description": "For all clubs this is the number of Male or Female runners for each Team (A and B), so 5 would indicate first 5 in team A, the second 5 in team B. Male/Female counted separately.",
    },
    "MAX_DIV_PTS": {
        "title": "Max Div Pts",
        "description": "The number of points for the top scoring team in a race, the next team receive one point less, and so on.",
    },
    "SEASON_FINAL_RACE": {
        "title": "Season Final Race",
        "description": "The Race Number of the final race in the season.",
    },
}


class SettingsPanel(tk.Frame):
    def __init__(self, parent, on_close=None, on_open_logs=None):
        super().__init__(parent, bg=WRRL_LIGHT)
        self._vars = {}
        self._on_close = on_close
        self._on_open_logs = on_open_logs
        self._project_root = Path(__file__).resolve().parents[2]
        self._build_ui()

    def _build_ui(self):
        frm = tk.Frame(self, bg=WRRL_LIGHT, padx=20, pady=20)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(0, weight=1, uniform="panel")
        frm.grid_columnconfigure(1, weight=1, uniform="panel")
        row = 0
        title = tk.Label(frm, text="League Settings", font=("Segoe UI", 16, "bold"), bg=WRRL_LIGHT, fg=WRRL_NAVY)
        title.grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 16))
        row += 1
        settings_frame = tk.Frame(frm, bg=WRRL_LIGHT, padx=8, pady=8)
        settings_frame.grid(row=row, column=0, sticky="nsew", padx=(0, 12))
        settings_frame.grid_columnconfigure(1, weight=1)

        info_frame = tk.Frame(frm, bg="#F4F6F8", padx=18, pady=18, highlightbackground="#D4DCE4", highlightthickness=1)
        info_frame.grid(row=row, column=1, sticky="nsew", padx=(12, 0))
        info_frame.grid_columnconfigure(0, weight=1)

        settings_row = 0
        for key, default in DEFAULT_SETTINGS.items():
            metadata = SETTING_METADATA.get(
                key,
                {
                    "title": key.replace("_", " ").title(),
                    "description": "",
                },
            )
            var = tk.IntVar(value=settings.get(key))
            entry = ttk.Entry(settings_frame, textvariable=var, width=8, font=("Segoe UI", 11))
            entry.grid(row=settings_row, column=0, rowspan=2, sticky="nw", padx=(0, 12), pady=(6, 8))
            label = tk.Label(
                settings_frame,
                text=metadata["title"],
                bg=WRRL_LIGHT,
                fg=WRRL_NAVY,
                font=("Segoe UI", 11, "bold"),
                wraplength=340,
                justify="left",
            )
            label.grid(row=settings_row, column=1, sticky="w", pady=(6, 2))
            self._vars[key] = var
            settings_row += 1
            if metadata["description"]:
                description = tk.Label(
                    settings_frame,
                    text=metadata["description"],
                    bg=WRRL_LIGHT,
                    fg=WRRL_NAVY,
                    font=("Segoe UI", 9),
                    wraplength=340,
                    justify="left",
                )
                description.grid(row=settings_row, column=1, sticky="w", pady=(0, 8))
            settings_row += 1

        self._build_info_panel(info_frame)

        row += 1
        btn_frame = tk.Frame(frm, bg=WRRL_LIGHT)
        btn_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(18,0))
        btn_frame.grid_columnconfigure(1, weight=1)
        save_btn = ttk.Button(btn_frame, text="Save", command=self._on_save)
        save_btn.grid(row=0, column=0, sticky="w", padx=(0, 10))
        logs_btn = ttk.Button(btn_frame, text="View Logs", command=self._on_open_logs_clicked)
        logs_btn.grid(row=0, column=1, sticky="w")
        close_btn = tk.Button(
            btn_frame,
            text="\u25c4 Dashboard",
            command=self._on_close_clicked,
            font=("Segoe UI", 10, "bold"),
            bg="#2d7a4a",
            fg="#ffffff",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground="#ffffff",
        )
        close_btn.grid(row=0, column=2, sticky="e")

    def _build_info_panel(self, parent):
        title = tk.Label(
            parent,
            text="System Information",
            bg="#F4F6F8",
            fg=WRRL_NAVY,
            font=("Segoe UI", 13, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = tk.Label(
            parent,
            text="Useful runtime and environment details for this installation.",
            bg="#F4F6F8",
            fg=WRRL_NAVY,
            font=("Segoe UI", 9),
            wraplength=360,
            justify="left",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(4, 14))

        for row, (label_text, value_text) in enumerate(self._get_info_rows(), start=2):
            item = tk.Frame(parent, bg="#F4F6F8")
            item.grid(row=row, column=0, sticky="ew", pady=(0, 10))
            item.grid_columnconfigure(0, weight=1)

            label = tk.Label(
                item,
                text=label_text,
                bg="#F4F6F8",
                fg=WRRL_NAVY,
                font=("Segoe UI", 9, "bold"),
                anchor="w",
            )
            label.grid(row=0, column=0, sticky="w")

            value = tk.Label(
                item,
                text=value_text,
                bg="#F4F6F8",
                fg=WRRL_NAVY,
                font=("Segoe UI", 9),
                wraplength=360,
                justify="left",
                anchor="w",
            )
            value.grid(row=1, column=0, sticky="w")

    def _get_info_rows(self):
        latest_source_change = self._get_latest_source_change()
        settings_path = Path.home() / ".wrrl_settings.json"
        return [
            ("User", getpass.getuser()),
            ("Machine", socket.gethostname()),
            ("Operating System", f"{platform.system()} {platform.release()}"),
            ("Python", platform.python_version()),
            ("Project Folder", str(self._project_root)),
            ("Settings File", str(settings_path)),
            ("Settings File Updated", self._format_modified_time(settings_path)),
            ("Latest Source Change", latest_source_change),
            ("Current Time", datetime.now().strftime("%d %b %Y %H:%M:%S")),
        ]

    def _get_latest_source_change(self):
        latest_path = None
        latest_mtime = None
        for path in self._project_root.rglob("*.py"):
            try:
                modified_time = path.stat().st_mtime
            except OSError:
                continue
            if latest_mtime is None or modified_time > latest_mtime:
                latest_mtime = modified_time
                latest_path = path

        if latest_path is None or latest_mtime is None:
            return "Unavailable"

        relative_path = latest_path.relative_to(self._project_root)
        timestamp = datetime.fromtimestamp(latest_mtime).strftime("%d %b %Y %H:%M")
        return f"{relative_path} at {timestamp}"

    def _format_modified_time(self, path):
        if not path.exists():
            return "Not created yet"
        try:
            return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d %b %Y %H:%M")
        except OSError:
            return "Unavailable"

    def _on_save(self):
        for key, var in self._vars.items():
            try:
                value = int(var.get())
                settings.set(key, value)
            except Exception:
                messagebox.showerror("Invalid Value", f"{key} must be an integer.")
                return
        if self._on_close:
            self._on_close()

    def _on_close_clicked(self):
        if self._on_close:
            self._on_close()

    def _on_open_logs_clicked(self):
        if self._on_open_logs:
            self._on_open_logs()
