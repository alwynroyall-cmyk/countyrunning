import getpass
import platform
import socket
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk, messagebox, filedialog, simpledialog
from ..settings import settings, DEFAULT_SETTINGS
from ..session_config import config as session_config
from ..input_layout import sort_existing_input_files
from ..output_layout import sort_existing_output_files
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
        self._path_value_labels = {}
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
        paths_frame = tk.Frame(frm, bg=WRRL_LIGHT)
        paths_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        self._build_paths_panel(paths_frame)

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

    def _build_paths_panel(self, parent):
        """Build the data paths configuration section."""
        # Title
        title = tk.Label(
            parent,
            text="Data Paths",
            font=("Segoe UI", 12, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        )
        title.pack(anchor="w", pady=(0, 12))

        # Data Root selector frame
        root_frame = tk.Frame(parent, bg=WRRL_LIGHT)
        root_frame.pack(fill="x", pady=(0, 12))

        root_label = tk.Label(
            root_frame,
            text="Data Root:",
            font=("Segoe UI", 10, "bold"),
            bg=WRRL_LIGHT,
            fg=WRRL_NAVY,
        )
        root_label.pack(side="left")

        self._data_root_display = tk.Label(
            root_frame,
            text=str(session_config.data_root) if session_config.data_root else "Not set",
            font=("Segoe UI", 9),
            bg=WRRL_LIGHT,
            fg="#ff9900" if not session_config.data_root else "#2d7a4a",
            anchor="w",
        )
        self._data_root_display.pack(side="left", fill="x", expand=True, padx=(8, 10))

        browse_btn = tk.Button(
            root_frame,
            text="Browse…",
            command=self._on_browse_data_root,
            font=("Segoe UI", 9),
            bg="#2d7a4a",
            fg="white",
            relief="flat",
            padx=10,
            pady=2,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground="white",
        )
        browse_btn.pack(side="right")

        setup_btn = tk.Button(
            root_frame,
            text="Set Up New Season...",
            command=self._on_setup_new_season,
            font=("Segoe UI", 9),
            bg="#2d7a4a",
            fg="white",
            relief="flat",
            padx=10,
            pady=2,
            cursor="hand2",
            activebackground="#1f5632",
            activeforeground="white",
        )
        setup_btn.pack(side="right", padx=(0, 8))

        # Path info frame
        info_frame = tk.Frame(parent, bg="#f5f5f5", padx=10, pady=10, highlightbackground="#d0d0d0", highlightthickness=1)
        info_frame.pack(fill="x")

        paths_title = tk.Label(
            info_frame,
            text="Paths for current season",
            font=("Segoe UI", 9, "bold"),
            bg="#f5f5f5",
            fg=WRRL_NAVY,
        )
        paths_title.pack(anchor="w", pady=(0, 8))

        for label_text, config_attr in [
            ("Input Folder:", "input_dir"),
            ("Output Folder:", "output_dir"),
            ("Events File:", "events_path"),
        ]:
            path_row = tk.Frame(info_frame, bg="#f5f5f5")
            path_row.pack(fill="x", pady=(0, 6))

            lbl = tk.Label(
                path_row,
                text=label_text,
                font=("Segoe UI", 9, "bold"),
                bg="#f5f5f5",
                fg=WRRL_NAVY,
                width=15,
                anchor="w",
            )
            lbl.pack(side="left")

            path_lbl = tk.Label(
                path_row,
                text="",
                font=("Segoe UI", 9),
                bg="#f5f5f5",
                fg="#b0b0b0",
                anchor="w",
            )
            path_lbl.pack(side="left", fill="x", expand=True)
            self._path_value_labels[config_attr] = path_lbl

        self._refresh_paths_display()

    def _refresh_paths_display(self):
        for config_attr, path_lbl in self._path_value_labels.items():
            path_value = getattr(session_config, config_attr, None)
            if path_value:
                display_text = path_value.name if config_attr == "events_path" else str(path_value)
                exists = path_value.exists() if hasattr(path_value, "exists") else False
                text_color = "#2d7a4a" if exists else "#ff9900"
            else:
                display_text = "—"
                text_color = "#b0b0b0"
            path_lbl.config(text=display_text, fg=text_color)

    def _on_browse_data_root(self):
        """Open a folder browser to select the data root."""
        initial = str(session_config.data_root) if session_config.data_root else "/"
        folder = filedialog.askdirectory(
            title="Select Data Root Folder (parent of all year folders)",
            initialdir=initial,
        )
        if folder:
            session_config.data_root = Path(folder)
            session_config.ensure_dirs()
            self._data_root_display.config(
                text=str(session_config.data_root),
                fg="#2d7a4a",
            )
            self._refresh_paths_display()
            messagebox.showinfo(
                "Data Root Updated",
                f"Data root set to:\n{session_config.data_root}",
            )

    def _on_setup_new_season(self):
        if not session_config.data_root:
            messagebox.showerror(
                "Data Root Missing",
                "Set Data Root first, then set up a season.",
            )
            return

        year = simpledialog.askinteger(
            "Set Up New Season",
            "Enter season year:",
            parent=self,
            initialvalue=session_config.year,
            minvalue=2020,
            maxvalue=2100,
        )
        if year is None:
            return

        session_config.year = year
        session_config.ensure_dirs()
        if session_config.input_dir:
            sort_existing_input_files(session_config.input_dir)
        if session_config.output_dir:
            sort_existing_output_files(session_config.output_dir)

        self._data_root_display.config(
            text=str(session_config.data_root),
            fg="#2d7a4a",
        )
        self._refresh_paths_display()
        messagebox.showinfo(
            "Season Ready",
            f"Season {year} is ready.\n\nInput: {session_config.input_dir}\nOutput: {session_config.output_dir}",
        )

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
