import tkinter as tk
from tkinter import ttk, messagebox
from ..settings import settings, DEFAULT_SETTINGS
from ..graphical.dashboard import WRRL_NAVY, WRRL_LIGHT, WRRL_GREEN

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("League Settings")
        self.configure(bg=WRRL_LIGHT)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self._vars = {}
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _build_ui(self):
        frm = tk.Frame(self, bg=WRRL_LIGHT, padx=20, pady=20)
        frm.pack(fill="both", expand=True)
        row = 0
        for key, default in DEFAULT_SETTINGS.items():
            label = tk.Label(frm, text=key.replace('_', ' ').title(), bg=WRRL_LIGHT, fg=WRRL_NAVY, font=("Segoe UI", 11))
            label.grid(row=row, column=0, sticky="w", pady=6)
            var = tk.IntVar(value=settings.get(key))
            entry = ttk.Entry(frm, textvariable=var, width=8, font=("Segoe UI", 11))
            entry.grid(row=row, column=1, sticky="w", pady=6, padx=(8,0))
            self._vars[key] = var
            row += 1
        btn_frame = tk.Frame(frm, bg=WRRL_LIGHT)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(18,0))
        save_btn = ttk.Button(btn_frame, text="Save", command=self._on_save)
        save_btn.pack(side="left", padx=(0,10))
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self._on_cancel)
        cancel_btn.pack(side="left")

    def _on_save(self):
        for key, var in self._vars.items():
            try:
                value = int(var.get())
                settings.set(key, value)
            except Exception:
                messagebox.showerror("Invalid Value", f"{key} must be an integer.")
                return
        self.destroy()

    def _on_cancel(self):
        self.destroy()
