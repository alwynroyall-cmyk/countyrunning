import tkinter as tk
import pytest

from league_scorer.graphical import dashboard


def test_create_action_button_returns_frame():
    root = tk.Tk()
    try:
        db = dashboard
        frame = db._create_action_button(root, "Test", "subtitle", lambda: None, 0, 0)
        assert hasattr(frame, "winfo_children")
    finally:
        root.destroy()
