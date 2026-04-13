"""
graphical — Qt-backed graphical interface for WRRL Admin Suite.
"""

from .qt import launch_dashboard, launch_qt_dashboard

launch = launch_dashboard

__all__ = ["launch", "launch_dashboard", "launch_qt_dashboard"]
