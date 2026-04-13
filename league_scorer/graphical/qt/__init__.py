"""Qt-backed graphical frontend for WRRL Admin Suite."""

from .dashboard import launch_dashboard

launch_qt_dashboard = launch_dashboard

__all__ = ["launch_dashboard", "launch_qt_dashboard"]
