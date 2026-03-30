"""
graphical — GUI and graphical interfaces for the Wiltshire League Management.
"""

from .dashboard import launch_dashboard
from .events_viewer import EventsViewerWindow
from .gui import launch

__all__ = ["launch", "launch_dashboard", "EventsViewerWindow"]
