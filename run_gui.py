"""
run_gui.py
----------
IDLE-friendly entry point.  Open this file in IDLE and press F5.

Expected layout:
    your_project/
    ├── run_gui.py          <-- this file
    └── league_scorer/      <-- your package folder
        ├── __init__.py
        ├── main.py
        ├── gui.py
        └── ...
"""

import sys
from pathlib import Path

# Make sure the package folder is discoverable regardless of how IDLE
# sets up sys.path when running a top-level script.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from league_scorer.graphical import launch_dashboard

if __name__ == "__main__":
    try:
        launch_dashboard()
    except KeyboardInterrupt:
        # Allow Ctrl+C from the terminal to exit cleanly without a traceback
        print("Interrupted; exiting.")
        sys.exit(0)
