# Dashboard UI: Run Autopilot Dirty State Visuals

- Added an amber state to the `Run Autopilot` action card when RAES or autopilot writes a `dirty` flag.
- The header run button and the small header status dot are now consistent with the RAES "DATA DIRTY" message.
- When dirty, the `Run Autopilot` card now pulses between two amber tones to draw attention.

Why: Make the prominent action on the dashboard clearly reflect data freshness and RAES edits pending.

Files changed:

- `league_scorer/graphical/dashboard.py` — added pulsing behaviour, amber colours, and dynamic updates.
- `league_scorer/graphical/gui.py` — header run button now turns amber when dirty.

