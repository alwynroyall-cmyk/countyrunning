WRRL League AI
===============

Release: 8.2.3

Overview
--------
WRRL League AI is a desktop scoring and audit workflow for Wiltshire Road Running League seasons.
It supports race ingestion, audit cleansing, issue review, scoring, staged checks, and report output.

Key Changes In 8.2.3
--------------------
- Broadened `Actionable Issues` to capture category, club, gender, name variant, and data-invalid row issues that require manual review.
- Removed `Candidates To Check` and `EA Checked` from Season Audit workbooks to reduce audit noise.
- Kept `Unrecognised Club Summary` for club review and retained the core audit review sheets.
- Added `Name Review` support for RAES and manual audit logging through `name_corrections.xlsx`.
- Updated autopilot to clear stale audit workbook files before rebuilding reports.
- Code is organized by functional package areas within `league_scorer/` and workflow wrappers are grouped under `scripts/publish/` and `scripts/autopilot/`.

Package Layout
--------------
The core package is now organized by functional areas:
- `league_scorer/config/` — settings and session configuration
- `league_scorer/input/` — race inputs, event loading, club data, and source discovery
- `league_scorer/output/` — output paths, writers, reporting, structured logging
- `league_scorer/process/` — scoring, race processing, validation, aggregation, rules
- `league_scorer/autopilot/` — audit, issue resolution, manual review helpers, staged checks support
- `league_scorer/publish/` — final publish workflow and club report generation
- `league_scorer/views/` — dashboard-connected results, enquiry, and autopilot view panels
- `league_scorer/graphical/` — main dashboard UI and supporting widgets
- `league_scorer/raes/` — RAES manual review modules

Run
---
- GUI: python run_gui.py
- Module entry: python -m league_scorer
- Autopilot CLI wrapper: python scripts/autopilot/run_full_autopilot.py --year <year> --data-root <root>
- Publish CLI wrapper: python scripts/publish/run_publish_results.py --year <year> --data-root <root>
- Package publish wrapper: python scripts/publish/package_publish.py --year <year> --data-root <root> [--dest <folder>] [--zip <archive.zip>] [--no-flatten]

Data Layout
-----------
The configured data root should contain one folder per season year:

	{data_root}/{year}/inputs
	{data_root}/{year}/outputs

Expected files:
- inputs/control/clubs.xlsx
- inputs/control/wrrl_events.xlsx
- inputs/control/name_corrections.xlsx

Notes
-----
- PDF conversion requires Microsoft Word (via docx2pdf).
- Unexpected publish or club report failures now write structured JSON/Markdown reports under the configured report directory.
- User settings are saved in:
	- ~/.wrrl_prefs.json
	- ~/.wrrl_settings.json
