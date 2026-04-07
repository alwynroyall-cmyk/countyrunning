WRRL League AI
===============

Release: 7.0.0

Overview
--------
WRRL League AI is a desktop scoring and audit workflow for Wiltshire Road Running League seasons.
It supports race ingestion, audit cleansing, issue review, scoring, staged checks, and report output.

Key Changes In 7.0.0
--------------------
- Structured season input folders under inputs/:
	- raw_data
	- series
	- control
	- audited
	- raw_data_archive
- Structured season output folders under outputs/ (publish/audit/quality/autopilot).
- Autopilot now rebuilds audited files from raw_data and enforces provenance.
- Autopilot now archives raw files to raw_data_archive (write-once by filename).
- Settings panel includes Set Up New Season.
- Autopilot completion dialog is more user-friendly and includes Review Messages.

Run
---
- GUI: python run_gui.py
- Module entry: python -m league_scorer

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
- User settings are saved in:
	- ~/.wrrl_prefs.json
	- ~/.wrrl_settings.json
