WRRL League AI
===============

Release: 8.0.1

Overview
--------
WRRL League AI is a desktop scoring and audit workflow for Wiltshire Road Running League seasons.
It supports race ingestion, audit cleansing, issue review, scoring, staged checks, and report output.

Key Changes In 8.0.1
--------------------
- Club report generation is now supported as DOCX output for all clubs.
- Simplified publish output layout:
	- outputs/publish/docx/
	- outputs/publish/pdf/
	- outputs/publish/standings/
	- outputs/publish/review-packs/
- Audit and quality execution artifacts remain under outputs/audit/ and outputs/quality/.
- Legacy publish/xlsx layout is now migrated into the new publish structure.
- Autopilot now uses a cleaner completion flow when club reports are generated after final publish.

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
