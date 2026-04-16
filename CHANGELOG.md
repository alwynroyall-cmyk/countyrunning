# Changelog

## 8.2.2

- Added RAES `Name Review` panel mode for audited runner name suggestions.
- Ensured RAES name corrections are written to the manual audit register and persisted to `name_corrections.xlsx`.
- Fixed autopilot audit refresh by clearing stale `outputs/audit/workbooks/` before rebuilding audit reports.
- Reduced Season Audit workbook output to the core review sheets and removed unused diagnostic-only sheets while retaining `Unrecognised Club Summary` for club review.
- Updated package metadata and documentation to reflect `8.2.2`.

## 8.2.0

- Updated UI consistency across Qt windows with matching button styling and refresh iconography.
- Added the Runner / Club Enquiry single-instance guard to prevent duplicate dashboard launches.
- Improved the Export Published PDFs workflow by showing an in-app progress dialog and retaining folder open behavior when complete.
- Standardized the enquiry screen toolbar and aligned it with other `View Events`, `View Results`, `View Autopilot`, and RAES panels.
- Bumped project version to `8.2.0` and updated user-facing documentation.

## 8.0.1

- Added support for club report generation as a single DOCX for all clubs.
- Reorganized publish output under `outputs/publish/docx/`, `outputs/publish/pdf/`, `outputs/publish/standings/`, and `outputs/publish/review-packs/`, with club reports isolated under `outputs/publish/docx/club-reports/`.
- Updated `league_scorer/` package layout for clearer functional separation.
- Added structured failure reporting for unexpected publish and club report failure handling.
- Added regression tests for publish failure handling and club report error reporting.
- Added `outputs/publish/package/` for grouped deployable publish artifacts, with files flattened into a single top-level package by default.
- Added `scripts/publish/package_publish.py` for packaging publish outputs and optionally copying or zipping them for deployment, with `--no-flatten` available to preserve nested publish directories.
- Improved stability of the GUI workflow and dirty-state refresh indicators.
- Added documentation notes in `ReadMe.txt` describing the new stability and report behavior.
