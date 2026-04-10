# Changelog

## 8.0.1

- Added support for club report generation as a single DOCX for all clubs.
- Reorganized publish output under `outputs/publish/docx/`, `outputs/publish/pdf/`, `outputs/publish/standings/`, and `outputs/publish/review-packs/`, with club reports isolated under `outputs/publish/docx/club-reports/`.
- Updated `league_scorer/` package layout for clearer functional separation.
- Added structured failure reporting for unexpected publish and club report failures.
- Added regression tests for publish failure handling and club report error reporting.
- Added `outputs/publish/package/` for grouped deployable publish artifacts, with files flattened into a single top-level package by default.
- Added `scripts/publish/package_publish.py` for packaging publish outputs and optionally copying or zipping them for deployment, with `--no-flatten` available to preserve nested publish directories.
- Improved stability of the GUI workflow and dirty-state refresh indicators.
- Added documentation notes in `ReadMe.txt` describing the new stability and report behavior.
