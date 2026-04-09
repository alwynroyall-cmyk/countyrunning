# WRRL League AI 8.0.1 Release Notes

## Summary

This release updates the club reporting workflow, simplifies published output layout, and aligns the documentation with the current season workflow.

## Key changes

- Added club report generation as a DOCX output for all clubs.
- Updated club report headers to show `WRRL | Season Summary {year}` plus the club name.
- Simplified published output structure to:
  - `outputs/publish/docx/`
  - `outputs/publish/pdf/`
  - `outputs/publish/standings/`
  - `outputs/publish/review-packs/`
- Added automatic migration from legacy `outputs/publish/xlsx/` folders into the new publish layout.
- Improved GUI publish flow to run club reports after final publish and remove the redundant completion dialog.
- Updated `ReadMe.txt` to reflect the current release and output folder structure.
