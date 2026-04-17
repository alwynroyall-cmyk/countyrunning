# WRRL League AI Combined Release Notes

This document consolidates all release notes from versions 8.0.1 through 8.3.0 for the WRRL League AI.

---

## 8.0.1 Release Notes

This release updates the club reporting workflow, simplifies published output layout, and aligns the documentation with the current season workflow.

### Key changes
- Added club report generation as a DOCX output for all clubs.
- Updated club report headers to show `WRRL | Season Summary {year}` plus the club name.
- Simplified published output structure to:
  - `outputs/publish/docx/`
  - `outputs/publish/docx/club-reports/`
  - `outputs/publish/pdf/`
  - `outputs/publish/standings/`
  - `outputs/publish/review-packs/`
- Added automatic migration from legacy `outputs/publish/xlsx/` folders into the new publish layout.
- Improved GUI publish flow to run club reports after final publish and remove the redundant completion dialog.
- Updated `ReadMe.txt` to reflect the current release and output folder structure.

---

## 8.2.0 Release Notes

This release focuses on GUI polish, workflow feedback, and consistency across the Qt dashboard.

### Highlights
- Updated Qt button styling and aligned refresh buttons across the enquiry, results, autopilot, RAES, and compare screens.
- Improved the Runner / Club Enquiry experience by standardizing toolbar presentation and removing duplicate refresh behavior.
- Added a single-instance guard for the Qt dashboard, preventing duplicate `run_gui.py` launches.
- Enhanced the Export Published PDFs workflow with an in-app progress dialog and completion confirmation before opening the export folder.
- Bumped the application version to `8.2.0` and updated user-facing docs to match.

#### Notes
- The export flow now gives visible feedback instead of immediately opening a file explorer with no status.
- This release preserves prior publish and club report generation functionality while focusing on desktop UI consistency.

---

## 8.2.2 Release Notes

This release adds improved RAES name review, audit refresh behavior for autopilot runs, and stronger manual audit logging for name corrections.

### Key improvements
- Added a new RAES `Name Review` panel mode to surface audited runner name candidates and allow manual review without changing existing review behavior.
- Ensured name corrections applied through RAES are recorded in the manual audit register so manual edits are tracked consistently.
- Persisted name alternate updates into `name_corrections.xlsx` when name changes are accepted through RAES.
- Fixed autopilot audit refresh so stale `outputs/audit/workbooks/` files are cleared before rebuilding audit reports, ensuring current audit outputs are generated on every run.
- Reduced Season Audit workbook output to the core review sheets and removed unused diagnostic-only sheets while keeping `Unrecognised Club Summary` for club review.
- Updated package metadata, installation guidance, and release docs to reflect version `8.2.2`.

---

## 8.2.3 Release Notes

This release refines audit actionability and improves the review workflow by focusing on manually checked items.

### Key improvements
- Expanded `Actionable Issues` to include category, club, gender, name variant, and data-invalid row issues that require manual checking.
- Removed the `Candidates To Check` and `EA Checked` audit sheets.
- Kept `Unrecognised Club Summary` for club review and retained the core audit review sheets.
- Updated package metadata and user-facing docs to version `8.2.3`.

---

## 8.3.0 Release Notes

This release adds a workbook comparison feature to the Qt dashboard and improves time-selection logic for scoring.

- Added `Compare Workbooks` support in the Qt dashboard so users can select two race workbooks and export a points-difference workbook.
- Added a standalone race workbook comparison utility in `scripts/race_compare.py` with both command-line and GUI modes.
- Added compare diagnostics for matched sheet pairs, runner/club column detection, scored row counts, and diff row totals.
- Updated scoring time column selection to prefer `Chip Time`, `Time`, or `Net Time` before falling back to `Gun Time`.
- Expanded `Actionable Issues` capture and audit workflow support for manual review.
- Updated package metadata, installation instructions, and changelog for release `8.3.0`.
