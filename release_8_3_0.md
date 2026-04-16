# WRRL League AI 8.3.0 Release Notes

This release adds a workbook comparison feature to the Qt dashboard and improves time-selection logic for scoring.

- Added `Compare Workbooks` support in the Qt dashboard so users can select two race workbooks and export a points-difference workbook.
- Added a standalone race workbook comparison utility in `scripts/race_compare.py` with both command-line and GUI modes.
- Added compare diagnostics for matched sheet pairs, runner/club column detection, scored row counts, and diff row totals.
- Updated scoring time column selection to prefer `Chip Time`, `Time`, or `Net Time` before falling back to `Gun Time`.
- Expanded `Actionable Issues` capture and audit workflow support for manual review.
- Updated package metadata, installation instructions, and changelog for release `8.3.0`.
