# WRRL League AI 8.2.2 Release Notes

This release adds improved RAES name review, audit refresh behavior for autopilot runs, and stronger manual audit logging for name corrections.

Key improvements:

- Added a new RAES `Name Review` panel mode to surface audited runner name candidates and allow manual review without changing existing review behavior.
- Ensured name corrections applied through RAES are recorded in the manual audit register so manual edits are tracked consistently.
- Persisted name alternate updates into `name_corrections.xlsx` when name changes are accepted through RAES.
- Fixed autopilot audit refresh so stale `outputs/audit/workbooks/` files are cleared before rebuilding audit reports, ensuring current audit outputs are generated on every run.
- Reduced Season Audit workbook output to the core review sheets and removed unused diagnostic-only sheets while keeping `Unrecognised Club Summary` for club review.
- Updated package metadata, installation guidance, and release docs to reflect version `8.2.2`.
