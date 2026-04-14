# WRRL League AI 8.2.0 Release Notes

This release focuses on GUI polish, workflow feedback, and consistency across the Qt dashboard.

## Highlights

- Updated Qt button styling and aligned refresh buttons across the enquiry, results, autopilot, RAES, and compare screens.
- Improved the Runner / Club Enquiry experience by standardizing toolbar presentation and removing duplicate refresh behavior.
- Added a single-instance guard for the Qt dashboard, preventing duplicate `run_gui.py` launches.
- Enhanced the Export Published PDFs workflow with an in-app progress dialog and completion confirmation before opening the export folder.
- Bumped the application version to `8.2.0` and updated user-facing docs to match.

## Notes

- The export flow now gives visible feedback instead of immediately opening a file explorer with no status.
- This release preserves prior publish and club report generation functionality while focusing on desktop UI consistency.
