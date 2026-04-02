# WRRL League Scorer - Operational Dependencies

This page captures runtime dependencies and environment assumptions that are not obvious from `requirements.txt` alone.

## Python Environment

- Python 3.11+ recommended.
- Install packages from the project root:

```sh
pip install -r requirements.txt
```

## PDF Output Dependency

- The app generates DOCX reports with `python-docx`.
- PDF conversion uses `docx2pdf`.
- On Windows, `docx2pdf` requires Microsoft Word to be installed.

If Word is not available, the app will still generate DOCX output and report a PDF conversion warning.

## Browser Automation Dependency (Playwright)

- `playwright` is used by import/audit tooling that relies on browser automation.
- After installing dependencies, install browser binaries once:

```sh
playwright install chromium
```

If browser binaries are missing, related import flows may fail until Playwright browsers are installed.

## Data/Layout Assumptions

- The scorer expects the configured data root to follow:
  - `{data_root}/{year}/inputs`
  - `{data_root}/{year}/outputs`
- Input workbooks are expected to be `.xlsx` files with expected WRRL columns.

## User Settings Files

- Session prefs path: `~/.wrrl_prefs.json`
- Scoring settings path: `~/.wrrl_settings.json`

## Quick Health Checklist

Before a season run:

1. Confirm Python environment is active and requirements are installed.
2. Confirm Microsoft Word is installed if PDF output is required.
3. Run `playwright install chromium` at least once on each machine.
4. Confirm data root and season folders are configured in the dashboard.
5. Confirm `clubs.xlsx` and events workbook are present in the season inputs folder.
