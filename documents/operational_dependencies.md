# WRRL League AI - Operational Dependencies

This page captures runtime dependencies and environment assumptions that are not obvious from `requirements.txt` alone.

## Python Environment

- Python 3.11+ recommended.
- Install packages from the project root:

```sh
pip install -r requirements.txt
```

## Supported Platforms

As of v6.1.0, the application runs on Windows, macOS, and Linux:

- **Windows**: fully supported; all features including file open and PDF export.
- **macOS**: GUI and scoring pipeline supported; PDF export requires LibreOffice or a compatible converter unless Microsoft Word is installed.
- **Linux**: GUI and scoring pipeline supported; file open uses `xdg-open`; PDF export requires LibreOffice or a compatible converter.

## PDF Output Dependency

- The app generates DOCX reports with `python-docx`.
- PDF conversion uses `docx2pdf`.
- On Windows, `docx2pdf` requires Microsoft Word to be installed.
- On macOS/Linux, PDF output is skipped unless a compatible conversion tool is available. The app reports a warning rather than failing.

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
- Structured input subfolders under `inputs/`:
  - `raw_data/` (source race files and imports)
  - `series/` (series round source files)
  - `control/` (clubs/events/name corrections)
  - `audited/` (generated cleanse output used for scoring)
  - `raw_data_archive/` (write-once archive copies)
- Structured output subfolders under `outputs/`:
  - `publish/` (docx/pdf/xlsx published outputs)
  - `audit/` (workbooks and manual changes)
  - `quality/` (staged checks and data quality reports)
  - `autopilot/runs/` (autopilot reports)
- Input workbooks are expected to be `.xlsx`/`.xlsm`/`.xls` files with expected WRRL columns.

## User Settings Files

- Session prefs path: `~/.wrrl_prefs.json`
- Scoring settings path: `~/.wrrl_settings.json`

## Quick Health Checklist

Before a season run:

1. Confirm Python environment is active and requirements are installed.
2. Confirm Microsoft Word is installed if PDF output is required.
3. Run `playwright install chromium` at least once on each machine.
4. Confirm data root and season folders are configured in the dashboard.
5. Confirm `clubs.xlsx`, `wrrl_events.xlsx`, and `name_corrections.xlsx` are present in `inputs/control`.
6. Confirm raw race files are present in `inputs/raw_data` before running autopilot.
