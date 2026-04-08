# WRRL League AI 6.1.0 Release Notes

Date: 2026-04-03

## Summary

Release 6.1.0 is a quality and portability pass. No new features are introduced. All changes focus on robustness, resource safety, UI responsiveness, and cross-platform compatibility.

---

## Cross-Platform Portability

- **File open (`os.startfile`)** replaced with a `sys.platform` dispatch in all four call sites (`dashboard.py`, `issue_reviewer.py`, `results_viewer.py`, `raw_archive_diff_viewer.py`):
  - `win32` → `os.startfile`
  - `darwin` → `subprocess.run(["open", ...])`
  - Linux / other → `subprocess.run(["xdg-open", ...])`
- **Font paths in `timeline_generator.py`** are now platform-specific:
  - Windows: `%SystemRoot%\Fonts` (Segoe UI / Arial)
  - macOS: `/Library/Fonts/` and `/System/Library/Fonts/Supplemental/` (Arial)
  - Linux: DejaVu Sans and Liberation Sans from standard system font directories
  - PIL `load_default()` fallback fires if no path resolves

---

## Resource Leak Fixes

- All five `pd.ExcelFile(...)` usages in `audit_data_service.py` converted to `with pd.ExcelFile(...) as xl:` context managers. Previously, file handles were never explicitly closed, causing lock contention on Windows.
- `audit_viewer._load_sheet_options` now uses `with pd.ExcelFile` and surfaces any read error to the panel via `_show_message` instead of silently clearing the dropdown.
- `_load_recently_resolved_issue_keys` in `audit_data_service.py` now calls `xls.parse(sheet_name)` rather than opening the same file a second time with `pd.read_excel`.

---

## Atomic File Save

- `club_editor._on_save` now uses `_atomic_save` (imported from `manual_edit_service`) instead of a direct `wb.save(path)`. A crash mid-write no longer corrupts the target file; only a completed write replaces it.

---

## Threading / UI Responsiveness

- `check_all_runners._scan_all_races` is now offloaded to a daemon thread. The GUI no longer freezes when the panel is opened with a large race file set. Results are posted back to the main thread via `after(0, ...)`.
- `check_all_runners._apply_selected` is now offloaded to a daemon thread. The Apply button is disabled during the write, then the UI refreshes on completion.
- `raw_archive_diff_viewer._load_selected_diff` is now offloaded to a daemon thread with a "Loading…" status indicator.

---

## Correctness Fixes

- `audit_cleanser.create_cleansed_race_file` return-type annotation corrected from `Tuple[..., ..., ..., ...]` (4) to `Tuple[..., ..., ...]` (3).
- `issue_reviewer` `os.startfile` exception handler now catches `OSError` as well as `AttributeError`.
- `runner_history_viewer._on_runner_typed` and `_on_club_typed`: the `starts_lower` set is now hoisted before the list comprehension (was rebuilt per element — O(n²) per keypress).
- `check_all_runners._apply_selected` now removes only suggestions for files that were written successfully. Files that failed are kept in the list so the operator can retry without re-scanning.
- `check_all_runners._load_eligible_clubs` now sets a status-bar message on failure instead of returning silently.

---

## Repository Hygiene

- `output/autopilot/` and `output/*.png` added to `.gitignore`.
- Previously committed autopilot report artefacts removed from git tracking (`git rm --cached`).
- Temporary working documents removed from `documents/`:
  - `blank_club_in_inputs.md` (operator data pass checklist, superseded by Check All Runners panel)
  - `Health Check 31st March 26.docx` (working draft, superseded by `health_check_report.md`)

---

## Dataclass Invariants

- `ClubInfo.__post_init__` raises `ValueError` if `div_a` or `div_b` is not in `{1, 2}`.
- `TeamRaceResult.__post_init__` raises `ValueError` if `team_id` is not in `{'A', 'B'}` or `division` is not in `{1, 2}`.

---

## Operational Notes

- No new Python package dependencies.
- The application now runs correctly on Linux and macOS as well as Windows, with the exception of PDF conversion (`docx2pdf` still requires Microsoft Word on Windows or LibreOffice on Linux/macOS if configured separately).
