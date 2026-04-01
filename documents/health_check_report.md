# WRRL League Scorer — Application Health Check Report

**Date:** June 2025  
**Version under review:** League-Report-Cleanup @ `37435ee` (= `main`)  
**Reviewer:** GitHub Copilot (Claude Sonnet 4.6)

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Architecture Overview](#2-architecture-overview)
3. [Data Processing Pipeline](#3-data-processing-pipeline)
4. [Error Handling](#4-error-handling)
5. [Configuration Management](#5-configuration-management)
6. [GUI Completeness](#6-gui-completeness)
7. [Dependencies](#7-dependencies)
8. [Repository Health](#8-repository-health)
9. [Testing](#9-testing)
10. [Portability](#10-portability)
11. [Catalogue of Issues](#11-catalogue-of-issues)
12. [Recommendations](#12-recommendations)

---

## 1. Executive Summary

The application is a well-structured Python scoring system for the Wiltshire Road Running League (WRRL). The processing core — individual ranking, team scoring, season aggregation, and report generation — is in good shape and produces correct output across 8 races. The codebase is clean following the recent code-quality pass.

However, several issues warrant attention before the application is considered production-ready. The most significant are the **complete absence of automated tests**, **data/output files committed to git** (polluting the repository and blocking a clean diff), and a **deleted `.gitignore`** that leaves the repository vulnerable to future accidental commits of generated artefacts. There are also minor version-string inconsistencies, hardcoded season constants that should be configurable, and two placeholder GUI features that are not yet implemented.

**Overall assessment: Good foundation. Address the repository hygiene and testing gap before the next season.**

| Area | Status |
|---|---|
| Processing pipeline correctness | ✅ Sound |
| Error handling | ✅ Good |
| GUI — core scorer workflow | ✅ Complete |
| GUI — View Results / Settings | ⚠️ Placeholders |
| Version consistency | ⚠️ Mismatch |
| Hardcoded season constants | ⚠️ Not configurable |
| Dependencies | ⚠️ Stale comment in requirements.txt |
| Repository hygiene | ❌ No .gitignore, output files committed |
| Automated testing | ❌ None |

---

## 2. Architecture Overview

### Module Map

```
league_scorer/
  __init__.py              Package entry, v3.1 declaration
  models.py                All data model dataclasses (RunnerRaceEntry etc.)
  exceptions.py            FatalError, ValidationError hierarchy
  session_config.py        Singleton year/path config + .wrrl_prefs.json persistence
  club_loader.py           Loads clubs.xlsx → {club_name: ClubInfo}
  events_loader.py         Loads events.xlsx → List[EventRecord]
  normalisation.py         Runner name deduplication and variant mapping
  race_processor.py        Parses a single race CSV → List[RunnerRaceEntry]
  individual_scoring.py    Per-race finishing position → points, gender/category rank
  team_scoring.py          5-scorer team points (TEAM_SIZE=5, MAX_DIV_PTS=20)
  season_aggregation.py    Best-6-of-8 cumulative totals (BEST_N=6)
  output_writer.py         Writes Results.xlsx (5-sheet workbook) + unused clubs.xlsx
  report_writer.py         Writes per-race DOCX/PDF + combined league update DOCX/PDF
  main.py                  LeagueScorer orchestration class

  graphical/
    __init__.py            Exports: launch, launch_dashboard, EventsViewerWindow
    dashboard.py           Main tkinter window (config, navigation, branding)
    gui.py                 LeagueScorerApp panel (race selector, log, run pipeline)
    events_viewer.py       Toplevel events schedule (sortable Treeview)
    timeline_generator.py  Pillow-based season timeline PNG generator
```

### Entry Points

| Entry point | Usage |
|---|---|
| `python -m league_scorer` (`__main__.py`) | Headless CLI run |
| `run_gui.py` | Launches GUI (suitable for IDLE/double-click) |
| `graphical.launch_dashboard()` | Programmatic GUI launch |

### Strengths

- Clear single-responsibility separation across modules.
- `models.py` holds all data classes; no data definitions are scattered through logic files.
- `exceptions.py` provides a clean two-level hierarchy (`FatalError` → pipeline abort; `ValidationError` → per-race warning).
- The GUI uses a **background thread with a queue** for the pipeline, keeping the UI responsive.
- `session_config.py` persists user preferences (year, data root) across restarts via `.wrrl_prefs.json`.

---

## 3. Data Processing Pipeline

### Flow (per run)

```
clubs.xlsx → club_loader
events.xlsx → events_loader
Race N CSV → race_processor → normalisation → individual_scoring
                                              team_scoring
                          (all races combined) → season_aggregation
                                              → output_writer (Results.xlsx)
                                              → report_writer (DOCX/PDF)
```

### Individual Scoring (`individual_scoring.py`)

Assigns points based on finishing position within division (Div 1 / Div 2), with separate gender (M/F) and age-category rankings. Tie-breaking is handled correctly: runners with identical chip times share a points value and the next position is skipped.

**Assessment:** Correct and well-implemented.

### Team Scoring (`team_scoring.py`)

A team's score is the sum of points from its top `TEAM_SIZE = 5` finishers. Maximum team score is capped at `MAX_DIV_PTS = 20`. Teams must have at least 3 finishers to qualify.

**Assessment:** Correct. The 3-scorer minimum, 5-scorer cap, and 20-point maximum are all implemented clearly. However, these constants are hardcoded (see §11-1).

### Season Aggregation (`season_aggregation.py`)

Takes all race entries for a runner and sums their best `BEST_N = 6` scores from up to `MAX_RACES = 8` races. This correctly implements the "best 6 of 8" scoring rule.

**Assessment:** Correct. `BEST_N` and `MAX_RACES` are hardcoded (see §11-1).

### Name Normalisation (`normalisation.py`)

Handles name variants across multiple race CSVs (e.g., "J. Smith" vs "John Smith") via a deterministic matching and alias-building process. This is one of the more complex parts of the codebase and is implemented carefully.

**Assessment:** Sound. No issues identified.

### Race File Processing (`race_processor.py`)

Reads a CSV, validates column presence, maps club names to `ClubInfo`, and emits `RunnerRaceEntry` objects. Unrecognised clubs are collected separately and written to `unused clubs.xlsx`.

**Assessment:** Correct. Validation raises `ValidationError` for missing columns; missing club is soft-handled (recorded, not fatal).

---

## 4. Error Handling

### Exception Hierarchy

```
Exception
└── FatalError      — pipeline abort (unrecoverable)
└── ValidationError — per-race warning (skipped, pipeline continues)
```

Both are defined in `exceptions.py` and used consistently.

### GUI Error Path (`gui.py`)

The background thread catches `FatalError` and generic `Exception` separately, both resulting in a visible error message in the log panel. The UI is re-enabled after either outcome. This is correct and safe.

### File I/O

- `output_writer.py` does not explicitly catch `OSError` when writing Excel files. An output-path permission error would propagate as an unhandled exception and surface as "Unexpected error" in the GUI log. This is acceptable but worth noting.
- `session_config.py` correctly silences `OSError` on preferences save (non-fatal).
- `report_writer.py` wraps `docx2pdf` import in a `try/except ImportError` and the PDF save in a `try/except`, silently skipping PDF generation on failure. **The user receives no prominent warning when PDF output is skipped** (see §11-5).

### Assessment

Error handling is well-structured. The one gap is the silent PDF-skip behaviour.

---

## 5. Configuration Management

### `session_config.py` — SessionConfig Singleton

Manages:
- `year` (current season, e.g. 2025)
- `data_root` (top-level data folder path)
- Derived paths: `input_dir`, `output_dir`, `events_file`
- `available_years()` — scans `data_root` for year-named subdirectories

Persisted to `~/.wrrl_prefs.json` (home directory) via `save()` / `load()`. The file format is plain JSON; load is guarded against `json.JSONDecodeError` and `ValueError`.

**Assessment:** Well-implemented for a single-user desktop tool.

### Hardcoded Season Constants

Several business-rule constants are embedded directly in source files rather than centralised or user-configurable:

| Constant | Value | Location | Rule |
|---|---|---|---|
| `BEST_N` | 6 | `season_aggregation.py` | Best 6 scores count |
| `MAX_RACES` | 8 | `output_writer.py` | Season has 8 races |
| `TEAM_SIZE` | 5 | `team_scoring.py` | Top 5 scorers per team |
| `MAX_DIV_PTS` | 20 | `team_scoring.py` | Max team points per race |
| `_SEASON_FINAL_RACE` | 8 | `report_writer.py` | Race at which end-of-season narrative triggers |

These constants are consistent with each other and match real league rules. However, they are scattered across five different files. If the league format changes, all five must be updated manually.

### Year Derivation Fallback

In `main.py`, the current year is derived from `output_dir.parent.name` (e.g. `.../2025/outputs` → `2025`). There is a hardcoded fallback to `2026` in two places:

```python
year = int(output_dir.parent.name)  # may fall back to 2026
```

If the directory naming convention ever changes (e.g. a flat layout), the fallback will silently produce incorrect year labels in reports without raising any error.

---

## 6. GUI Completeness

### Dashboard (`dashboard.py`)

| Button / Feature | Status |
|---|---|
| Run Scorer | ✅ Fully implemented |
| View Events | ✅ Fully implemented (opens EventsViewerWindow) |
| View Timeline | ✅ Fully implemented (generates PNG, opens in system viewer) |
| View Results | ⚠️ Placeholder — shows messagebox only |
| Settings | ⚠️ Placeholder — shows messagebox only |
| Year selector | ✅ Functional |
| Data root picker | ✅ Functional |

### Scorer Panel (`gui.py`)

The scorer panel is fully implemented:
- Race file discovery from input directory
- Scrollable checkbox selector with Select All / None
- Background threading with animated progress bar
- Colour-coded log output (INFO, WARNING, ERROR, SUCCESS)
- Duplicate-line collapsing with repeat counter (×N)

**Assessment:** Core workflow is complete. "View Results" and "Settings" are stubs that display a "coming soon" messagebox. These are acceptable for a current release but should be noted as outstanding work.

### Events Viewer (`events_viewer.py`)

A sortable, colour-coded Treeview showing the full Championship Events schedule. Status colours (upcoming/completed/TBC) are applied via row tags. Functional and complete.

### Timeline Generator (`timeline_generator.py`)

Generates a race-progress timeline as a 1500×700 PNG using Pillow. Layout constants are all centralised at the top of the file. Font loading uses hardcoded Windows paths with a fallback to the Pillow default bitmap font. See §10 for portability note.

---

## 7. Dependencies

### `requirements.txt`

```
pandas>=1.5.0
openpyxl>=3.0.10
Pillow>=10.0.0
python-docx>=1.1.0
docx2pdf>=0.1.8
```

**Observations:**

1. **`docx2pdf>=0.1.8`** is listed as a hard requirement but is treated as optional at runtime (wrapped in `try/except ImportError`). On a system where `docx2pdf` cannot produce PDFs (e.g., no Microsoft Word installed), the install will still succeed but PDF output will silently be skipped. The `requirements.txt` could carry a comment noting the Word dependency, or the package could be moved to an `extras_require` section if a `pyproject.toml` is ever introduced.

2. **Stale project structure comment.** `requirements.txt` contains a block comment listing the project file structure that still references the deleted `scorer.py`:
   ```
   # league_scorer/scorer.py
   ```
   This should be removed.

3. **No version upper bounds.** This is generally acceptable for a single-user application with a managed venv, but `pandas` in particular has had breaking API changes between major versions. Pinning a tested range (e.g. `pandas>=1.5.0,<3.0`) would improve reproducibility.

4. **No `tkinter` or `threading` listed.** Both are stdlib modules, so no entry is needed — this is correct.

---

## 8. Repository Health

### `.gitignore` Deleted

The `.gitignore` file shows as **deleted** in `git status`:

```
deleted:    .gitignore
```

This means nothing is being excluded from `git add`. The following artefacts are currently untracked or modified-but-committed and could be accidentally staged in a future commit:

- `data/2025/outputs/*.docx`, `*.pdf`, `*.xlsx` — generated output files (already committed from an earlier session)
- `league_scorer/__pycache__/` — Python bytecode (untracked)
- `league_scorer/graphical/__pycache__/dashboard.cpython-314.pyc` — bytecode (modified/committed)
- `.venv/` — virtual environment
- `output/division-format-check/` — scratch output folder (untracked)
- `data/2025/outputs/sample.docx` — manual edit sample (untracked)

### Generated Output Files in Repository

The `data/2025/outputs/` folder contains all generated race reports (`.docx`, `.pdf`, `.xlsx`) committed to git. These are build artefacts — they change every time the scorer runs and should not be under version control. They inflate repository size and make `git diff` and PR reviews noisy.

**This is the most significant repository health issue.**

### `__pycache__` Partially Committed

`league_scorer/graphical/__pycache__/dashboard.cpython-314.pyc` is tracked by git and shows as modified. Bytecode files should never be committed.

### Commit History

```
37435ee  Code cleanup (current)
62765de  Refine league report layout and narrative
cd4f2d8  Refine report layout and bump version to 2.1
d2821ef  Clean up outputs and simplify category reporting
a6e55a9  Initial commit
```

Five commits total. History is linear and well-described. No merge commits or force-pushes in history.

### Branch State

Active branch `League-Report-Cleanup` is fully synced with `origin/League-Report-Cleanup` and `origin/main` — all three point to `37435ee`. No divergence.

---

## 9. Testing

**There are zero automated tests in this codebase.**

No test files, no test directory, no test runner configuration (`pytest.ini`, `setup.cfg [tool:pytest]`, `pyproject.toml`).

This is the most significant quality risk for a scoring application where correctness is critical. A miscalculation in points, tie-breaking, team composition, or season aggregation could go unnoticed until a club queries their standings.

### Suggested Test Coverage Priorities

| Module | Test scenarios |
|---|---|
| `individual_scoring.py` | Correct points for positions 1–N; tie-breaking (equal chip time); division boundary |
| `team_scoring.py` | Top-5 selection; fewer than 3 finishers; 20-point cap; tie resolution |
| `season_aggregation.py` | Best-6-of-8 selection; runner with fewer than 6 races; runner with exactly 6 races |
| `normalisation.py` | Exact match; known variant match; new runner creation; initials match |
| `race_processor.py` | Valid CSV parse; missing column → ValidationError; unrecognised club handling |
| `report_writer.py` | Race 8 triggers end-of-season narrative; Race 7 triggers mid-season narrative |

A suite of ~30 unit tests for the above would provide meaningful regression coverage. `pytest` with `pytest-cov` is the recommended toolchain.

---

## 10. Portability

The application is currently **Windows-only by design** (the target environment is a Windows desktop). The following items would need addressing before cross-platform use:

1. **`timeline_generator.py` — hardcoded Windows font paths:**
   ```python
   "C:/Windows/Fonts/segoeuib.ttf"
   "C:/Windows/Fonts/segoeui.ttf"
   ```
   A fallback to the Pillow default bitmap font is in place, so the generator does not crash on non-Windows, but the rendered timeline will use a lower-quality font. If portability is ever needed, use `matplotlib` font resolution or bundle a TTF.

2. **GUI font `"Segoe UI"`** is used throughout `dashboard.py` and `gui.py`. This font is Windows-only. On macOS/Linux tkinter will silently substitute a system font; visual fidelity will degrade but no crash will occur.

3. **`docx2pdf`** requires Microsoft Word to be installed on the host machine. This is an inherent Windows/macOS constraint. The try/except wrapper handles missing installation gracefully.

4. **No POSIX path issues.** All path handling uses `pathlib.Path` throughout. No raw string path literals with backslashes were found.

---

## 11. Catalogue of Issues

### Severity: High

| # | Issue | Location | Impact |
|---|---|---|---|
| H-1 | No automated tests | Entire codebase | Scoring bugs could go undetected |
| H-2 | `.gitignore` deleted | Repository root | Build artefacts risk being committed |
| H-3 | Generated output files committed to git | `data/2025/outputs/` | Repository bloat, noisy diffs, poor hygiene |

### Severity: Medium

| # | Issue | Location | Impact |
|---|---|---|---|
| M-1 | Version string mismatch: `__init__.py` says `v3.1`, report footer says `v2.1` | `__init__.py`, `report_writer.py` | User confusion; unclear which is authoritative |
| M-2 | "View Results" button is a placeholder | `dashboard.py` `_on_view_results()` | Feature gap — user cannot view results from GUI |
| M-3 | "Settings" button is a placeholder | `dashboard.py` `_on_settings()` | Feature gap — no in-app configuration UI |
| M-4 | Season constants hardcoded in 5 separate files | `season_aggregation.py`, `output_writer.py`, `team_scoring.py`, `report_writer.py` | Fragile if league rules change |
| M-5 | Year fallback silently produces `2026` if dir convention changes | `main.py` (×2) | Silent data integrity bug |
| M-6 | `__pycache__` bytecode tracked by git | `graphical/__pycache__/` | Repository pollution |

### Severity: Low

| # | Issue | Location | Impact |
|---|---|---|---|
| L-1 | PDF generation silently skipped with no prominent user warning | `report_writer.py` | User may not realise no PDF was produced |
| L-2 | Stale project-structure comment in `requirements.txt` still references deleted `scorer.py` | `requirements.txt` | Minor documentation inaccuracy |
| L-3 | No version upper bounds on dependencies | `requirements.txt` | Possible breakage on future major releases |
| L-4 | Windows-only font paths in timeline generator | `timeline_generator.py` | Visual degradation on non-Windows (not a current concern) |

---

## 12. Recommendations

### Immediate (before next season)

1. **Restore `.gitignore`** — at minimum exclude `__pycache__/`, `*.pyc`, `.venv/`, `output/`, and ideally move generated output files out of the tracked tree.

2. **Move generated outputs out of git** — add `data/*/outputs/` to `.gitignore` (once restored) and remove the currently tracked output files with `git rm --cached`. Generated reports are artefacts, not source.

3. **Fix version mismatch** — decide on a single authoritative version string. Recommend updating `report_writer.py`'s footer constant to match `__init__.py` (`v3.1`), then keeping both in sync via a `__version__` constant in `__init__.py`.

4. **Remove stale `requirements.txt` comment** referencing `scorer.py`.

### Short-term (next development sprint)

5. **Add a `constants.py` module** (or add to `session_config.py`) centralising all season rules: `BEST_N`, `MAX_RACES`, `TEAM_SIZE`, `MAX_DIV_PTS`, `SEASON_FINAL_RACE`. Import from there in all consuming modules.

6. **Harden year derivation** — raise a clear `FatalError` rather than falling back silently to `2026` when the output directory does not follow the `YEAR/outputs` convention.

7. **Surface PDF failure prominently** — when `docx2pdf` conversion fails or is unavailable, log a `WARNING` that is conspicuous in the GUI log (currently the failure is silently swallowed).

8. **Implement _View Results_** — the most useful outstanding feature. A `Results.xlsx` viewer panel or simple summary display would complete the GUI workflow.

### Ongoing

9. **Write a test suite.** Start with `pytest` and target the six modules identified in §9. Aim for coverage of all scoring-rule edge cases. Even 30 well-chosen unit tests would substantially reduce the risk of a silent scoring regression.

10. **Add `docx2pdf` installation note to `requirements.txt`** (or `README`) clarifying that PDF output requires Microsoft Word to be installed.

---

*End of report.*
