# Audit Todo List

This file tracks audit work that is intentionally deferred even though the rule direction is known.

## Input Folder Restructure (implemented in v6.0.0)

New folder layout under `inputs/`:

- `raw_data/` — race file downloads and Sporthive imports; the only place edits may be made
- `series/` — individual round files as downloaded; consolidation merges them into `raw_data/`
- `control/` — clubs.xlsx and events spreadsheet
- `audited/` — output of the audit/cleanse routines; source for scoring
- `raw_data_archive/` — write-once copy of each raw_data file (keyed by filename only)

Design decisions:

- Results are generated from `audited/` only; `source_loader.discover_race_files()` looks only there
- The autopilot always runs audit before scoring, so audited/ is always populated
- Raw data edits trigger a "dirty" state; audited file is only refreshed by re-running the audit routines
- Series consolidation checks club name consistency across rounds before writing to `raw_data/`; previous consolidated file is archived before replacement
- Archive key = filename only; re-archiving requires the reconcile_archive routine below

### Remaining TODO items

- **raw_data_diff**: comparison routine to detect post-archive modifications to raw_data files (raw vs archive mtime/hash)
- **reconcile_archive routine**: manual utility to allow deliberate overwrite of an archived file when the raw_data source has been intentionally corrected; requires explicit confirmation

## Deferred Workflow Items

- Build a suspected-name-variant review workflow similar to club matching.
- Show fuzzy suggestions for likely same-person name variants in a manual review popout.
- Keep all selections manual and approval-based.
- Do not allow automatic merge or automatic source-data correction.
- Decide whether approved name-variant resolutions should write to a maintained alias/identity lookup in a future release.
