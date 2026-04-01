# Audit Test Matrix

This file tracks concrete scenarios for the planned v4.0 audit feature.

## Status Key

- `planned`: scenario agreed, not yet implemented
- `ready`: scenario is specified clearly enough for implementation
- `done`: scenario has been implemented and verified

## Test Cases

| Test ID | Area | Scenario | Expected Audit Result | Status |
| --- | --- | --- | --- | --- |
| `AUD-TST-001` | Row audit | A row has a valid time but the runner name cell is blank. | Raise `AUD-ROW-005` and report the row as missing a runner name with result data. | ready |
| `AUD-TST-001A` | Row audit | A missing-name row is reported in audit. | The row-audit entry should show race, source row, captured time, and any available club, category, or sex values so the source row can be found quickly. | ready |
| `AUD-TST-002` | Row audit | A workbook contains fully blank spacer rows with no runner name and no time. | Do not raise a formal audit issue for the blank spacer rows. | ready |
| `AUD-TST-003` | Club audit | An unrecognised raw club appears in several races and has one strong fuzzy match to a known club. | Include one aggregated row in `Unrecognised Club Summary` with count, races seen, best match, and confidence score. | planned |
| `AUD-TST-004` | Club audit | An unrecognised raw club appears with no credible fuzzy match. | Include the raw club in `Unrecognised Club Summary` with no accepted mapping and low or blank confidence. | planned |
| `AUD-TST-005` | Club audit | A reviewer accepts the suggested best match for an unrecognised club in the audit UI. | A popout review box with `Tick`, `Current Club`, `Proposed Club`, and `Message` should allow the approved `raw club -> preferred club` pairing to be written into the club mapping file through `Add conversions to Club Lookup`, with confidence, occurrence count, and races seen shown inside `Message`. All ticks should start cleared, no bulk-select actions should be available, and a confirmation summary should appear before writing changes. | ready |
| `AUD-TST-005A` | Club audit | A reviewer selects a club conversion where the raw club already exists in the club lookup file with the same proposed mapping. | The confirmation summary should warn that an exact mapping already exists before writing so the user can avoid redundant updates. | ready |
| `AUD-TST-005B` | Club audit | A reviewer selects a club conversion where the raw club already exists in the club lookup file with a different mapping. | The confirmation summary should warn that a conflicting mapping already exists before writing and provide a quick return path so the user can go back and untick that selection before any changes are applied. Returning should preserve the same selection state. | ready |
| `AUD-TST-005C` | Club audit | A reviewer returns from a conflicting existing-match warning to the selection popout. | The same selection state should reopen and the conflicting selected rows should be visibly highlighted so the user can untick them quickly. | ready |
| `AUD-TST-006` | Category audit | After category normalisation, female runner categories in one race are all standard league bands such as `V40`, `V50`, and `V60`. | Do not raise `AUD-RACE-006` or `AUD-RACE-007`; treat the scheme as already aligned with league categories. | ready |
| `AUD-TST-007` | Category audit | After category normalisation, even one female runner category in a race proves EA-style 5-year usage, such as `V35`, `V40`, `V45`, or `V50`. | Raise `AUD-RACE-006` and derive league categories for affected rows, reporting original and derived values. No minimum threshold is required. | ready |
| `AUD-TST-008` | Category audit | After category normalisation, female runner categories in one race show a mixed or ambiguous pattern such as `V35`, `V45`, and `V55`. | Raise `AUD-RACE-007` and flag the race for External Data Check rather than silently forcing a full conversion. | ready |
| `AUD-TST-008A` | Category audit | A race is flagged `External Data Check` because female categories are mixed or ambiguous. | The race-audit entry should list the distinct female categories that triggered the warning. | ready |
| `AUD-TST-009` | Runner audit | The same runner changes category across races in a way that matches expected birthday progression from the category first seen in chronological race order. | Use the first chronological race as the baseline category and keep this at warning level only if surfaced at all; do not escalate to an error by default. | ready |
| `AUD-TST-009A` | Runner audit | A runner appears in races that include one race flagged `External Data Check`. | Keep downstream runner audit findings visible. If needed, mark the category-based finding as dependent on the unresolved race-level category issue rather than suppressing it. | ready |
| `AUD-TST-009B` | Runner audit | The same normalised runner name appears across races for two different clubs, with no stronger identity evidence than the exact name match. | Raise `AUD-RUNNER-007` as an ambiguous identity collision. Do not auto-merge the season records and do not raise `AUD-RUNNER-003` from name match alone. | ready |
| `AUD-TST-009C` | Runner audit | A candidate same-runner match shares name and club context, but sex differs across records. | Raise `AUD-RUNNER-008` as a warning and stop automatic identity escalation for that candidate set. Do not continue into gender inconsistency or club inconsistency automatically. | ready |
| `AUD-TST-009D` | Runner audit | A manually trusted or otherwise confirmed runner identity appears across races for two different eligible clubs. | Raise `AUD-RUNNER-003` for club inconsistency because the identity is trusted beyond simple exact name matching. | ready |
| `AUD-TST-009E` | Runner audit | A runner appears eligible in one race and non-league in another because club lookup is unresolved in one case. | Still raise `AUD-RUNNER-004`, but mark it as dependent on club lookup review rather than suppressing it. | ready |
| `AUD-TST-009F` | Runner audit | Two similar names appear likely to be the same person. | Raise `AUD-RUNNER-005` with fuzzy suggestions for review only. Do not auto-merge records or correct data in audit. | ready |
| `AUD-TST-010` | Duplicate audit | Duplicate runner rows in one race have the same club, category, and sex but different times. | Consolidate to the quickest time and report the duplicate under `AUD-ROW-008`. | ready |
| `AUD-TST-010A` | Duplicate audit | A duplicate runner warning is shown for a routine deduplication case. | The audit entry should show the kept row and discarded row references so the deduplication decision is traceable. | ready |
| `AUD-TST-011` | Duplicate audit | Duplicate runner rows in one race disagree on club, category, or sex. | Keep the quickest time in results output and raise `AUD-ROW-010` separately in audit for attribute conflict. | ready |
| `AUD-TST-011A` | Duplicate audit | A race contains one or more duplicate attribute conflicts. | The race audit should surface a high-severity duplicate conflict summary in addition to the row-level issue entries. | ready |
| `AUD-TST-012` | Process status | DOCX output succeeds but PDF export fails. | Show a process warning only; do not create an audit issue. | ready |
| `AUD-TST-013` | Audit output | The first audit test drive is run against a realistic season dataset. | Emit all current audit issue types, including informational items, into a workbook with one audit view per sheet, then show those same audit views on screen. | ready |
| `AUD-TST-014` | Audit output | The user selects exactly one race for audit. | Write the workbook into `outputs/audit` using the file name `Race N - Audit.xlsx`. | ready |
| `AUD-TST-015` | Audit output | The user selects multiple races for audit. | Write the workbook into `outputs/audit` using the file name `Season Audit.xlsx`. | ready |
| `AUD-TST-016` | Audit output | A generated audit sheet is reviewed in workbook or on screen. | The sheet should include `Status`, `Depends On`, and `Next Step` columns using the controlled wording defined for the first audit pass. | ready |
| `AUD-TST-017` | Audited file | The user requests a cleansed version of a single race file after audit review. | Write a new workbook beside the source file in the same `inputs` folder using the default name `Race N - Event Name (audited).xlsx`. | ready |
| `AUD-TST-018` | Audited file | The default audited file name already exists. | Do not overwrite it. Require the user to choose a new name with the standard rename or save dialog. | ready |
| `AUD-TST-019` | Audited file | A cleansed race workbook is generated from audit output. | Preserve the original row order and source structure where practical, write normalised values where the audit is confident, and add a `Comments` column for changes and unresolved issues. | ready |
| `AUD-TST-020` | Audited file | Some audit findings remain unresolved when the cleansed file is written. | Keep those unresolved items visible in the `Comments` column rather than pretending they were fixed. | ready |