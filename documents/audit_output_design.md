# Audit Output Design

This file captures the recommended first-pass layout and wording for audit output.

## Output Location And Naming

- Audit output folder: `outputs/audit`
- Input selection should allow the user to choose the race file set to review
- If exactly one race is selected, output file name: `Race N - Audit.xlsx`
- If multiple races are selected, output file name: `Season Audit.xlsx`
- Screen views should mirror the workbook sheets after the workbook is written

## Follow-On Audited File Stage

- After the audit output is trusted, audit may write a new derived workbook beside the original source race file in the same `inputs` folder.
- Default naming for that later stage: `Race N - Event Name (audited).xlsx`
- The original source file must never be overwritten.
- If the default audited filename already exists, the user must choose a new name.
- The audited file should preserve the source structure where practical, write cleaned values where the audit is confident, and add a `Comments` column for changes and unresolved issues.

## Workbook Sheets

Recommended first-pass sheets:

1. `Race Audit Summary`
2. `Actionable Issues`
3. `Row Audit`
4. `Runner Audit`
5. `Club Audit`
6. `Unrecognised Club Summary`
7. `Candidates To Check`
8. `EA Checked`

Keep one audit view per sheet and mirror the same views on screen.

## Recommended Columns

### Race Audit Summary

- `Severity Max`
- `Race`
- `Race File`
- `Scheme Status`
- `Error Count`
- `Warning Count`
- `Info Count`
- `Issue Codes`
- `Status`
- `Depends On`
- `Summary`
- `Next Step`

### Actionable Issues

- `Type`
- `Severity`
- `Issue Code`
- `Race`
- `Source Row`
- `Key`
- `Name`
- `Club`
- `Message`
- `Next Step`

### Row Audit

- `Severity`
- `Race`
- `Race File`
- `Source Row`
- `Issue Code`
- `Runner Name`
- `Time`
- `Club`
- `Sex`
- `Category`
- `Status`
- `Depends On`
- `Message`
- `Next Step`

### Runner Audit

- `Severity`
- `Runner Key`
- `Display Name`
- `Issue Code`
- `Clubs Seen`
- `Sexes Seen`
- `Categories Seen`
- `Races Seen`
- `Status`
- `Depends On`
- `Message`
- `Next Step`

### Club Audit

- `Severity`
- `Raw Club`
- `Issue Code`
- `Preferred Club`
- `Confidence`
- `Occurrences`
- `Races Seen`
- `Status`
- `Depends On`
- `Message`
- `Next Step`

### Unrecognised Club Summary

- `Raw Club`
- `Best Match`
- `Confidence`
- `Occurrences`
- `Races Seen`
- `Status`
- `Message`

## Status Wording

Use a short controlled set of status values.

- `Open`: issue is ready for review now
- `Dependent`: issue is valid, but another issue should usually be reviewed first
- `Manual Review`: issue needs a human decision rather than a straightforward correction
- `Ready To Fix`: source data can be corrected directly now
- `Informational`: noteworthy, but no action is required

## Dependency Wording

Use a short controlled set of dependency values.

- `None`
- `Club Lookup`
- `External Category Check`
- `Identity Review`
- `Source Data Correction`
- `Duplicate Conflict Review`

## Recommended Rules For Status And Dependency

### Missing Name With Result Data

- `Status`: `Ready To Fix`
- `Depends On`: `Source Data Correction`
- `Next Step`: `Correct the missing runner name in the source workbook`

### Unrecognised Club

- `Status`: `Open`
- `Depends On`: `Club Lookup`
- `Next Step`: `Review club suggestion and decide whether to add a club lookup conversion`

### Exact Name Collision Across Clubs

- `Status`: `Manual Review`
- `Depends On`: `Identity Review`
- `Next Step`: `Decide whether the matching names are the same runner or different people`

### Candidate Identity Sex Conflict

- `Status`: `Dependent`
- `Depends On`: `Identity Review`
- `Next Step`: `Resolve identity before escalating into gender or club inconsistency`

### Eligibility Inconsistency With Unresolved Club Lookup

- `Status`: `Dependent`
- `Depends On`: `Club Lookup`
- `Next Step`: `Review club mapping first, then confirm eligibility state`

### EA 5-Year Category Scheme Detected

- `Status`: `Open`
- `Depends On`: `None`
- `Next Step`: `Review derived categories and confirm the race used EA 5-year bands`

### External Category Data Check Required

- `Status`: `Dependent`
- `Depends On`: `External Category Check`
- `Next Step`: `Confirm the race category scheme before trusting downstream category findings`

### Routine Duplicate Runner

- `Status`: `Open`
- `Depends On`: `None`
- `Next Step`: `Review kept and discarded rows if traceability is needed`

### Duplicate Runner Attribute Conflict

- `Status`: `Manual Review`
- `Depends On`: `Duplicate Conflict Review`
- `Next Step`: `Check conflicting club, sex, or category values in the source rows`

## Sorting Recommendation

Sort within each sheet using this order:

1. `Severity`
2. `Status`
3. `Race`
4. `Source Row` or entity name

Use severity order `error`, `warning`, `info`.

Use status order `Manual Review`, `Dependent`, `Ready To Fix`, `Open`, `Informational`.

## Screen Presentation Recommendation

- `View Audit` should behave as a workbook/sheet browser and expose all workbook sheets.
- `Review Issues` should be a dedicated panel over `Actionable Issues` with:
	- code and race filters
	- open-source-file action
	- runner-history jump action
	- quick-fix actions where supported by issue code