# Manual Update Automation Findings

## Source

- File analyzed: `quarantine and copy/Manual_Data_Audit.xlsx`
- Sheet: `Manual Changes`
- Total rows: 151

## Key patterns

### 1. Category normalization is the dominant pattern

- `category` corrections: 140 rows
- Most frequent old→new mappings:
  - `FV45` → `V50` (12 rows)
  - `LV40` → `V40` (12 rows)
  - `Vet 50` → `V50` (11 rows)
  - `Vet 40` → `V40` (10 rows)
  - `FV35` → `V40` (8 rows)
  - `FV45` → `V40` (7 rows)
  - `FV55` → `V50` (5 rows)
  - `Ages 35 - 44` → `V40` (5 rows)
  - `Ages 45 - 54` → `V40` (5 rows)
  - `FV65+` → `V60` (4 rows)
- There is also a repeated mapping of `MSen` → `Sen`.

### 2. Runner-based corrections are repeated

- 44 runners appear in 2 or more rows.
- Examples of frequent runners:
  - `Peter Campbell` (5 rows)
  - `Matthew Waite` (5 rows)
  - `Victoria Matar` (4 rows)
  - `Timothy Eddy` (4 rows)
  - `Jack Clarke` (4 rows)
- This suggests automation could focus on applying a single runner's resolved category consistently across all race files.

### 3. Club normalization is present but lower volume

- `club` corrections: 6 rows
- Most common club fix:
  - `AVON VALLEY RUNNERS` → `Avon Valley Runners` (3 rows)

### 4. Data quality issues to clean before automation

- Blank `Field` entries: 4 rows
- Blank `Runner` entries: 4 rows
- These likely indicate malformed audit rows and should be validated or ignored before automation.

## Strong automation candidates

### A. Normalize category labels

Implement a rule-based category mapper to convert legacy labels into canonical codes.
- Possible mappings:
  - `FV45`, `Vet 50`, `F55`, `Ages 45 - 54` → `V50`
  - `LV40`, `Vet 40`, `FV35`, `Ages 35 - 44`, `Ages 40 - 49` → `V40`
  - `FV55`, `Ages 55 +` → `V50` or `V60` depending on rules
  - `FV65+` → `V60`
  - `MV60` → `V60`
  - `Vet 70+`, `MV70+` → `V70`
  - `MSen`, `Senior` → `Sen`
- This would likely automate the majority of manual category edits.

### B. Apply runner resolutions across files

- Use the existing runner-field propagation approach to apply resolved categories for a runner across all raw input files.
- This is especially useful when a runner appears in multiple race files with inconsistent category labels.

### C. Normalize club aliases

- Add a small alias map for club names such as `AVON VALLEY RUNNERS` → `Avon Valley Runners`.
- This is a lower-volume but clearly repeated fix.

### D. Guard against invalid audit rows

- Reject or flag manual audit rows with no `Field` or no `Runner`.
- Add validation in the importer prior to automation.

## Recommendations for later implementation

1. Collect more automation proposals in `AI_Automation/` before coding.
2. Prioritize category normalization first, since it is the largest repeated set of manual corrections.
3. Next, add runner-level propagation logic using the existing UI/service patterns.
4. Finally, add club alias normalization and audit-row validation.

## Notes

- Existing code already contains relevant service hooks:
  - `league_scorer/manual_edit_service.py`
  - `league_scorer/manual_data_audit.py`
- The current workbook confirms concrete, repeatable automation opportunities.
