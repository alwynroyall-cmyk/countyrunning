# Audit Issue Register

This is the working catalogue for the planned audit feature targeting v4.0.

Purpose:

- define the classic data-quality issue types
- classify each issue by severity and scoring impact
- give each issue a stable code for outputs, review, and tests

Status:

- draft 1 for review
- intended to be edited as rules are added, merged, split, or removed

## Severity Model

- `error`: data cannot be trusted for scoring without review or correction
- `warning`: scoring can continue, but a human should review the case
- `info`: unusual or notable, but not normally action-blocking

## Scoring Impact

- `high`: could change individual points, team scores, standings, or eligibility
- `medium`: could affect interpretation, audit confidence, or downstream review
- `low`: mostly descriptive or housekeeping

## Scope

- `row`: tied to a single source row in a race file
- `runner`: tied to one runner across one or more races
- `race`: tied to a whole race file or processing run for that race
- `season`: tied to cross-race consistency or whole-season outputs
- `club`: tied to club mapping or club configuration data

## Issue Types

| Code | Issue Type | Severity | Impact | Scope | Description | Suggested Audit Handling |
| --- | --- | --- | --- | --- | --- | --- |
| `AUD-RACE-001` | Missing Required Column | error | high | race | A race file is missing one or more required columns such as Name, Club, Gender, or Category. | Block race from scoring and report exact missing columns. |
| `AUD-RACE-002` | Missing Time Column | error | high | race | No valid time-like column can be identified in the race file. | Block race from scoring and report expected column names. |
| `AUD-RACE-003` | Race File Unreadable | error | high | race | Workbook cannot be opened or parsed. | Block race and report file-level failure. |
| `AUD-RACE-004` | Duplicate Race Number File | warning | medium | race | More than one file resolves to the same race number. | Keep chosen file, report ignored file names. |
| `AUD-RACE-005` | Race Skipped | error | high | race | A whole race was skipped because of a fatal race-level processing failure. | Record explicit reason in race audit summary. |
| `AUD-RACE-006` | EA 5-Year Category Scheme Detected | warning | medium | race | After category normalisation, the female runner categories include any EA-style 5-year veteran band evidence, which is enough to show that the race is using EA-style categories rather than pure league scoring bands. | Record that category normalisation was derived from EA-style source categories and allow row-level derived-category warnings where applicable. |
| `AUD-RACE-007` | External Category Data Check Required | warning | high | race | After category normalisation, the female runner categories show a mixed or ambiguous scheme such as `V35`, `V45`, `V55`, indicating that the race cannot be trusted for direct league-category mapping. | Flag race for External Data Check and avoid silent remapping where evidence is weak. |
| `AUD-ROW-001` | Invalid Gender For Eligible Runner | warning | high | row | An eligible runner row has a gender/sex value that cannot be normalised. | Exclude from scoring and report row for correction. |
| `AUD-ROW-002` | Invalid Time For Eligible Runner | warning | high | row | An eligible runner row has a missing or invalid time. | Exclude from scoring and report row for correction. |
| `AUD-ROW-003` | Missing Category For Eligible Runner | warning | medium | row | An eligible runner has a blank category and is defaulted to `Sen`. | Score runner, but report that category was defaulted. |
| `AUD-ROW-004` | Category Derived From EA 5-Year Band | warning | medium | row | An eligible runner category has been translated from an EA-style 5-year band to the nearest league category because there is race-level evidence that the race used 5-year bands. | Score runner using the derived league category and report the original value alongside the derived category. |
| `AUD-ROW-005` | Missing Runner Name With Result Data | warning | medium | row | A row has a valid time but no usable runner name. | Exclude from scoring and report the row because the missing name suggests incomplete source data rather than a harmless blank line. |
| `AUD-ROW-006` | Unrecognised Club | warning | high | row | A row belongs to a club not mapped into the league and is treated as non-league. | Keep row for race results, exclude from scoring, and include in the unrecognised-club summary with possible fuzzy-match suggestions. |
| `AUD-ROW-008` | Duplicate Eligible Runner In Race | warning | high | row | The same eligible runner appears more than once in a race and one row is removed. | Keep fastest eligible row and report discarded entry. |
| `AUD-ROW-009` | Duplicate Ineligible Runner In Race | info | low | row | The same non-league runner appears more than once in a race. | Keep both rows unchanged and do not treat as an error. |
| `AUD-ROW-010` | Duplicate Runner Attribute Conflict | error | high | row | Duplicate rows for the same runner disagree on club, category, or sex. | Keep the quickest time in the results output, but report the attribute conflict separately in audit because audit must not alter the results-building flow. |
| `AUD-RUNNER-001` | Gender Inconsistency Across Season | error | high | runner | A runner with trusted identity matching appears with conflicting genders across races. Trusted identity for automatic audit means name, club, and sex match; if sex differs, identity is not trusted automatically. | Raise a warning-level identity stop case when sex conflicts on an otherwise candidate match. Reserve this error for manually trusted or otherwise confirmed identity matches only. |
| `AUD-RUNNER-002` | Category Inconsistency Across Season | warning | medium | runner | The same runner appears with materially inconsistent categories across races after allowing for expected birthday progression and race-level EA-style category conversion, using the first chronological race as the baseline category for that runner. | Report sequence of categories and ask for review. |
| `AUD-RUNNER-003` | Club Inconsistency Across Season | error | high | runner | A manually trusted or otherwise confirmed runner identity appears under different eligible clubs across races. | Flag for manual review; may affect points and team scoring. Do not raise this from automated name matching alone. |
| `AUD-RUNNER-004` | Eligibility Inconsistency Across Season | warning | high | runner | The same runner appears eligible in some races and non-league in others. | Surface the warning even when club lookup is unresolved, but mark it as dependent on club-mapping review because audit is iterative and cascading findings should stay visible. |
| `AUD-RUNNER-005` | Suspected Name Variant | warning | medium | runner | Two similar names may refer to the same person but are not confidently merged. | Review only. Show fuzzy-match suggestions, require manual source-data correction, and do not auto-merge audit identities. |
| `AUD-RUNNER-006` | Implausible Time Progression | info | low | runner | Times for a runner look unusually inconsistent across races. | Report for curiosity/review only, not blocking. |
| `AUD-RUNNER-007` | Exact Name Collision Across Clubs | warning | high | runner | The same normalised runner name appears for different clubs, but the identity may represent two different people rather than one runner changing clubs. | Treat as ambiguous identity. Do not auto-merge season records or raise club inconsistency from name match alone; require manual review first. |
| `AUD-RUNNER-008` | Candidate Identity Sex Conflict | warning | high | runner | A candidate same-runner match shares name and club context, but sex differs across records. | Raise a warning and stop automatic identity escalation for that candidate set. Do not continue into gender inconsistency or club inconsistency until manually resolved. |
| `AUD-CLUB-001` | Club Mapping Missing | error | high | club | A club in race data has no mapping to a preferred league club. | Report raw club and likely candidate if available. |
| `AUD-CLUB-002` | Club Division Mismatch | warning | high | club | Club configuration contains inconsistent division information. | Keep first config, flag for admin review. |
| `AUD-CLUB-003` | Club Alias Collision | warning | medium | club | The same raw alias may map ambiguously to more than one preferred club. | Flag for mapping cleanup. |
| `AUD-CLUB-004` | Unrecognised Club Summary With Best Match | warning | medium | club | Aggregate all unrecognised clubs across the selected races and provide the best fuzzy match against known league clubs. | Do not change source data or mappings automatically; show the best single candidate and a confidence score only. |
| `AUD-SEASON-001` | Best-N Boundary Risk | info | medium | season | A data issue changes whether a runner counts in best-N scoring. | Surface as scoring-impactful summary detail. |
| `AUD-SEASON-002` | Team Composition Risk | warning | high | season | A data issue could change A/B team composition or division points. | Highlight separately from general warnings. |

## Proposed First-Cut Audit Views

- `Race Audit Summary`: one row per race with counts by issue type and severity
- `Actionable Issues`: compact high-signal queue for manual correction and follow-up
- `Runner Audit`: one row per runner-level issue across the season
- `Row Audit`: one row per source-row issue for exact correction work
- `Club Audit`: one row per club mapping/configuration issue
- `Unrecognised Club Summary`: retained as a workbook view, but currently deferred as an active population target in this branch
- `Review Issues`: dedicated UI panel backed by `Actionable Issues` with filtering, source-file navigation, runner-history jump, and code-specific quick-fix actions
- `Missing Name Review`: row-audit entries for missing runner names should show race, source row, captured time, and any available club, category, or sex values so the source row can be located quickly. These cases should stay manual only and be corrected in source data, not inside audit.
- `Category Scheme Review`: race-audit entries should classify each race as `League Bands`, `EA 5-Year`, or `External Data Check`. For `EA 5-Year`, row audit should show original and derived categories. For `External Data Check`, race audit should list the distinct female categories that triggered the warning.
- `Duplicate Conflict Review`: duplicate-runner audit should group entries by runner within race, showing the kept row and discarded row references. When club, category, and sex all match, the audit should stay warning-level. When they conflict, the race should surface a high-severity duplicate conflict summary, but results should still keep the quickest time.
- `Runner Identity Review`: exact name matches across different clubs should first be treated as possible identity collisions. Audit should avoid collapsing these records into one season runner until a human confirms they are the same person.
- `Runner Identity Review`: automatic trusted identity for audit should require name, club, and sex. Suspected name variants should be review-only with fuzzy suggestions, and any future approval workflow should mirror the club-matching process rather than auto-correcting records.
- `Audit Workbook Output`: first test drive should emit all current audit issue types into a workbook with one audit view per sheet, and those same views should then be shown on screen. Nothing should be suppressed just because another upstream issue exists; dependency notes can be added, but cascading findings should remain visible.

## Proposed Priorities For v4.0

Build first:

- All current issue types should be emitted in the first audit test drive, including warning- and info-level items.
- Interactive resolution for suspected name variants is deferred, but the issue itself should still be reported.
- `Actionable Issues` sheet and `Review Issues` workflow
- audit run scopes (`single file` and `all discovered files`) with replace-existing control
- staged-checks quality gate and data-quality report output

Defer or make optional later:

- `AUD-CLUB-001` to `AUD-CLUB-004` population and club conversion approval workflows
- `AUD-RUNNER-005`
- `AUD-RUNNER-006`
- `AUD-SEASON-001`
- `AUD-SEASON-002`

## Review Questions

- No open questions remain for the club-conversion, missing-name, category-scheme, duplicate-conflict, runner-identity, or first audit-output workflows.
