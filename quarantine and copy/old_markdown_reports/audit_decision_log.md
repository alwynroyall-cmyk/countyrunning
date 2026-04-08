# Audit Decision Log

This file records agreed audit decisions as they are made.

## 2026-04-01 - Initial Club Audit Decisions

### Decision 1

Blank club values are not audit issues by themselves.

Reason:
- not every runner is affiliated
- absence of a club name is not automatically bad data
- we should not create noise from valid unaffiliated entries

Effect:
- remove blank club as a standalone audit issue
- do not flag missing club merely because the club cell is empty

### Decision 2

Unrecognised clubs should remain audit-visible.

Reason:
- some raw club names will be genuine non-league affiliations
- some will be misspellings or unmapped league clubs
- this is still important because it can affect eligibility and scoring

Effect:
- keep unrecognised club as an audit issue
- keep source data unchanged
- do not auto-correct mappings during audit

### Decision 3

Audit should include an aggregate unrecognised-club summary across all selected races.

Reason:
- this is one of the most labour-intensive manual cleanup tasks
- a cross-race view is more useful than isolated row-by-row warnings

Effect:
- one audit view should group unrecognised raw clubs across all races
- include occurrence count and races seen

### Decision 4

Unrecognised-club audit should offer fuzzy-match suggestions, but suggestions must never change data automatically.

Reason:
- suggestions can reduce manual review effort
- automatic mapping would risk silent data corruption

Effect:
- show the best single likely known-club candidate only
- include a confidence score where possible
- keep decisions human-controlled

### Decision 5

Rows with missing runner names should only become audit issues when there is evidence that a real result row exists.

Reason:
- empty lines and formatting artefacts are common in race workbooks
- a missing name matters when the row still carries a valid time

Effect:
- a valid time is enough evidence on its own
- audit and report rows that have a usable time but no runner name
- do not create noise from harmless blank rows

### Decision 6

Category inconsistencies should be handled differently depending on whether the source race is clearly using EA 5-year bands.

Reason:
- some category differences are normal birthday progression and should remain warnings, not hard failures
- some races use EA 5-year bands, which can often be translated to league categories when the scheme is clearly present
- mixed schemes or uncertain mappings need manual verification rather than silent remapping

Effect:
- treat expected birthday progression as warning-level review, not an error by default
- after normalising categories, scan female runner categories for race-level evidence of the scheme in use
- if female categories are all standard league bands such as `V40` and `V50`, treat the scheme as already correct
- if female categories show EA-style 5-year progression such as `V35`, `V40`, `V45`, derive the nearest league category and report the derivation
- if female categories show a mixed or ambiguous pattern such as `V35`, `V45`, `V55`, flag `External Data Check` instead of forcing a conversion
- there is no minimum threshold for EA detection; even one female category entry that proves EA usage is enough to classify the race as EA-style
- for cross-season runner category review, use the first chronological race as the baseline category for that runner

### Decision 7

Duplicate runner rows should consolidate to the quickest time, but club, category, and sex are expected to match.

Reason:
- duplicate result lines often represent the same performance duplicated in export or entry data
- conflicting runner attributes suggest a source-data problem that should not be hidden

Effect:
- keep the quickest time in the normal results-building flow
- report attribute conflicts when duplicate rows disagree on club, category, or sex
- keep audit reporting separate so audit does not affect the results output chosen for scoring

### Decision 8

PDF export failures are process issues, not audit issues.

Reason:
- they describe output-generation state rather than source-data quality
- they still matter operationally, but they do not belong in the audit data model

Effect:
- keep PDF export warnings visible in run/output status only
- remove PDF export failures from the audit register

### Decision 9

If the best-match club suggestion is accepted by a human, it should be possible to add that raw-club to preferred-club pairing into the club mapping file.

Reason:
- this turns repetitive manual review into a durable rules improvement
- the mapping should still remain a deliberate user action

Effect:
- keep fuzzy suggestion generation read-only
- design the audit workflow with a popout review box showing `Tick`, `Current Club`, `Proposed Club`, and `Message`
- label the action button `Add conversions to Club Lookup`
- include confidence, occurrence count, and races seen within the `Message` field rather than adding separate columns
- start with all ticks cleared so adding a club conversion is always an active user choice
- keep the popout fully manual with no bulk actions such as `Select All` or `Select High Confidence`
- show a confirmation summary of the selected mappings before writing them to the club lookup file
- warn in that confirmation summary when a selected raw club already exists in the club lookup file
- distinguish exact existing matches from conflicting existing matches in that warning
- when a conflicting existing match is shown, allow the user to return and untick it before proceeding
- provide a quick return path back to the selection popout for that untick step
- reopen the same selection state when returning so the user's existing choices are preserved
- highlight the conflicting selected rows when returning so the user can identify and untick them quickly
- allow approved mappings to be written into the club mapping file from that review step

### Decision 10

Missing-name audit should focus on fast source-row recovery rather than in-audit correction.

Reason:
- these rows usually need source workbook correction, not derived-system fixes
- the reviewer needs enough context to locate the bad row quickly without opening multiple views

Effect:
- row-audit entries for missing names should show race, source row, captured time, and any available club, category, or sex values
- blank spacer rows should continue to stay out of the formal audit
- the audit UI should not try to repair missing names directly

### Decision 11

Category-scheme review should classify races explicitly and suppress avoidable downstream noise.

Reason:
- category problems are easiest to review at race level first, before looking at runner-level consistency
- races flagged as uncertain should not create misleading runner-level category inconsistency noise

Effect:
- each race should be classified in audit as `League Bands`, `EA 5-Year`, or `External Data Check`
- `EA 5-Year` races should show original and derived categories at row level
- `External Data Check` races should list the distinct female categories that triggered the warning
- runner-level category inconsistency review should stay visible even when a race is flagged as `External Data Check`, but it should be marked as dependent on the race-level category issue

### Decision 12

Duplicate-runner audit should separate straightforward deduplication from true attribute conflicts.

Reason:
- some duplicates are routine export noise and only need traceability
- conflicting duplicate attributes are materially different and deserve stronger visibility

Effect:
- duplicate audit should group entries by runner within race
- show kept row and discarded row references for duplicate handling
- when club, category, and sex all match, keep the quickest time and report a warning only
- when those attributes conflict, surface a high-severity duplicate conflict summary while still keeping the quickest time in results

### Decision 13

Exact runner-name matches across different clubs should be treated as ambiguous identity unless there is stronger evidence that they are the same person.

Reason:
- the same name can legitimately belong to two different runners in different clubs
- treating name match alone as identity truth would create false club-change audit findings and potentially bad season merges

Effect:
- do not raise club inconsistency from exact name match alone when clubs differ
- raise an identity-collision warning instead and hold those records apart until reviewed
- only raise `Club Inconsistency Across Season` where identity is trusted beyond a simple exact name match

### Decision 14

Automatic trusted identity for audit should require name, club, and sex.

Reason:
- this is the safest rule available for a first audit test drive
- weaker matching would create false positives across common names and club changes

Effect:
- automatic audit should only trust same-runner identity when name, club, and sex align
- if sex differs inside a candidate match, raise a warning and stop automatic escalation for that candidate
- `Club Inconsistency Across Season` should require manual or stronger identity confirmation beyond automated matching

### Decision 15

Suspected name variants should be review-only with fuzzy suggestions and no automatic correction.

Reason:
- name-variant handling is useful, but it is also one of the easiest ways to create wrong merges
- this needs a careful workflow similar to club matching, not a quick heuristic patch

Effect:
- report suspected name variants with fuzzy suggestions only
- require manual correction in source data
- defer any interactive approval workflow for this area into a separate future task

### Decision 16

Eligibility inconsistency should remain visible even when unresolved club lookup may be the underlying cause.

Reason:
- the audit process is iterative and the user wants cascading findings to remain visible
- hiding downstream issues would slow down review for someone who is comfortable with layered error output

Effect:
- emit eligibility inconsistency warnings even when club mapping is unresolved
- mark those warnings as dependent on club lookup review where appropriate
- do not suppress downstream findings solely because an upstream issue exists

### Decision 17

The first audit test drive should emit everything and present it in a workbook first, then on screen.

Reason:
- the fastest way to learn from a first audit pass is to expose the whole shape of the data problems
- tuning noise can happen after there is real output to inspect

Effect:
- emit all current issue types in the first audit pass, including informational items
- write audit output into a workbook with one audit view per sheet
- show those same audit views on screen after generation

### Decision 18

The future workflow for suspected name variants should mirror the club-matching approach, but it is intentionally deferred.

Reason:
- there is value in a guided approval flow here, but it is not needed to test-drive the audit rules
- deferring it keeps the current scope tight while preserving the design intent

Effect:
- add suspected-name-variant workflow to a todo list
- keep current implementation scope to reporting and review only

### Decision 19

The first audit output should use a fixed workbook-and-screen layout rather than an open-ended export.

Reason:
- the first test drive will be judged on whether the output is reviewable, not just whether the rules exist
- fixed views make it easier to compare runs and refine noise later

Effect:
- emit workbook sheets for `Race Audit Summary`, `Actionable Issues`, `Row Audit`, `Runner Audit`, `Club Audit`, `Unrecognised Club Summary`, `Candidates To Check`, and `EA Checked`
- use `Actionable Issues` as the data source for the on-screen `Review Issues` workflow
- mirror workbook sheets in the `View Audit` screen after writing the workbook
- use the column definitions captured in `audit_output_design.md`

### Decision 20

Audit output naming and location should be predictable from the user’s selection.

Reason:
- the user wants selection-driven input and obvious output placement
- audit files need to be easy to locate beside other run outputs

Effect:
- write audit workbooks into `outputs/audit`
- when exactly one race is selected, name the workbook `Race N - Audit.xlsx`
- when multiple races are selected, name the workbook `Season Audit.xlsx`

### Decision 21

Dependency and status wording should be short, fixed, and visible on every audit sheet.

Reason:
- cascading findings are acceptable, but reviewers still need a quick way to see what should be tackled first
- controlled wording makes screen filtering and workbook sorting practical

Effect:
- use status values `Open`, `Dependent`, `Manual Review`, `Ready To Fix`, and `Informational`
- use dependency values `None`, `Club Lookup`, `External Category Check`, `Identity Review`, `Source Data Correction`, and `Duplicate Conflict Review`
- include `Status`, `Depends On`, and `Next Step` columns in the main audit sheets

### Decision 22

Once the audit output is trusted, audit should be able to write a cleansed race file automatically as a new derived workbook.

Reason:
- this creates a controlled bridge between audit review and a cleaner scoring input
- the cleansed file should reduce repeated manual fixes and later allow the scoring process to be simplified

Effect:
- write a new audited workbook beside the original source file in the same `inputs` folder
- use the naming pattern `Race N - Event Name (audited).xlsx`
- never overwrite the original source file under any circumstances
- if the audited filename already exists, require a new name via the standard rename or save dialog
- confirm clearly after the audited file has been written

### Decision 23

The audited file should preserve the source workbook shape while adding cleaned values and a visible comments trail.

Reason:
- reviewers need to compare the audited file to the original without losing the source context
- silent replacement would make later investigation harder

Effect:
- keep the original row order and source structure where practical
- write normalised and cleansed values as far as the audit can determine them safely
- add a `Comments` column to describe changes, defaults, derivations, and unresolved issues
- keep unresolved issues visible rather than pretending they were fixed

### Decision 24

Scoring cleanup is a later phase and should not be rushed into the first audited-file implementation.

Reason:
- the audited-file workflow needs to prove itself before the scorer is simplified around it
- changing both the audit and scoring boundary at once would make regression analysis harder

Effect:
- build the audited-file workflow first
- keep the current scoring process in place for now
- revisit scoring simplification only after the audited-file output has been proven in practice

## Open Questions

- No open questions remain for the club-conversion, missing-name, category-scheme, duplicate-conflict, runner-name collision, first audit-output, or audited-file workflows.