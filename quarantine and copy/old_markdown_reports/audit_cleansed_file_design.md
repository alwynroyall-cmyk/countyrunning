# Audit Cleansed File Design

This file captures the agreed design for the later audit-to-cleansed-file stage.

## Core Principle

- Audit remains the first step.
- Cleansing is only introduced after the audit output is trusted.
- The cleansed file is a new derived file and must never overwrite the source input file.

## File Location And Naming

- Cleansed files must be written into the same `inputs` folder as the original race file.
- Cleansed files must sit next to the original source file.
- Default file name format: `Race N - Event Name (audited).xlsx`
- If the default audited file name already exists, do not overwrite it.
- In that case, prompt the user for a different name using the standard rename or save dialog.
- After a successful write, confirm clearly that the audited file has been created.

## Write Behaviour

- Never overwrite the original input file.
- Never silently overwrite an existing audited file.
- Write the audited file automatically when requested; do not add an extra approval prompt unless the target name already exists.
- If the target name already exists, require the user to choose a new file name.

## Workbook Shape

- Keep the workbook shape as close to the source race file as possible.
- Preserve the original rows and their order.
- Preserve the original columns wherever possible.
- Add normalised and cleansed values into the workbook rather than replacing source values invisibly.

## Cleansed Content

The audited file should contain, as far as the audit has been able to determine:

- normalised club
- normalised sex
- normalised category
- normalised eligibility state
- any other cleaned values the audit has high confidence in

## Comments Column

- Add a new column labelled `Comments`.
- Use `Comments` to explain what was changed, defaulted, derived, or left unresolved.
- If multiple audit notes apply to the same row, combine them in the `Comments` column in a readable form.

Examples of comment content:

- `Club normalised from 'Avon Vallley' to 'Avon Valley Runners'`
- `Category derived from 'V35' to 'Sen' using EA 5-year audit rule`
- `Missing category defaulted to 'Sen'`
- `Unresolved identity issue remains for manual review`

## Safety Boundary

- Only write cleansed values where the audit result is confident enough for automatic derivation.
- Keep unresolved or manual-review issues visible in `Comments`.
- Do not pretend unresolved issues have been fixed.

## Relationship To Scoring

- The scoring process will continue unchanged at first.
- Once the audited-file workflow is proven, scoring can be simplified later to rely more on the audited input.
- Scoring cleanup is a later phase and is intentionally separate from the first audited-file implementation.