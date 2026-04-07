# WRRL Staged Checks Report

Generated: 2026-04-07T16:34:58.427505Z
Success: False

## Stage 1 - Raw Ingest Validation
Status: passed

Validated 8 raw workbook(s); failures=0, warnings=0.

## Stage 2 - Raw To Audited Consolidation
Status: passed

Raw races=8, audited races=8; quality gate=80.0% (observed=99.78%).

### Stage 2 Quality Gate Details

- Threshold: 80.0%
- Observed Success: 99.78%
- Gate Passed: True
- Audited Blank Category %: 0.0
- Audited Invalid Time %: 0.88
- Data Quality Report: output\quality\data-quality\year-2025\data_quality_report.md

### Suggested Fixes

- Normalize common time formatting variants in source files before parsing (for example stray spaces, punctuation, and hh:mm:ss variants).

## Stage 3 - Audit Generation Validation
Status: passed

Audit workbook generated with 25 actionable issue row(s).

## Stage 4 - Main Scoring Regression
Status: failed

Results fingerprint differs from baseline.
