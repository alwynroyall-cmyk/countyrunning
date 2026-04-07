# WRRL Data Quality Analysis

Generated: 2026-04-07T16:34:54.790812+00:00
Season: 2025
Data Root: C:\Users\alwyn\OneDrive\Documents\WiltshireAthletics\WRRL

## Summary

| Stage | Files | Rows | Blank Category % | Invalid Time % | Blank Name % | Blank Gender % |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Raw | 8 | 4571 | 0.0 | 0.88 | 0.0 | 0.0 |
| Audited Inputs | 8 | 4563 | 0.0 | 0.88 | 0.0 | 0.02 |

## Top Blank Category Hotspots (Audited Inputs)

No blank-category hotspots detected.

## Interpretation

- Prioritize files with highest blank-category percentage at source/audited-input stage.
- Re-run the staged checks after source cleanups to validate downstream impact.

## Suggested Fixes

- Pre-clean time values before parse (trim whitespace, standardize separators, and handle hh:mm:ss / mm:ss variants).
