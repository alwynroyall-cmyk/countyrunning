# Autopilot Generated Files - Complete Inventory

**Last Updated:** 2026-08-29

## Quick Summary

The autopilot process generates **500+ files annually** across 10 output categories. This document catalogs what's created so you can identify files to archive or ignore.

---

## Output Categories & Files

### 1. AUDIT FILES
**Location:** `output/audit/workbooks/` and `output/audit/manual-changes/`

Files per season:
- `Season Audit.xlsx` (multi-race) or `Race {N} - Audit.xlsx` (per-race audits)
- `manual_data_audit.xlsx` (if manual corrections exist)

**Size:** Typically 2-5 MB per race audit
**Keep?** YES - needed for issue tracking and manual review records

---

### 2. INPUT PROCESSING FILES  
**Location:** `input/{year}/inputs/audited/`

Files per race:
- `Race {N} - Audited.xlsx` - Cleaned input data
- `Race {N} - Consolidated.xlsx` - Merged series data
- `.archive/` subdirectory - Backup of raw inputs

**Size:** 1-3 MB per race
**Keep?** YES - base data for scoring pipeline

---

### 3. DATA QUALITY REPORTS
**Location:** `output/quality/data-quality/year-{year}/`

Files per season:
- `data_quality_report.json` - Machine-readable metrics
- `data_quality_report.md` - Human-readable summary

**Contents:** Blank cell %, invalid time formats, schema warnings, data hotspots
**Size:** 100-500 KB
**Keep?** OPTIONAL - historical records only, can be archived after season ends

---

### 4. STAGED CHECKS REPORTS
**Location:** `output/quality/staged-checks/year-{year}/`

Files per season:
- `staged_checks_report.json` - Validation results per stage
- `staged_checks_report.md` - Human-readable pass/fail

**Contents:** Raw ingest, consolidation, audit, scoring regression status
**Size:** 50-200 KB
**Keep?** OPTIONAL - debugging only

---

### 5. SCORING & RESULTS WORKBOOKS
**Location:** `output/publish/standings/` and `output/publish/review-packs/`

Files per season:
- `Season Standings R{N:02d} {year}.xlsx` - Main results workbook (1 per race)
- `{race_name}_review_pack.xlsx` - Detailed point breakdown
- `Category Mismatch TODO.xlsx` - Flagged issues
- `Time Query TODO.xlsx` - Flagged issues

**Size:** 2-10 MB per race workbook
**Keep?** YES - needed for category/time reviews

---

### 6. RACE SCORING CARDS (DOCX/PDF)
**Location:** `output/publish/docx/` and `output/publish/pdf/`

Files per race:
- `Race {N} Scoring Card {year}.docx` - Issue summary + results
- `Race {N} Scoring Card {year}.pdf` - PDF version
- `League Update R{N:02d} {year}.docx` - Cumulative standings
- `League Update R{N:02d} {year}.pdf` - PDF version

**Size:** 500 KB - 2 MB each
**Keep?** YES if distributed; OPTIONAL if just for internal review

---

### 7. CLUB REPORTS
**Location:** `output/publish/docx/club-reports/` and `output/publish/pdf/club-reports/`

Files per club (typically 30-50 clubs):
- `{ClubName} Season Report {year}.docx` - Team breakdowns + runner analytics
- `{ClubName} Season Report {year}.pdf` - PDF version
- Auto-generated pie charts for gender/category distribution

**Size:** 200 KB - 1 MB per club report
**Total:** 6-50 MB per season (all clubs)
**Keep?** YES if sent to clubs; OPTIONAL if just internal use

---

### 8. PUBLISH PACKAGES
**Location:** `output/publish/package/`

Files per season:
- Flattened or hierarchical bundle of all PDFs and workbooks
- Optional ZIP file `publish_artifacts_{year}.zip`

**Size:** 50-200 MB per season (depending on club reports included)
**Keep?** NO - temporary staging; keep only final exports

---

### 9. AUTOPILOT EXECUTION REPORTS
**Location:** `output/autopilot/runs/year-{year}/`

Files per autopilot run:
- `autopilot_report.json` - Structured execution status
- `autopilot_report.md` - Markdown version
- Multiple timestamped variants if run multiple times

**Contents:** Success/fail, fixes applied, audit snapshots, staged check results
**Size:** 10-50 KB each
**Keep?** YES - execution history and metrics

---

### 10. PUBLISH STATUS REPORTS
**Location:** `output/autopilot/runs/year-{year}/`

Files per publish operation:
- `publish_results.json` - PDF conversion, club report generation status
- `publish_results.md` - Markdown version

**Size:** 5-20 KB each
**Keep?** OPTIONAL - debugging only

---

### 11. LOGS
**Location:** `output/logs/`

Files per session:
- Various `.log` files with debug info

**Size:** 100 KB - 10 MB depending on verbosity
**Keep?** OPTIONAL - debugging only, can be purged quarterly

---

---

## FILE TYPE REFERENCE

| Extension | Quantity | Purpose | Priority |
|-----------|----------|---------|----------|
| `.xlsx` | 4-10 per race | Audit, results, review packs, TODO items | KEEP |
| `.docx` | 2+ per race + 1+ per club | Scoring cards, league updates, club reports | KEEP or ARCHIVE |
| `.pdf` | 2+ per race + 1+ per club | PDF versions of reports | KEEP or ARCHIVE |
| `.json` | 4-8 per season | Status & metrics reports | OPTIONAL |
| `.md` | 4-8 per season | Human-readable summaries | OPTIONAL |
| `.log` | Variable (10-50+) | Application logs | OPTIONAL |

---

## Recommended Cleanup Strategy

### KEEP (Active Work)
```
output/audit/                          # All audit workbooks
output/publish/standings/              # Results workbooks
output/publish/review-packs/           # Category/time query files
output/publish/docx/                   # Reports being reviewed/distributed
output/publish/pdf/                    # Final published PDFs
output/autopilot/runs/year-{year}/     # Recent autopilot reports
```

### ARCHIVE (End of Season)
```
output/publish/docx/club-reports/      # After sent to clubs
output/publish/pdf/club-reports/       # After sent to clubs
input/{year}/inputs/audited/           # After season complete
```

### DELETE (Safe to Remove)
```
output/publish/package/                # Temporary staging (keep ZIP only)
output/quality/                        # After season ends (historical only)
output/logs/                           # Older than 3 months
```

### IGNORE (Development/Temporary)
```
output/autopilot_runs/                 # Legacy directory
output/staged_reports/                 # Temporary processing
output/staged-checks/                  # Temporary processing
quarantine and copy/                   # Legacy files
```

---

## Estimated Annual Output Volume

- **Audit files:** 5-20 MB (1-4 races)
- **Input audited files:** 10-30 MB
- **Data quality reports:** 1-2 MB
- **Results workbooks:** 20-40 MB
- **Scoring cards (DOCX):** 10-20 MB
- **Scoring cards (PDF):** 15-30 MB
- **Club reports (DOCX):** 10-50 MB
- **Club reports (PDF):** 15-75 MB
- **Status/metrics JSON:** 100 KB
- **Logs:** 100 MB - 1 GB (depending on verbosity)

**Total:** 150 MB - 1.5 GB per season (including logs)

---

## Automation Tips

To reduce clutter, consider:

1. **Archive old seasons** - Move `output/year-{N}` directories to cold storage after 2 years
2. **Auto-cleanup logs** - Delete `.log` files older than 90 days
3. **Compress club reports** - After season, ZIP all club report PDFs for archival
4. **Remove staging** - Regularly clean `output/publish/package/` intermediate files

---

## References

- **Main autopilot script:** `scripts/run_full_autopilot.py`
- **Output writer:** `league_scorer/output_writer.py`
- **Output layout:** `league_scorer/output_layout.py`
- **Report writer:** `league_scorer/report_writer.py`
