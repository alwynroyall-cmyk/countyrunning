# Enhancements and Failures - August 26, 2026

## Input Structure Redesign

### Problem Statement
- Too many PDF files generated - they look awful and add clutter
- Too many folders and places to look for data
- Input directory structure is confusing with multiple source locations

### Solution: Simplified Input Structure

#### Three Input Categories Only

1. **Control** - Season configuration files
   - Clubs data
   - Events catalog
   - Configuration settings (e.g., points scoring rules)
   - Any season-level parameters

2. **Events** - Simple single event races
   - One-off races with a single occurrence
   - Single file per event

3. **Series** - Multi-race series events
   - Events consisting of multiple race occurrences
   - Currently: Westbury, Heddington, Lightning Bolt
   - One folder per series type
   - Each race occurrence gets its own input file

### Series File Processing

#### Input Naming Convention
- Format: `Race {n} {Event Name} Series #{occurrence}`
- Example: `Race 3 Heddington 5k Series #2`

#### Output/Consolidation
- Generate one consolidated file per series
- Output naming: `Race {n} {Event Name} Series Round {n}`
- Example: `Race 3 Heddington 5k Series Round 2`
- Replace series occurrence number with "Round" designation
- **Remove previous consolidated file** before creating new one
- **Place consolidated file in Events folder** (same location as other event files, not in Series subfolder)

#### Series Folder Structure
```
inputs/
├── Control/
│   ├── Clubs.xlsx
│   ├── Events.xlsx
│   └── Configuration.xlsx
├── Events/
│   ├── Race 1 Sprint Event.xlsx
│   ├── Race 2 Marathon Event.xlsx
│   └── ...
└── Series/
    ├── Westbury/
    │   ├── Race 3 Westbury 10k Series #1.xlsx
    │   ├── Race 3 Westbury 10k Series #2.xlsx
    │   └── Race 3 Westbury 10k Series #3.xlsx
    ├── Heddington/
    │   ├── Race 5 Heddington 5k Series #1.xlsx
    │   ├── Race 5 Heddington 5k Series #2.xlsx
    │   └── ...
    └── Lightning Bolt/
        ├── Race 6 Lightning Bolt Series #1.xlsx
        └── ...
```

### Removal of PDF Files
- All PDF files to be removed from output
- PDFs look awful and don't add value
- Focus on Excel workbooks and data files only
- Reduces output clutter and storage requirements

### Benefits
- Clear, simple input organization
- Reduced file clutter in outputs
- Easier to locate and manage data
- Streamlined consolidation process for series races

### Series Consolidation Workflow
1. Read all race occurrence files from Series/{SeriesName}/ folder
2. Consolidate results into single file
3. **Delete any previous consolidated file** for that series
4. Save new consolidated file to `inputs/Events/` folder with naming convention: `Race {n} {Event Name} Series Round {occurrence}`
5. Series input files remain in `inputs/Series/{SeriesName}/` for reference/audit trail

### Input File Archiving
**Status:** Already in place

- Archive created **before running any processing**
- Archive location: `inputs/.archive/` (single copy of each file)
- Files archived:
  - All Events files from `inputs/Events/`
  - All consolidated Series files from `inputs/Events/`
- Files NOT archived:
  - Individual Series occurrence files (kept in `inputs/Series/{SeriesName}/` as source reference)
  - Control folder files (assumed stable and maintained separately)
  
**Purpose:** Maintain a baseline snapshot of input state before any processing or consolidation occurs, enabling rollback and audit trail verification

---

## Processing Steps

### Step 1: Audit Process
**Status:** No changes - process remains as today

**Input:** 
- Events files from `inputs/Events/` (includes consolidated Series files)
- Control files from `inputs/Control/`

**Process:**
- Validate and clean race data
- Identify issues and inconsistencies
- Generate audit reports
- **Audited files are recreated from scratch** for each run from the Events and Series source files

**Output:**
- Audited input files placed in `inputs/audited/` folder
- This becomes the **single clean source of truth** for all downstream table generation processing
- Previous audited files are overwritten/replaced on each run

**Important Constraint:**
- Any data modifications or corrections must ONLY be applied to the original source files in `inputs/Events/` or `inputs/Control/`
- The audited versions in `inputs/audited/` are derived outputs and regenerated on each run
- Data corrections are made upstream in source files, then audit is re-run to generate fresh audited files

---

## Summary of Input Definition Changes

✅ **Inputs simplified to 3 categories:**
- Control (season config)
- Events (single races + consolidated series)
- Series (multi-occurrence input files)

✅ **Series consolidation workflow:**
- Individual race files consolidated to Events folder
- Previous consolidated files removed
- Source series files retained for audit trail

✅ **Input file archiving:**
- Single copy of all Events and consolidated Series files archived before processing
- Located in `inputs/.archive/`

✅ **Audit process:**
- Remains unchanged from current implementation
- Recreated from scratch on each run
- Outputs to `inputs/audited/` as single source of truth
- All data corrections made to source files, audit regenerated

---

## Downstream Processing

**Audit and Table Generation Processes:** Remain as currently implemented

**Output File Cleanup:** To follow - removal of unnecessary files (PDFs, etc.) and consolidation of output locations

---

## Output File Cleanup

### JSON Files
- **Remove all JSON files** - not used or needed
- Includes:
  - `autopilot_report.json`
  - `publish_results.json`
  - `data_quality_report.json`
  - `staged_checks_report.json`
  - Any other JSON outputs
  
**Rationale:** No operational value; markdown equivalents provide all needed human-readable information

### Logs Folder
- **Remove `output/logs/` folder entirely** - currently empty and not needed
- No operational logging being generated or used
- Eliminates unnecessary directory structure

### Manifests Folder
- **Remove `output/manifests/` folder entirely** - currently empty and not needed
- Eliminates unnecessary directory structure

### Data Quality Folder
- **Remove `output/quality/` folder entirely**
- Data quality report to be appended to the autopilot run reports instead
- Staged checks report to be appended to the autopilot run reports instead
- Consolidates quality metrics and validation checks into the main run report output
- Eliminates separate directory structure

### RAES Folder
- **Keep `output/raes/` folder** - actively used
- Contains important data transformations and outputs
- Do not remove or consolidate

### PDF Folders
- **Remove `output/publish/pdf/` folder entirely** - includes all subdirectories
  - `output/publish/pdf/`
  - `output/publish/pdf/club-reports/`
- PDFs are not needed; Excel workbooks provide data in usable format
- Eliminates file clutter and storage requirements

### Package Folder
- **Remove `output/publish/package/` folder after each run**
- This is a transitory staging folder used during publish operations
- Delete after publish/export is complete
- Do not retain between runs

### Review Packs Folder
- **Remove `output/publish/review-packs/` folder**
- Data is already contained in other Excel workbooks (standings, results)
- Redundant output; no additional value provided
- Either stop producing review-packs entirely, or remove after data is consolidated into other reports

### Category Review File
- **Remove "Season {year} - Category review" file after run**
- Transitory output used during review process
- Not needed after category review is complete
- Delete once review process is finished

### Season Standings File
- **Keep `Season Standings R{N:02d} {year}.xlsx`** - CRITICAL OUTPUT
- This is the singularly most important file generated by the system
- Contains complete race results, all divisions, all categories
- Primary deliverable for season reporting
- Must be retained and distributed

---

## File Cleanup Summary

**Removed/Deleted:**
- All PDF files and `/pdf/` folders
- All JSON files
- `output/logs/` folder
- `output/manifests/` folder
- `output/quality/` folder (data quality + staged checks)
- `output/publish/review-packs/` folder
- `output/publish/package/` folder (after each run)
- Season Category Review files (after run)

**Kept:**
- `input/.archive/` - baseline snapshot
- `inputs/audited/` - clean source of truth
- `output/audit/` - audit workbooks
- `output/publish/docx/` - distributable reports
- `output/publish/standings/` - Season Standings workbook (CRITICAL)
- `output/raes/` - active transformations
- `output/autopilot/runs/` - run reports with integrated quality/staged check metrics

**Result:** Cleaner file structure, only essential files retained, focus on Excel-based data formats

---

## Content Optimization

### Objective
Maximize the value and clarity of the files we're keeping, ensuring data is presented optimally for end users

### Areas to Review
1. **Season Standings workbook structure and sheet organization**
2. **Audit workbook content and presentation**
3. **Report clarity and usability**
4. **Data completeness and accuracy**
5. **Visual presentation and formatting**

### Next Steps
- Review current content in each critical file
- Identify gaps or unclear presentations
- Optimize sheet layouts, column organization, summaries
- Ensure consistent formatting and naming
- Maximize user-friendliness and data accessibility

---

## Net Impact Analysis of Proposed Changes

### Processing Impact
**IMPROVED** ✅
- Simpler input folder structure (3 categories vs. current multi-level complexity)
- Less file path management and conditional logic
- Faster processing (no PDF generation, no JSON generation)
- Clearer data flow: Events → Audit → Audited → Standings
- Series consolidation logic simplified with single output location

### Code Impact
**REDUCED COMPLEXITY** ✅
- **Remove:** PDF generation modules (~500+ lines eliminated)
- **Remove:** JSON report generation (~200+ lines eliminated)
- **Remove:** Manifest/logging folder management (~100+ lines eliminated)
- **Simplify:** Output writer (fewer file types and destinations to manage)
- **Refactor:** Quality/staged-checks integration into run reports (consolidates reporting pipeline)
- **Net:** 800-1000+ lines of code elimination/simplification

### Input/Output Structure Impact
**MUCH CLEANER** ✅

Current complexity:
```
inputs/
  - events/
  - series/{multiple}/
  - control/
  - raes/
  - (multiple nested levels)

output/
  - audit/
  - publish/docx, pdf, standings, review-packs, package/
  - quality/data-quality, staged-checks/
  - logs/
  - manifests/
  - raes/
  - autopilot/runs/
  (10+ top-level folders, multiple redundancies)
```

Proposed clarity:
```
inputs/
  - Control/
  - Events/
  - Series/
  - .archive/

output/
  - audit/
  - publish/docx, standings/
  - raes/
  - autopilot/runs/ (with integrated quality metrics)
  (4 clean folders, no redundancy)
```

### User Experience Impact
**IMPROVED** ✅
- No confusing PDF files to ignore
- Single clear location for all important data
- Season Standings as clear primary output
- Audit workbooks in obvious location
- No mystery folders (logs, manifests, package)
- Clear file naming conventions

### Maintenance Impact
**REDUCED BURDEN** ✅
- Fewer file types to manage
- Simpler folder structure to debug
- Fewer edge cases in file generation
- Less storage/backup complexity
- Easier to document and explain to new users

### Data Loss Risk
**LOW/NONE** ✅
- No data is lost, just reorganized
- Redundant files removed (review-packs, category review, etc.)
- Transitory files cleaned after use (package folder)
- All essential data retained (audit, standings, reports)

---

## Summary: NET POSITIVE IMPACT

| Dimension | Current | Proposed | Impact |
|-----------|---------|----------|--------|
| Code lines | 1000+ | 0-200 | **800+ reduction** |
| Processing time | Baseline | -10-15% | **Faster** |
| Folder complexity | 10+ folders | 4 folders | **60% reduction** |
| User confusion | High | Low | **Much clearer** |
| Data redundancy | Moderate | None | **Elimination** |
| Maintenance overhead | High | Low | **Significant reduction** |

**Conclusion:** Implementing these changes results in:
- ✅ **Simpler processing** - fewer files to generate, clearer pipeline
- ✅ **Less code** - eliminates 800+ lines of PDF/JSON/manifest handling
- ✅ **Cleaner structure** - 60% reduction in folder complexity
- ✅ **Not more complex** - substantially LESS complex overall

**Recommendation:** HIGH PRIORITY for implementation. These are genuine improvements with minimal downside and significant benefits.

---

## Potential Downsides and Mitigations

### 1. PDF Removal
**Downside:**
- Some users may want printable/shareable PDF formats
- Loss of formatted presentation for external distribution
- Excel files less convenient for email sharing than single PDF

**Mitigations:**
- ✅ Provide export-to-PDF on-demand functionality within Season Standings workbook
- ✅ Create a simple PDF export script for users who need formatted outputs
- ✅ Document that Excel provides filtering/sorting for viewing, reducing PDF need
- ✅ Keep DOCX report generation (which can be printed to PDF by users if needed)
- Impact: Minimal - users who need PDFs can generate them when needed

---

### 2. JSON Removal
**Downside:**
- No structured data for programmatic integrations
- Loss of machine-readable status reports
- Third-party tools/dashboards can't consume JSON output

**Mitigations:**
- ✅ If programmatic access needed, keep one JSON file (autopilot_report.json) for CI/CD pipelines
- ✅ Generate JSON on-demand rather than storing permanently
- ✅ Excel workbooks are quasi-structured and can be read programmatically via openpyxl/xlrd
- ✅ Markdown reports can be parsed for key metrics if needed
- Impact: Low - most users don't consume JSON; can be added back selectively

---

### 3. Series Consolidation Loss of Individual Race Details
**Downside:**
- Individual Series race files deleted/consolidated - can't easily access individual race-only data
- If a specific race within series needs review, must refer back to Series/{SeriesName}/ folder

**Mitigations:**
- ✅ Individual series files retained in `inputs/Series/{SeriesName}/` for audit trail/reference
- ✅ Consolidated file includes all race data with clear separation per race
- ✅ Audit workbooks contain per-race details from consolidation
- ✅ Can re-run consolidation if individual race data needs extraction
- Impact: Low - source files retained; consolidation is deterministic

---

### 4. Single Season Standings File (Critical Output)
**Downside:**
- No backup/redundancy if file gets corrupted
- All eggs in one basket for primary deliverable
- Large file might have performance issues

**Mitigations:**
- ✅ Archive Season Standings file immediately after generation
- ✅ Maintain version history (keep previous rounds' standings)
- ✅ Store backup copy in version control or separate location
- ✅ Implement file validation/integrity checks on generation
- ✅ Use Excel built-in repair/recovery if corruption occurs
- Impact: Manageable - standard file backup practices mitigate risk

---

### 5. Removal of Review-Packs
**Downside:**
- If users were using review-packs for detailed point verification
- Detailed race-by-race breakdown not in separate file

**Mitigations:**
- ✅ Season Standings workbook can include a detailed race breakdown sheet
- ✅ Review-packs data is still accessible in audit workbooks
- ✅ Create detailed verification sheets within main standings if needed
- ✅ Document that detailed data is in Standing sheets, not separate file
- Impact: Low - data not lost, just integrated into other files

---

### 6. Integrated Quality/Staged-Checks in Run Reports
**Downside:**
- Run reports become longer/more complex
- Harder to focus on single aspect (autopilot status vs. quality metrics)
- Report might be harder to navigate

**Mitigations:**
- ✅ Use separate sections/headings within single report (clear structure)
- ✅ Create table of contents in markdown report for easy navigation
- ✅ Keep quality metrics as appendix section (easy to skip if not needed)
- ✅ Use clear visual separators between sections
- ✅ Provide executive summary at top, details below
- Impact: Low - better organization mitigates complexity

---

### 7. Simplified Folder Structure Loss of Organization
**Downside:**
- All quality/staging data now grouped under autopilot runs
- Might be harder to find specific metric files if structure is unclear

**Mitigations:**
- ✅ Consistent, documented folder structure (documented in this file)
- ✅ Clear file naming conventions
- ✅ README/index file in output folder explaining structure
- ✅ Predictable paths for all outputs
- Impact: Low - benefits outweigh organization concerns; documentation key

---

### 8. Loss of Legacy Folder Locations
**Downside:**
- Scripts/tools that reference old paths (e.g., `output/quality/data-quality/`) will break
- External integrations expecting old structure will fail

**Mitigations:**
- ✅ Audit all code references to removed folders and update paths
- ✅ Update any external scripts/automation that read from removed locations
- ✅ Create migration script to move data if backward compatibility needed
- ✅ Version bump to indicate breaking change
- ✅ Document migration path in release notes
- Impact: Medium - manageable through code updates and documentation

---

## Downside Summary & Overall Assessment

| Issue | Severity | Mitigation Effort | Overall Risk |
|-------|----------|-------------------|--------------|
| PDF removal | Low | Low | **Very Low** |
| JSON removal | Very Low | Low | **Very Low** |
| Series detail loss | Low | Low | **Very Low** |
| Single standings file | Medium | Medium | **Low** |
| Review-packs removal | Low | Low | **Very Low** |
| Report complexity | Low | Low | **Very Low** |
| Folder structure | Low | Low | **Very Low** |
| Legacy path breaks | Medium | Medium | **Low** |

---

## Final Recommendation

**Downsides are MANAGEABLE and have clear mitigations.**

No fundamental blockers exist. Implementation strategy:
1. ✅ Document new folder structure clearly
2. ✅ Audit and update all code references to removed paths
3. ✅ Add file validation/backup for critical Season Standings file
4. ✅ Create migration guide for any external integrations
5. ✅ Keep individual series files as audit trail
6. ✅ Implement on-demand PDF export if users request it
7. ✅ Test thoroughly before production rollout

**Risk Level:** **LOW** - Downsides are minor, mitigations straightforward, benefits substantial.

**Go/No-Go:** **GO** - Recommended for implementation with mitigations in place.

---

## Implementation Plan - Phase 1: Output Structure Cleanup

### Critical Files to Modify (Priority Order)

**Priority 1 - Configuration (MASTER):**
- [league_scorer/output_layout.py](league_scorer/output_layout.py) - Lines 10-31, 46-68
  - Remove: `quality_data_dir`, `quality_staged_checks_dir`, `logs_dir`, `manifests_dir`, `publish_review_packs_dir`, `publish_pdf_dir`
  - Keep: `audit_dir`, `publish_docx_dir`, `publish_xlsx_dir`, `autopilot_runs_dir`, `raes_dir`
  - Function: `ensure_output_subdirs()` - remove creation of deleted folders
  - Function: `sort_existing_output_files()` - update logic for removed folders

**Priority 2 - PDF Generation (Remove):**
- [league_scorer/report_writer.py](league_scorer/report_writer.py) - Lines 1144-1150, 1338-1344
  - Remove: `_pdf_conversion_enabled()` method
  - Remove: PDF export code (keep DOCX only)
- [league_scorer/publish/publish.py](league_scorer/publish/publish.py) - Lines 159-176
  - Remove: PDF conversion section in `publish_results()`
  - Keep: DOCX generation
- [league_scorer/publish/club_report.py](league_scorer/publish/club_report.py) - Lines 490-498
  - Remove: PDF conversion logic
  - Keep: DOCX generation
- [league_scorer/graphical/qt/dashboard.py](league_scorer/graphical/qt/dashboard.py) - Lines 1174-1214
  - Remove: `_on_export_published_pdfs()` method
  - Remove: Export PDFs button/menu item

**Priority 3 - JSON Generation (Remove):**
- [scripts/run_full_autopilot.py](scripts/run_full_autopilot.py) - Lines 132-134
  - Remove: JSON report writing (keep Markdown only)
  - Update: `_write_report()` to skip JSON
- [scripts/autopilot/run_full_autopilot.py](scripts/autopilot/run_full_autopilot.py) - Lines 132-134
  - Same changes as above (duplicate file)
- [scripts/run_staged_checks.py](scripts/run_staged_checks.py) - Lines 304-311
  - Remove: JSON report writing
  - Keep: Markdown report
- [scripts/analyse_data_quality.py](scripts/analyse_data_quality.py) - Lines 299-305
  - Remove: JSON report generation
  - Keep: Markdown report
- [league_scorer/publish/publish.py](league_scorer/publish/publish.py) - Lines 70-72
  - Remove: `publish_results.json` generation
  - Keep: `publish_results.md`
- [league_scorer/publish/club_report.py](league_scorer/publish/club_report.py) - Lines 511-513, 542-544
  - Remove: Club report JSON summary
  - Keep: DOCX generation

**Priority 4 - Review-Packs Removal (Remove):**
- [league_scorer/output_layout.py](league_scorer/output_layout.py) - Lines 23, 57, 77, 134
  - Remove: `publish_review_packs_dir` definition
  - Remove: Legacy migration logic for review-packs
- [league_scorer/main.py](league_scorer/main.py) - Lines 285, 290
  - Remove: Review-packs generation calls
- [league_scorer/process/main.py](league_scorer/process/main.py) - Lines 285, 290
  - Remove: Review-packs generation calls
- [league_scorer/output_writer.py](league_scorer/output_writer.py)
  - Remove: Functions that write category/time query review files

**Priority 5 - Quality Folder Cleanup (Remove):**
- [league_scorer/output_layout.py](league_scorer/output_layout.py) - Lines 28, 62, 82
  - Remove: `quality_data_dir`, `quality_staged_checks_dir` definitions
- [league_scorer/graphical/qt/dashboard.py](league_scorer/graphical/qt/dashboard.py) - Lines 1157
  - Remove: View staged checks button/reference
- [league_scorer/graphical/qt/view_autopilot.py](league_scorer/graphical/qt/view_autopilot.py) - Line 188
  - Remove: Staged checks folder display

**Priority 6 - Manifests & Logs Cleanup (Remove):**
- [league_scorer/output_layout.py](league_scorer/output_layout.py) - Lines 31, 65, 85
  - Remove: `manifests_dir`, `logs_dir` definitions
  - Remove: Creation of these folders in `ensure_output_subdirs()`

**Priority 7 - Keep These (NO CHANGES):**
- [league_scorer/raes/raes_service.py](league_scorer/raes/raes_service.py) - RAES JSON output
- [league_scorer/raes/raes_write_service.py](league_scorer/raes/raes_write_service.py) - RAES JSON changes
- [league_scorer/graphical/qt/raes_window.py](league_scorer/graphical/qt/raes_window.py) - RAES GUI
- [league_scorer/series_consolidation.py](league_scorer/series_consolidation.py) - Series consolidation (output path update only)

---

## Implementation Steps

### Step 1: Update output_layout.py (MASTER CONFIG)
1. Remove folder definitions: quality, logs, manifests, review_packs, pdf
2. Remove folder creation logic in `ensure_output_subdirs()`
3. Remove `sort_existing_output_files()` logic for removed folders
4. Remove legacy migration code for moved folders

### Step 2: Remove PDF Generation
1. Remove PDF conversion from report_writer.py
2. Remove PDF conversion from publish.py
3. Remove PDF conversion from club_report.py
4. Remove PDF export GUI button from dashboard.py

### Step 3: Remove JSON Generation
1. Remove JSON from autopilot report generation (keep Markdown)
2. Remove JSON from staged checks (keep Markdown)
3. Remove JSON from data quality (keep Markdown)
4. Remove JSON from publish results (keep Markdown)
5. Keep RAES JSON (still needed)
6. Keep user settings/logs JSON (in home directory, not output)

### Step 4: Remove Review-Packs
1. Remove review-packs folder definition
2. Remove review-packs generation in main.py/process/main.py
3. Remove category/time query review file generation

### Step 5: Remove Quality Folder References
1. Remove quality folder definition
2. Remove GUI references to staged checks viewing

### Step 6: Clean Up Manifests & Logs
1. Remove manifests folder definition
2. Remove logs folder definition

### Step 7: Update All Remaining References
1. Search for all references to removed paths
2. Update import/reference statements
3. Update GUI references
4. Update documentation

### Step 8: Test & Validate
1. Run autopilot process - verify only essential files generated
2. Verify output folder structure matches design
3. Verify Season Standings workbook is created correctly
4. Verify RAES files still work
5. Verify DOCX reports still generated

---

## File Statistics

**Total Files to Modify:** ~25 Python files
**Lines of Code to Remove:** ~800-1000 lines
**New Folder Structure:** 4 top-level folders (audit, publish, autopilot, raes)
**Removed Output Types:** PDF files, JSON reports, quality folder, manifests, logs

---

## Rollback Plan

If issues arise:
1. Revert changes from git (`git reset --hard HEAD~N`)
2. Previous version remains on main branch
3. No production data affected (all in output folder)
4. Input files unchanged (still in inputs/ directory)
