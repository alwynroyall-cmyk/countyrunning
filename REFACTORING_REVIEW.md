# v9.0 Input/Output Restructuring - Review & Next Improvements

## 📊 Current State (11 Commits Applied)

### Modified Files: 14 total
- `league_scorer/output_layout.py` - OutputPaths dataclass (66 ↓ lines removed)
- `league_scorer/publish/publish.py` - Main publish workflow (86 ↓ lines removed)
- `league_scorer/publish/club_report.py` - Chart generation (32 ↓ lines)
- `league_scorer/report_writer.py` - DOCX generation (42 ↓ lines)
- `league_scorer/main.py` - Race processing (16 ↓ lines)
- `league_scorer/process/main.py` - Modular processing (16 ↓ lines)
- `league_scorer/graphical/qt/dashboard.py` - GUI (61 ↓ lines)
- `league_scorer/graphical/qt/view_autopilot.py` - Report viewer (2 ↓ lines)
- `scripts/run_full_autopilot.py` - CLI wrapper (12 ↓ lines)
- `scripts/autopilot/run_full_autopilot.py` - Core autopilot (12 ↓ lines)
- `scripts/run_staged_checks.py` - Staged checks (4 ↓ lines)
- `scripts/analyse_data_quality.py` - Data quality (7 ↓ lines)
- `tests/unit/test_analyse_data_quality.py` - Test updates (3 ↓ lines)

### Current Output Structure
```
output/
├── audit/
│   ├── workbooks/          # Audit Excel files
│   └── manual-changes/     # Manual edits
├── publish/
│   ├── docx/               # Race cards, league updates
│   ├── docx/club-reports/  # Club reports + embedded charts
│   ├── standings/          # Season standings Excel
│   └── package/            # Published package
├── autopilot/
│   └── runs/               # Staged checks, data quality reports
└── raes/                   # RAES JSON integration
```

---

## 🎯 Areas for Next-Level Improvements

### 1. **Dead Code & Orphaned Functions** ⚠️
**Current Status:** Multiple functions still exist but are no longer called

**Files to Review:**
- `league_scorer/output_layout.py`
  - `export_publish_pdfs()` (lines ~121-135) - Defined but import removed from publish.py
  - `package_publish_artifacts()` - Check if still needed
  
**Action Items:**
- [ ] Verify if `export_publish_pdfs()` can be deleted entirely
- [ ] Check if `package_publish_artifacts()` is actually called anywhere
- [ ] Remove `_move_tree_contents()` if legacy folder migration is complete
- [ ] Remove legacy folder migration code if v8.x archives aren't needed

---

### 2. **Error Handling & Cleanup on Exceptions** 🔧
**Current Status:** Some exception handlers catch but don't clean up properly

**Areas:**
- `publish.py` lines 130-165: Empty try block with PDF comment
- `club_report.py`: Chart generation has fallback but no error logging
- `publish.py`: Warnings appended but not detailed in error cases

**Recommendations:**
- [ ] Replace empty try-except for PDF code with explicit logging
- [ ] Add detailed error messages to club report chart generation failures
- [ ] Create structured error reports instead of generic messages

---

### 3. **Test Coverage Gaps** 🧪
**Current Status:** Integration tests reference removed functionality

**Problem Areas:**
- `tests/integration/test_publish_workflows.py` (lines 131+, 164+, 266+)
  - Tests still check for `.publish_results.json` file existence
  - Tests verify PDF exports that no longer exist
  - Test fixtures create PDF files unnecessarily

**Action Items:**
- [ ] Update test fixtures to remove PDF file creation
- [ ] Update assertions from `.json` to `.md` reports
- [ ] Rename test methods to reflect DOCX-only workflow
- [ ] Add tests for Markdown report generation

---

### 4. **Documentation Drift** 📝
**Current Status:** Several docs still mention removed features

**Files:**
- `ReadMe.txt` (line 41) - Still mentions PDF export workflow
- `documents/AUTOPILOT_OUTPUT_INVENTORY.md` - Lists JSON files no longer generated
- `scripts/run_publish_results.py` - Docstring outdated

**Action Items:**
- [ ] Update `ReadMe.txt` to reflect Markdown-only reports
- [ ] Update `AUTOPILOT_OUTPUT_INVENTORY.md` with v9.0 structure
- [ ] Update CLI script docstrings and help text
- [ ] Add v9.0 migration guide to `INSTALL.md`

---

### 5. **Commented-Out Code & Debug Remnants** 🗑️
**Current Status:** Several files have placeholder comments

**Issues:**
- `publish.py` line 130: `# PDF generation has been removed in v9.0 - only DOCX files are now generated` inside empty try block
- Possible debug logging statements left in place

**Action Items:**
- [ ] Remove empty try-except with comment
- [ ] Clean up all v9.0 migration comments
- [ ] Verify no debug print statements remain
- [ ] Remove doc TODOs related to removed features

---

### 6. **Payload Schema Inconsistency** ⚖️
**Current Status:** Report payloads have inconsistent structure after removals

**Issue Examples:**
- `publish.py` payload includes `club_reports_generated` but might be missing other fields
- Different workflows (autopilot, staged-checks, publish) have different summary structures
- No schema validation

**Recommendations:**
- [ ] Define shared payload TypedDict/Pydantic model
- [ ] Standardize all report payloads to consistent schema
- [ ] Add validation before writing reports
- [ ] Document expected summary fields

---

### 7. **Unused Imports & Parameters** 📦
**Current Status:** Some cleaned up, but review needed for completeness

**Likely Issues:**
- `scripts/run_publish_results.py`: Check if all imports are used
- `league_scorer/publish/publish.py`: Verify all imports necessary
- Dashboard imports: Verify all action handlers are connected

**Action Items:**
- [ ] Run full import analysis (e.g., with `vulture` or `pylint`)
- [ ] Remove unused imports systematically
- [ ] Verify all optional parameters have defaults or are required

---

### 8. **Output Path Redundancy** 🗂️
**Current Status:** Multiple docx paths point to same directory

**Issue:**
```python
publish_docx_race_cards_dir=publish_dir / "docx",
publish_docx_league_updates_dir=publish_dir / "docx",
publish_docx_club_reports_dir=publish_dir / "docx" / "club-reports",
```

First two are identical and might not need separate attributes.

**Considerations:**
- [ ] Evaluate if `publish_docx_race_cards_dir` and `publish_docx_league_updates_dir` can be consolidated
- [ ] Consider flattening structure if separate folders not needed
- [ ] Document reasoning for current structure

---

### 9. **CLI Integration** 🖥️
**Current Status:** CLI scripts simplified but may have usability gaps

**Checks Needed:**
- `scripts/run_publish_results.py`: Verify help text is accurate
- `scripts/run_full_autopilot.py`: Check if all flags still work
- Autopilot GUI integration: Verify progress reporting works

**Action Items:**
- [ ] Test all CLI workflows end-to-end
- [ ] Verify progress messages align with new structure
- [ ] Update help text for removed options
- [ ] Ensure error messages reference correct report paths

---

### 10. **Legacy Migration Code** 🔄
**Current Status:** `sort_existing_output_files()` handles folder migration

**Questions:**
- Is this still needed if v8.x is no longer supported?
- Are there production systems still using old folder structure?
- Should migration code be removed for cleaner codebase?

**Decision Points:**
- [ ] Confirm if legacy migration is still required
- [ ] If yes, add logging to track migrations
- [ ] If no, plan removal for v9.1 after migration period

---

## 📋 Priority Matrix for Next Improvements

| Priority | Item | Effort | Impact | Status |
|----------|------|--------|--------|--------|
| 🔴 HIGH | Remove dead code (`export_publish_pdfs`, etc.) | Low | High | To Do |
| 🔴 HIGH | Update integration tests (PDF → Markdown) | Medium | High | To Do |
| 🔴 HIGH | Fix empty try block in publish.py | Low | Medium | To Do |
| 🟠 MEDIUM | Standardize report payload schemas | Medium | High | To Do |
| 🟠 MEDIUM | Update documentation (README, INSTALL) | Medium | Medium | To Do |
| 🟠 MEDIUM | Run import analysis & cleanup | Low | Low | To Do |
| 🟡 LOW | Consolidate docx path attributes | Low | Low | To Do |
| 🟡 LOW | Decide on legacy migration code | Medium | Low | Decision Needed |

---

## ✅ Completed (This Session)

- ✅ Removed PDF generation code across 6 files
- ✅ Removed JSON generation across 8 files  
- ✅ Removed review-pack generation calls
- ✅ Fixed AttributeError chain (3 commits)
- ✅ Fixed NameError (warnings list)
- ✅ Fixed KeyError (None key handling)
- ✅ Removed undefined variable references
- ✅ Python syntax validated on all files

---

## 🚀 Next Session Recommendations

**Quick Wins (30 min):**
1. Remove empty try block from publish.py line 130
2. Delete `export_publish_pdfs()` function if confirmed unused
3. Run Python import check with vulture

**Medium Tasks (1-2 hours):**
4. Update integration tests (PDF → Markdown assertions)
5. Update ReadMe.txt and INSTALL.md with v9.0 info
6. Standardize report payload schema

**Future Session:**
7. Decide on legacy migration code retention
8. Merge 9.0-Input-Output-Restructure → main after full testing
