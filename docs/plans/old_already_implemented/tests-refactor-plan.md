# Test Suite Audit & Refactoring Plan

**Date:** 2026-04-03 (updated after da0f8db merge: demo actuals seeding)
**Audited:** 200 tests across 23 files (including conftest.py)
**Method:** Testing Anti-Patterns skill — verify real behavior, eliminate waste

---

## CRITICAL: Tautological & Empty Tests (5 tests)

These tests **always pass** regardless of code correctness. They give false confidence and must be fixed or removed immediately.

### tests/test_grouping.py

| Test | Issue | Action |
|------|-------|--------|
| `test_apply_year_groups_default_expands_year1` (L55) | **Zero assertions** — body is just `wb.close()` | **FIX:** Assert year 1 cols not hidden, year 2 cols hidden |
| `test_apply_year_groups_with_current_month` (L66) | **Zero assertions** — body is just `wb.close()` | **FIX:** Assert correct year expanded based on current_month |
| `test_year_groups_applied_to_plan` (L7) | Only checks `dim is not None` — column C always exists after any write | **FIX:** Assert `outlineLevel >= 1` |
| `test_year_groups_applied_to_capacity` (L19) | Same issue | **FIX:** Assert `outlineLevel >= 1` |
| `test_year_groups_applied_to_actuals` (L30) | Same issue | **FIX:** Assert `outlineLevel >= 1` |

### tests/test_styles.py

| Test | Issue | Action |
|------|-------|--------|
| `test_apply_year_groups_basic` (L26) | `assert outlineLevel == 1 or True` — **tautology**, always True | **FIX:** Remove `or True`, assert `outlineLevel == 1` |
| `test_apply_year_groups_with_current_month` (L37) | **Zero assertions** | **FIX:** Assert correct year expanded |

---

## REDUNDANT: Fully Duplicate Tests (3 tests -> remove)

| Test | Duplicated by | Action |
|------|--------------|--------|
| `test_recalc.py::test_recalculation_zero_errors` | `test_e2e.py::test_demo_mode_recalc_zero_errors` | **REMOVE** entire `test_recalc.py` |
| `test_grouping.py::test_apply_year_groups_creates_groups` (L41) | `test_styles.py::test_apply_year_groups_basic` | **REMOVE** from test_grouping.py (keep in test_styles.py after fix) |

---

## DEAD CODE

| Location | Issue | Action |
|----------|-------|--------|
| `test_integration.py:48-54` | `for person_name, group ... pass` loop does nothing | **REMOVE** dead loop |

---

## ANTI-PATTERN: Testing Internals

| Location | Issue | Action |
|----------|-------|--------|
| `test_capacity.py:238` | Imports `_wp_config_row` (private function) | **REFACTOR:** Compute expected config row from layout instead |

---

## NEW CONCERN: Demo Config Coupling in Builders

The recent merge (da0f8db) introduced function-level imports of `DEMO_CURRENT_MONTH` and `DEMO_ACTUALS_MULTIPLIERS` in `actuals.py` **outside** the `if work_packages is None` guard. These imports run even on rebuild path (wasted work, not a bug). The `test_no_module_level_config_import` test still passes because they're function-level, not module-level — but the spirit of decoupling is violated.

| Location | Issue | Action |
|----------|-------|--------|
| `actuals.py:37-39` | Imports `DEMO_CURRENT_MONTH`, `DEMO_ACTUALS_MULTIPLIERS` unconditionally inside function | **FIX:** Move inside `if mode == "demo"` guard |
| `workbook.py:143` | Imports `DEMO_CURRENT_MONTH` inside `_build_all_sheets` | **FIX:** Move inside `if mode == "demo"` guard |

---

## FIXTURE WASTE: Redundant Workbook Generations

Each `generate_workbook()` / `init_workbook()` call takes ~0.1s. Current waste:

| File | Current | Optimal | Savings |
|------|---------|---------|---------|
| `test_protection.py` | 6 workbooks (1 per test) | 1 shared module fixture | 5 generations |
| `test_grouping.py` | 3 workbooks (1 per test) | 1 shared module fixture | 2 generations |
| `test_capacity.py` | 3 workbooks (3 fixtures) | 2 (full_wb + isolated build) | 1 generation |
| `test_workbook.py` | 1 own fixture | Use conftest `fresh_wb_path` | 1 generation |

**Action:** 
- `test_protection.py`: Create a single `module`-scoped fixture, share across all 6 tests
- `test_grouping.py`: Create a single `module`-scoped fixture for the 3 full-workbook tests
- `test_capacity.py`: Eliminate the extra Team build in `test_capacity_wp_row_budget_formula` — derive share columns from the existing capacity_layout
- `test_workbook.py`: Replace local `fresh_path` fixture with conftest's `fresh_wb_path`

---

## CONSOLIDATION: Tests That Should Be Merged

### test_reader_new.py: 3 Config reader tests -> 1

```
test_read_config_sheet_roundtrip      — checks title, months, rate, wp count, wp fields
test_read_config_sheet_params_start_date — checks start date (subset of above)
test_read_config_sheet_wp_names       — checks wp names (subset of above)
```

**Action:** Merge all 3 into `test_read_config_sheet_roundtrip`. Add start_date and wp_name checks there.

### test_grouping.py: 3 full-workbook tests -> 1 parametrized

```
test_year_groups_applied_to_plan
test_year_groups_applied_to_capacity
test_year_groups_applied_to_actuals
```

**Action:** Parametrize on `["Plan", "Capacity", "Actuals"]` with shared fixture.

---

## FRAGILE: Hardcoded Values

| Location | Issue | Action |
|----------|-------|--------|
| `test_trends.py:181` | `trends_ws.cell(32, 1)` — hardcoded row 32 for section header | **FIX:** Scan for "PERSON LOAD" like other tests do |
| `test_trends.py:273-276` | `for r in [40, 41, 42, 43]` — hardcoded helper row numbers | **FIX:** Discover helper rows dynamically (between chart data and chart) |
| `test_trends.py:68` | `test_current_month_default_value` imports `DEMO_CURRENT_MONTH` from config — couples to demo seed constant | **OK** (intentional after da0f8db) |
| `test_conditional_formatting.py:117-125` | `TestSummaryConditionalFormatting` accesses CF rules by positional index (`rules[y]`, `rules[num_years]`) | **FIX:** Match by fill color + formula pattern per dev-workflow.md rule |

---

## WEAK ASSERTION

| Location | Issue | Action |
|----------|-------|--------|
| `test_protection.py:98-109` | `test_protection_password` checks protection exists but doesn't verify password is "grant" | **FIX:** Verify `ws.protection.password` matches expected hash or test unprotect with "grant" |

---

## KEEP AS-IS (High-Value Tests)

These tests are well-written, test real behavior, and should not be changed:

- **test_config.py** (2 tests) — fast seed data integrity check
- **test_config_sheet.py** (5 tests) — clean builder unit tests
- **test_parameterization.py** (7 tests) — architectural guard (AST-based import check)
- **test_rebuild.py** (16 tests) — critical rebuild/preservation/validation
- **test_integration.py** (8 tests, after dead code fix) — cross-sheet formula linkage
- **test_e2e.py** (9 tests) — LibreOffice recalc verification
- **test_plan_v2.py** (4 tests) — Plan↔Capacity formula + font color
- **test_actuals.py** (13 tests) — Actuals formulas + demo seeding + cumulative columns
- **test_merged_cells.py** (5 tests) — merge range correctness
- **test_conditional_formatting.py** (10 tests, after index fix) — CF visual regression
- **test_team.py** (7 tests) — Team builder
- **test_reader_v2.py** (8 tests) — Reader round-trip
- **test_cli.py** (10 tests) — CLI modes + snapshot/restore

---

## Execution Order

### Phase 1: Fix Critical (tautological tests)
1. Fix 5 empty/tautological tests in test_grouping.py and test_styles.py
2. Remove test_recalc.py (redundant)
3. Fix dead code in test_integration.py

### Phase 2: Consolidate Fixtures
4. test_protection.py — shared module fixture
5. test_grouping.py — shared module fixture + parametrize sheets
6. test_workbook.py — use conftest fixture
7. test_capacity.py — eliminate extra Team build

### Phase 3: Clean Up
8. Merge 3 Config reader tests in test_reader_new.py
9. Fix hardcoded rows in test_trends.py
10. Fix CF index-based access in test_conditional_formatting.py
11. Remove private function import in test_capacity.py
12. Fix test_protection_password assertion
13. Guard demo config imports in actuals.py and workbook.py (move inside `if mode == "demo"`)

---

## Expected Outcome

| Metric | Before | After |
|--------|--------|-------|
| Total tests | 200 | ~193 |
| Tautological tests | 5 | 0 |
| Redundant tests | 3 | 0 |
| Workbook generations | ~25 | ~16 |
| Tests with zero assertions | 4 | 0 |
