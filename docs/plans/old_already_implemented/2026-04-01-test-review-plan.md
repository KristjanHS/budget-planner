# Test Suite Critical Review Plan

**Date:** 2026-04-01
**Current state:** 163 tests, 11 files, ~2K LOC production code
**Goal:** Keep high-value tests, remove duplicates, consolidate the rest

---

## Executive Summary

The test suite has grown organically across batches. It has significant overlap
(same invariants checked 2-3x in different files), cosmetic-heavy tests that
don't catch real bugs, and over-parametrization that inflates count without
adding signal. Proposed reduction: **163 → ~65-75 tests**.

---

## Findings by Category

### 1. DUPLICATE TESTS (same assertion, different file)

| What's checked | Location A (KEEP) | Location B (REMOVE) |
|---|---|---|
| Actuals survive --update | `test_integration::test_roundtrip_actuals_still_preserved` | `test_roundtrip::test_update_preserves_actuals` |
| Empty config → ValueError | `test_integration::test_update_refuses_empty_config_table` | `test_reader_v2::test_read_empty_config_table_raises_error` |
| Snapshot/restore roundtrip | `test_cli::test_cli_restore_overwrites_from_snapshot` | `test_e2e::test_snapshot_restore_roundtrip` |
| Planned FTE has formulas | `test_plan_v2::test_planned_fte_cells_are_formulas_referencing_capacity` | `test_workbook::test_planned_fte_has_formulas` (strict subset) |
| Actual FTE = 0 in fresh wb | `test_plan_v2::test_plan_actual_block_unchanged` | `test_workbook::test_actual_fte_zeros` |
| Seed values correct | `test_integration::test_team_shares_to_capacity_seed_values` | `test_capacity::test_capacity_grant_wp_seed_value` |
| Discover layout keys | used implicitly by all tests | `test_roundtrip::test_discover_layout_finds_all_blocks` (explicit key list check) |
| Snapshot creates file | `test_cli::test_cli_snapshot_creates_copy` | `test_cli::test_cli_snapshot_is_exact_copy` (exact ⊃ exists) |

**Action:** Remove the "Location B" tests. Keep the more thorough version.

### 2. INACTIVE MONTHS — tested 3x

The invariant "inactive months must be empty" is checked in:
- `test_workbook::test_inactive_months_empty` (Plan: planned + actual blocks)
- `test_plan_v2::test_planned_fte_inactive_months_still_empty` (Plan: planned block only)
- `test_capacity::test_capacity_inactive_months_empty` (Capacity sheet)

**Action:** Keep `test_workbook::test_inactive_months_empty` (covers both Plan blocks).
Keep `test_capacity::test_capacity_inactive_months_empty` (different sheet, includes grey fill check).
Remove `test_plan_v2::test_planned_fte_inactive_months_still_empty` (strict subset of test_workbook).

### 3. OVER-PARAMETRIZATION

Several tests parametrize across all 8 WPs when the code path is identical for
each WP. Two representative WPs (one fully-active M1-M60, one partially-active)
would catch the same bugs with 2 tests instead of 8.

| Test | Current params | Proposed |
|---|---|---|
| `test_wp_config_valid` | 8 | 2 (WP1 full-span, WP3 partial) |
| `test_planned_fte_cells_are_formulas_referencing_capacity` | 8 | REMOVE (subset of formula_refs_match) |
| `test_planned_fte_formula_refs_match_capacity_layout` | 8 | 2 |
| `test_planned_fte_cells_have_black_font` | 8 | 2 |
| `test_plan_actual_block_unchanged` | 8 | REMOVE (dup of actual_fte_zeros) |
| `test_actual_fte_zeros` | 8 | 2 |
| `test_inactive_months_empty` (workbook) | 8 | 2 |
| `test_capacity_grant_wp_seed_value` | 12 | REMOVE (dup of integration seed test) |
| `test_capacity_inactive_months_empty` | 9 | 2-3 |

**Savings:** ~50 parametrized tests reduced to ~10-12.

### 4. LOW-VALUE TESTS TO REMOVE

| Test | Why low value |
|---|---|
| `test_capacity_sheet_exists` | Caught by ANY Capacity test. If the sheet doesn't exist, 30 other tests fail. |
| `test_capacity_tab_color_purple` | Cosmetic. Not a functional invariant. |
| `test_team_tab_color_orange` | Cosmetic. Not a functional invariant. |
| `test_config_header_exists` | "WP ID" text in first 20 rows. Any config table test catches this. |
| `test_meta_has_content` | Checks len >= 100. Arbitrary threshold, caught by modification_guide check. |
| `test_planned_block_exists` | "PLANNED" text in header. Any planned formula test catches this. |
| `test_actual_block_exists` | "ACTUAL" text in header. Any actual value test catches this. |
| `test_capacity_returns_layout_dict` | Layout dict tested implicitly by every test using it. |
| `test_discover_layout_finds_all_blocks` | Same — layout used everywhere, explicit key check is redundant. |
| `test_formula_count_positive` | If zero formulas → zero_errors test is meaningless, but the recalc result itself would show this. |

### 5. CONFIG TESTS (test_config.py) — QUESTIONABLE VALUE

All 4 tests validate static Python dictionaries (TEAM, EXTERNAL_COMMITMENTS).
If someone makes a typo in config.py, the workbook generation itself will fail
with a clearer error than "test_team_data_has_required_fields".

**Action:** Keep `test_team_wp_shares_reference_valid_wps` and
`test_external_commitments_persons_in_team` (referential integrity catches
typos early with clear messages). Remove the schema checks (required fields) —
Python would already TypeError on missing keys.

### 6. ANTI-PATTERN: TESTING SEED DATA, NOT BEHAVIOR

`test_capacity_external_config_values_match_seed` checks that Capacity sheet
cells equal the hardcoded values in EXTERNAL_COMMITMENTS. This is testing that
the builder copies config → sheet, which is trivially correct and tested more
meaningfully by the round-trip and integration tests.

**Action:** Remove. The integration tests verify the data flow end-to-end.

### 7. SNAPSHOT/RESTORE OVER-TESTING (test_cli.py)

5 tests for snapshot + 3 for restore. The meaningful behaviors are:
- Snapshot creates a file (1 test)
- Restore recovers from snapshot (1 test)
- Restore without snapshot raises error (1 test)
- Restore creates timestamped backup (1 test)

`test_cli_snapshot_is_exact_copy` is redundant with `test_cli_restore_overwrites_from_snapshot` — if restore works, the snapshot was correct.

**Action:** Remove `test_cli_snapshot_is_exact_copy`.

---

## Proposed Test Structure After Cleanup

### test_workbook.py — Structural Invariants (~12 tests)
- `test_sheets_exist` (all 5 sheets)
- `test_named_ranges` (project_start, fte_rate)
- `test_project_start_date` + `test_fte_rate` → merge into `test_global_params`
- `test_wp_config_valid` (2 params: full-span + partial)
- `test_inactive_months_empty` (2 params)
- `test_actual_fte_zeros` (2 params)
- `test_summary_wp_count`
- `test_meta_has_modification_guide`

### test_plan_v2.py — Plan ↔ Capacity Linkage (~6 tests)
- `test_planned_fte_formula_refs_match_capacity_layout` (2 params)
- `test_planned_fte_cells_have_black_font` (2 params)

### test_team.py — Team Sheet (~5 tests)
- `test_team_sheet_layout_keys`
- `test_team_roster_has_all_persons`
- `test_team_roster_max_fte_values`
- `test_team_roster_cells_are_blue_input`
- `test_team_share_matrix_values_match_config`
- `test_team_share_validation_row_has_sum_formulas`

### test_capacity.py — Capacity Sheet (~10 tests)
- `test_capacity_external_config_block_exists`
- `test_capacity_has_person_group_per_team_member`
- `test_capacity_person_has_grant_wp_rows`
- `test_capacity_person_has_external_project_rows`
- `test_capacity_person_has_total_row`
- `test_capacity_person_has_remaining_row`
- `test_capacity_wp_rows_are_blue_input`
- `test_capacity_total_row_is_sum_formula`
- `test_capacity_remaining_row_formula`
- `test_capacity_inactive_months_empty` (2-3 params)

### test_reader_v2.py — Reader Functions (~7 tests)
- `test_read_plan_config_extracts_wp_definitions`
- `test_read_team_roster_names_and_fte`
- `test_read_team_share_matrix`
- `test_read_capacity_discovers_person_groups`
- `test_read_capacity_classifies_grant_vs_external_rows`
- `test_read_capacity_extracts_monthly_values`
- `test_read_capacity_handles_variable_geometry`

### test_integration.py — Cross-Sheet Data Flow + Round-trip (~6 tests)
- `test_capacity_to_plan_formula_linkage`
- `test_team_shares_to_capacity_seed_values`
- `test_roundtrip_capacity_edits_preserved`
- `test_roundtrip_actuals_still_preserved`
- `test_roundtrip_capacity_external_edits_preserved`
- `test_roundtrip_team_roster_preserved`

### test_cli.py — CLI Modes + Snapshot/Restore (~6 tests)
- `test_cli_default_is_demo_mode`
- `test_cli_skeleton_mode`
- `test_cli_skeleton_has_cell_comments`
- `test_cli_snapshot_creates_copy`
- `test_cli_restore_overwrites_from_snapshot`
- `test_cli_restore_creates_timestamped_backup`
- `test_cli_restore_no_snapshot_raises_error`

### test_config.py — Config Referential Integrity (~2 tests)
- `test_team_wp_shares_reference_valid_wps`
- `test_external_commitments_persons_in_team`

### test_recalc.py — LibreOffice Recalc (~1 test)
- `test_recalculation_zero_errors`

### test_e2e.py — End-to-End with LibreOffice (~4 tests)
- `test_demo_mode_recalc_zero_errors`
- `test_skeleton_mode_recalc_zero_errors`
- `test_demo_mode_formula_count`
- `test_demo_mode_plan_planned_equals_capacity_sums`

### Merged/removed files
- `test_roundtrip.py` → DELETE entirely (all 4 tests absorbed elsewhere)

---

## Execution Plan

### Step 1: Remove pure duplicates
Remove the 8 tests from "Location B" in Section 1.
Delete `test_roundtrip.py` entirely.
**Risk:** None — identical coverage exists elsewhere.

### Step 2: Reduce over-parametrization
Change parametrize decorators from `range(NUM_WPS)` to `[0, 2]` (WP1=full-span, WP3=partial).
Update `_grant_seed_params()` and `_inactive_params()` similarly.
**Risk:** Low — bug in WP5 but not WP1/WP3 is theoretically possible but extremely unlikely since the code iterates identically.

### Step 3: Remove low-value tests
Remove the 10 tests from Section 4.
Remove 2 schema tests from test_config.py (Section 5).
Remove `test_capacity_external_config_values_match_seed` (Section 6).
Remove `test_cli_snapshot_is_exact_copy` (Section 7).
**Risk:** None — all invariants covered by remaining tests.

### Step 4: Minor consolidations
Merge `test_project_start_date` + `test_fte_rate` → `test_global_params`.
Remove `test_read_workbook_returns_all_keys` (integration smell — tests dict shape, not behavior).
**Risk:** None.

### Step 5: Verify
Run full test suite. Confirm all remaining tests pass.
Count should be ~65-75.

---

## Dependency Graph

```
Step 1 ──┐
Step 2 ──┼──→ Step 5 (verify)
Step 3 ──┤
Step 4 ──┘
```

Steps 1-4 can run in parallel (different files, no overlap).
Step 5 must run after all others complete.

---

## What We're NOT Removing (and why)

- **Font color tests** (blue=input, black=formula, green=cross-ref): "Font color convention is sacred" per project rules. Keeping 1 per sheet.
- **Recalc tests**: LibreOffice recalc is the only way to verify formulas actually compute correctly. High value.
- **Round-trip preservation tests**: The core UX promise — user edits survive regeneration.
- **Cross-sheet formula linkage**: The v2 architecture's defining feature.
- **Reader tests**: reader.py is the most complex module (420 LOC). Its tests catch real bugs.
