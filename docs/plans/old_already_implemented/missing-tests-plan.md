# Missing Integration & E2E Tests Plan

**Date:** 2026-04-02
**Baseline:** 64 tests passing across 13 test files (v3 workbook structure)

## Methodology

Compared every feature in each source module against existing test coverage.
Focused on cross-module data flows and full-workflow scenarios that unit tests miss.

---

## GAP 1: Conditional Formatting Rules (Integration)

**What's untested:** Capacity, Summary, Plan, and Actuals all register conditional formatting rules. Zero tests verify these exist or have correct formulas.

**Risk:** A typo in a conditional formatting formula silently breaks overcommitment/budget alerts. Auditors rely on these visual signals.

**Concrete features to test:**

| Sheet | Rule | Condition |
|-------|------|-----------|
| Capacity | TOTAL row amber fill | `>0.8 * Team!$B${roster_row}` |
| Capacity | TOTAL row red fill | `> Team!$B${roster_row}` |
| Capacity | REMAINING row red fill | `< 0` |
| Summary | Year variance red fill | `< 0` |
| Summary | Project total variance red fill | `< 0` |
| Summary | Remaining amber fill | `< ceiling * 0.1` |
| Summary | Remaining red fill | `< 0` |
| Plan | Planned variance yellow fill | `ABS() > 0.01` |
| Actuals | Variance red fill | `< 0` |

**File:** `tests/test_conditional_formatting.py`
**Tests:** ~9 parameterized tests, one per rule above
**Approach:** Read `ws.conditional_formatting` list, verify rule count and formula patterns.

---

## GAP 2: Summary ← Plan + Actuals Cross-Sheet Formulas (Integration)

**What's untested:** Summary sheet formulas reference both Plan (planned FTE) and Actuals (actual EUR). No test verifies these formulas are wired to the correct rows/columns.

**Risk:** After the v3 redesign that added Actuals integration, Summary formulas could reference wrong rows if layout shifts.

**Concrete tests:**

1. **Summary year Plan FTE formulas reference Plan planned block rows** — for each WP, verify `=SUM(Plan!{first_month_col}{planned_row}:{last_month_col}{planned_row})` pattern
2. **Summary year Actual EUR formulas reference Actuals WP rows** — for each WP, verify `=SUM(Actuals!{col}{row}:{col}{row})` pattern
3. **Summary Plan EUR = Plan FTE * fte_rate** — verify formula structure
4. **Summary Variance EUR = Plan EUR - Actual EUR** — verify formula structure
5. **Summary project totals sum across year columns** — verify SUM formulas
6. **Budget Overview section references** — Personnel Actual refs Actuals total, Personnel Ceiling refs `partner_ceiling`, Remaining = Ceiling - Actual, % Consumed formula

**File:** `tests/test_integration.py` (extend existing)
**Tests:** ~6 new tests

---

## GAP 3: Capacity Validation Column Formulas (Integration)

**What's untested:** Each Capacity WP row has Sum/Budget/Variance validation columns. No test verifies:
- Sum formula covers the correct month range
- Budget formula references the correct Team share cell × Plan config FTE
- Variance = Budget - Sum

**Risk:** Validation columns are the user's primary tool for catching allocation errors. Wrong formulas would show misleading guidance.

**Concrete tests:**

1. **Capacity WP row Sum formula spans MONTH_COL_START:MONTH_COL_END** — verify SUM range
2. **Capacity WP row Budget formula = Team!{share_col}{share_row} * Plan!C{config_row}** — verify cross-sheet ref
3. **Capacity WP row Variance formula = Budget - Sum** — verify structure
4. **External row Sum/Budget/Variance formulas** — same pattern but Budget refs config block

**File:** `tests/test_capacity.py` (extend existing)
**Tests:** ~4 new tests

---

## GAP 4: Freeze Panes (Integration)

**What's untested:** All 5 data sheets set `freeze_panes`. Not tested anywhere.

**Risk:** Low — cosmetic — but freeze panes are critical UX for a 60-month grid.

**Concrete tests:**

1. **Plan freeze panes at `C{planned_first_wp_row}`**
2. **Capacity freeze panes at `C{first_person_header_row + 1}`**
3. **Actuals freeze panes at `C{first_wp_row}`**
4. **Summary freeze panes at `C{first_wp_row}`**

**File:** `tests/test_workbook.py` (extend existing)
**Tests:** 4 parameterized

---

## GAP 5: Merged Cell Ranges (Integration)

**What's untested:** Year headers in Capacity, Plan, and Actuals merge cells across 12 columns per year. Summary merges year group headers across 4 columns. Section headers merge across columns. No test verifies merge correctness.

**Risk:** Medium — merged cells that span the wrong range cause confusing layout or misaligned data.

**Concrete tests:**

1. **Plan year headers merge exactly 12 cells (or fewer for final year)** — check `ws.merged_cells.ranges`
2. **Capacity year headers merge correctly** — same pattern
3. **Summary year headers merge across 4 columns per year group**
4. **Section headers (CONFIGURATION, PLANNED FTE, etc.) span correct column range**

**File:** `tests/test_merged_cells.py`
**Tests:** ~4 tests

---

## GAP 6: Reader Round-Trip for External Commitments (Integration)

**What's untested:** `read_capacity_sheet()` extracts the external commitments config block, but no test verifies the extracted data matches `EXTERNAL_COMMITMENTS` from config.

**Risk:** If config block layout changes, the reader silently returns wrong data.

**Concrete tests:**

1. **read_capacity_sheet external_commitments matches config** — verify person, project, fte, start, end for each entry

**File:** `tests/test_reader_v2.py` (extend existing)
**Tests:** 1 new test

---

## GAP 7: Summary Values After Recalc (E2E)

**What's untested:** After LibreOffice recalc, no test verifies Summary sheet computed values. Currently only Plan←Capacity is verified.

**Risk:** Summary could show wrong annual totals — the primary reporting surface for project managers.

**Concrete tests:**

1. **Summary Year 1 Plan FTE > 0 after recalc** — loaded with `data_only=True`
2. **Summary project total Plan FTE equals sum of yearly Plan FTE** (consistency check)
3. **Summary Actual EUR values are 0** (fresh workbook has no actuals)
4. **Budget Overview: Personnel Actual = 0, Personnel Ceiling > 0** after recalc

**File:** `tests/test_e2e.py` (extend existing)
**Tests:** ~4 new tests (all require LibreOffice, auto-skip)

---

## GAP 8: Actuals Sheet Post-Recalc (E2E)

**What's untested:** After recalc, Actuals total should be 0 (no actuals entered), ceiling should match partner_ceiling.

**Risk:** Actuals is new in v3 — never been E2E validated.

**Concrete tests:**

1. **Actuals TOTAL Sum = 0 after recalc** (fresh workbook)
2. **Actuals TOTAL Ceiling = partner_ceiling value after recalc**

**File:** `tests/test_e2e.py` (extend existing)
**Tests:** 2 new tests

---

## GAP 9: Skeleton Mode Reader Round-Trip (E2E)

**What's untested:** Generate skeleton → reader should find all monthly values as None. Currently only demo mode is round-tripped through the reader.

**Risk:** Skeleton mode is the production entry point. If the reader misclassifies empty cells, update mode would break.

**Concrete tests:**

1. **read_capacity_sheet on skeleton returns all-None monthly values for every person/WP**

**File:** `tests/test_reader_v2.py` (extend existing)
**Tests:** 1 new test

---

## GAP 10: Generate Overwrites With Backup (E2E)

**What's untested:** `generate_workbook()` creates a timestamped backup when the output file already exists (`workbook.py:90-94`). Only `restore_workbook` backup behavior is tested.

**Risk:** Users rely on this to avoid data loss when regenerating.

**Concrete tests:**

1. **generate_workbook creates timestamped backup when file exists**
2. **generate_workbook creates output directory if missing** (line 87)

**File:** `tests/test_cli.py` (extend existing)
**Tests:** 2 new tests

---

## Summary Table

| # | Gap | Type | File | New Tests |
|---|-----|------|------|-----------|
| 1 | Conditional formatting rules | Integration | test_conditional_formatting.py (new) | ~9 |
| 2 | Summary ← Plan + Actuals formulas | Integration | test_integration.py | ~6 |
| 3 | Capacity validation column formulas | Integration | test_capacity.py | ~4 |
| 4 | Freeze panes | Integration | test_workbook.py | 4 |
| 5 | Merged cell ranges | Integration | test_merged_cells.py (new) | ~4 |
| 6 | Reader external commitments round-trip | Integration | test_reader_v2.py | 1 |
| 7 | Summary values after recalc | E2E | test_e2e.py | ~4 |
| 8 | Actuals post-recalc | E2E | test_e2e.py | 2 |
| 9 | Skeleton reader round-trip | E2E | test_reader_v2.py | 1 |
| 10 | Generate overwrite backup | E2E | test_cli.py | 2 |
| | **TOTAL** | | | **~37** |

## Dependency Graph (for parallel implementation)

```
Independent groups (can run as parallel subagents):

Group A: GAP 1 (conditional formatting) — new file, no conflicts
Group B: GAP 5 (merged cells) — new file, no conflicts
Group C: GAP 2 + GAP 3 (cross-sheet formulas) — both modify test_integration.py + test_capacity.py
Group D: GAP 4 + GAP 6 (freeze panes + reader ext) — modify test_workbook.py + test_reader_v2.py
Group E: GAP 7 + GAP 8 (recalc E2E) — both modify test_e2e.py
Group F: GAP 9 + GAP 10 (skeleton reader + overwrite) — modify test_reader_v2.py + test_cli.py

File conflict check:
- test_reader_v2.py: GAP 6 (Group D) + GAP 9 (Group F) → must be sequential
- All other groups are file-independent

Revised parallel plan:
  Batch 1 (parallel): A, B, C, E
  Batch 2 (parallel): D, F  (after batch 1 completes, to avoid test_reader_v2.py conflicts)
```

## Priority Order (if implementing incrementally)

1. **GAP 1** — Conditional formatting (highest audit risk, completely untested)
2. **GAP 2** — Summary cross-sheet formulas (new v3 feature, primary reporting surface)
3. **GAP 3** — Capacity validation columns (user-facing guidance tool)
4. **GAP 7+8** — Post-recalc E2E for Summary + Actuals (validates the whole v3 chain)
5. **GAP 10** — Generate overwrite backup (data safety)
6. **GAP 6+9** — Reader round-trips (foundation for future update mode)
7. **GAP 4** — Freeze panes (UX, low risk)
8. **GAP 5** — Merged cells (cosmetic, low risk)
