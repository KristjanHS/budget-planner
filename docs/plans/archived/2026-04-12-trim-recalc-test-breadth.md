# Trim recalc test breadth

**Date:** 2026-04-12
**Goal:** Reduce `pytest` runtime by deleting low-value recalc tests and consolidating a duplicated fixture. No marker split — shrink the slow suite, don't hide it.

## Baseline

- 307 tests in **63s**
- Top 5 durations consume ~38s (60% of runtime) — all LibreOffice recalc
- `scripts/recalc.py` costs ~7s per subprocess call

## Diagnostic

| Hot spot | Cost | Observation |
|---|---|---|
| `test_e2e.py` demo + empty module recalc setups | 8.96s + 7.08s | Module-scoped fixtures (already shared) |
| `TestRebuildRecalc` (3 tests) | ~22s | Each test runs `recalc.py` **independently** — not shared |
| Everything else | <0.6s each | Rebuild round-trips, acceptable |

Deleting tests that share a module-scoped recalc fixture does **not** save time directly — it reduces assertion noise and maintenance burden. Time savings come from:
1. Deleting tests that spawn their own recalc subprocess.
2. Collapsing `TestRebuildRecalc` to one recalc call.

## Testing gate (reference)

From `~/.claude/rules/testing.md`:

1. Testing my code or the language?
2. Already covered by a higher-layer test?
3. Could this test ever fail from a real bug?
4. Too thoroughly testing a trivial function?
5. Testing an impossible input?

## Deletions

Each entry cites the failing gate.

### 1. `test_demo_mode_formula_count` (`tests/test_e2e.py:92`)

```python
def test_demo_mode_formula_count(demo_recalc):
    assert demo_recalc["total_formulas"] > 200
```

**Gate #1, #3.** Tests openpyxl's formula-counting, not our logic. Only fails if a whole builder is deleted — which would be caught by every other e2e test first.

### 2. `test_demo_mode_plan_planned_equals_capacity_sums` (`tests/test_e2e.py:100`)

Docstring claims "Plan Planned FTE values should be numeric (sum of Capacity)". Actual assertion:

```python
assert val is not None and isinstance(val, (int, float))
```

**Gate #2.** It doesn't check the sum equality — just numeric-ness. That's implied by `test_demo_mode_recalc_zero_errors`: if recalc produces zero errors, the formula resolved to a number.

### 3. `test_summary_year1_plan_fte_numeric` (`tests/test_e2e.py:139`)

```python
assert val is None or isinstance(val, (int, float))
```

**Gate #2.** Same pattern as above. `test_summary_total_plan_fte_equals_year_sum` already reads these cells and performs arithmetic on them — if they were non-numeric, that test would fail first.

### 4. `test_actuals_total_sum_positive` (`tests/test_e2e.py:246`)

Asserts Actuals TOTAL Sum > 0 in demo mode.

**Gate #2.** Covered by `test_summary_actual_eur_seeded` — both depend on the same seeded data reaching Summary/Actuals via the same SUMIF chain. If one works, the other does.

### 5. `test_changed_file_has_autofit_columns_after_recalc` (`tests/test_rebuild_e2e.py:747`)

Verifies `_changed` file has non-default column widths after recalc runs the autofit macro.

**Gate #1.** The only thing this tests is whether LibreOffice's autofit macro ran and modified widths — which is LibreOffice behavior, not our code. **Saves ~7.5s** because it runs its own `recalc.py` subprocess.

## Consolidation

### `TestRebuildRecalc` class-scoped fixture (`tests/test_rebuild_e2e.py:678`)

After deleting the autofit test, the remaining two tests both:
1. Set `CONFIG_FIRST_WP_ROW, col 3` to `50.0`.
2. Call `rebuild_workbook(init_path)`.
3. Spawn `recalc.py`.
4. Assert on the result.

Replace with a class-scoped fixture that performs steps 1–3 once and caches the recalc result + workbook path. Both tests become pure assertions against the cached state.

**Saves ~7s** (two recalc calls → one).

## Expected outcome

- **Runtime:** 63s → ~48s (~24% faster)
- **Test count:** 307 → 302
- **Coverage:** Unchanged. Every deleted test is either redundant or tests non-owned code.

## Execution order

1. Delete `test_changed_file_has_autofit_columns_after_recalc` (biggest time win; isolated).
2. Delete the four low-value e2e assertions.
3. Consolidate `TestRebuildRecalc` to class-scoped fixture.
4. Run `pytest tests/ -q --durations=10`. Confirm green + record new baseline.

## Kept tests (high-value recalc canaries)

- `test_demo_mode_recalc_zero_errors` — demo-mode formula canary
- `test_empty_mode_recalc_zero_errors` — empty-mode distinct code path (SUMIF-on-empty etc.)
- `test_summary_total_plan_fte_equals_year_sum` — cross-cell summation logic
- `test_summary_actual_eur_seeded` — seeded actuals flow to Summary
- `test_summary_budget_overview_after_recalc` — `sum(FTE * per-WP rate)` ceiling (core programme invariant)
- `test_actuals_total_ceiling_equals_partner_ceiling` — cross-sheet ceiling equality
- `test_recalc_zero_errors_after_fte_change` — rebuild doesn't corrupt formulas
- `test_recalc_ceiling_reflects_changed_fte` — rebuild preserves computational correctness
