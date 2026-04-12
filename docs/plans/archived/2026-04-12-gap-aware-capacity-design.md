# Gap-Aware Capacity Distribution — Design

**Status:** Ready to implement
**Date:** 2026-04-12
**Author:** brainstorm session

## Problem

A WP in `WP_SCHEDULE` may have multiple active ranges (gaps in between — e.g. `WP1` active months 8–17 and 20–26, inactive 18–19). When a person is assigned to the WP across the gap (e.g. Alice → WP1, window M10–20, `total_fte=5.5`), the current builder:

- Renders gap months M18, M19 as **blue-editable cells** (not grey) and fills them with distributed FTE.
- Helper columns `Sum`, `Budget`, `Variance`, `Sig Assigned`, `Balance` all agree that Alice is committing 5.5 PM to WP1 — but 1.0 PM of that is allocated to months when the WP is inactive (auditor-invalid).
- `Plan!Assigned PM` independently filters by `wp_active_months()` and reports 4.5 — creating a silent cross-sheet discrepancy.

Reproducer: `scripts/investigate_gapped_wp.py` (generates `output/investigate_gapped_wp.xlsx`, runs LO recalc, prints 4 FAILs).

The user wants a single assignment row to cover the whole window and the tool to exclude gap months everywhere — without splitting the WP into sub-WPs.

## Chosen semantics (α — redistribute)

`Total FTE` stays the user-entered scalar. The builder distributes it across the **intersection of the assignment window and `wp_active_months(wp_id, wp_schedule)`** — not across the raw window span.

Example: Alice, window M10–20, `total_fte=5.5`, WP1 gap at M18–19.
- Valid months in window: `{10, 11, 12, 13, 14, 15, 16, 17, 20}` (9 months)
- Per-month FTE after `distribute(5.5, 9, curve)` ≈ 0.611
- Gap months M18, M19 → grey + empty

Alternative (β — shrink total) was rejected: it would turn `Total FTE` into a derived value and create a "phantom gap" in the committed number that users would trip over.

## Implementation

### Single code change: `budget_planner/capacity.py::_write_person_grid`

Current (`capacity.py:143-147`) loops over all project months and decides blue-input vs grey based only on `asgn_start <= month_num <= asgn_end`.

Replace with an intersection-based check:

```python
wp_active = wp_active_months(wp_id, wp_schedule)          # frozenset[int]
valid_months = {m for m in range(asgn_start, asgn_end + 1)
                if m in wp_active}
monthly_values = distribute(asgn["total_fte"], len(valid_months), wp_curve)

for mi in range(project_params["months"]):
    month_num = mi + 1
    if month_num in valid_months:
        ...blue-input cell with next distributed value...
    else:
        ...grey fill, empty...
```

Requires importing `wp_active_months` from `budget_planner.layout` (already exists — used by `plan.py`, `actuals.py`).

### What stays the same

- Config assignment schema (`person`, `wp`, `total_fte`, `start`, `end`) — unchanged.
- All Excel formulas in every sheet — unchanged.
- Reader / rebuild round-trip logic — unchanged (assignment scalars drive re-distribution on rebuild).
- Plan sheet — already filters by `wp_active_months`; no change needed.

### Why no helper-column changes

Under α, the scalar `Total FTE` **is** the actually-distributed amount (no phantom FTE in gap months). Therefore `SUMIFS` on Total FTE still yields the correct `Sig Assigned` and `Balance`. The cross-sheet discrepancy (Plan 4.5 vs Capacity 5.5) disappears because Capacity now agrees with Plan at 5.5.

### Post-fix reconciliation table

| Sheet / column | Mechanism | Value |
|---|---|---|
| Capacity M18/M19 | builder emits grey + empty | — |
| Capacity `Sum` | `SUM(row)` | 5.5 |
| Capacity `Budget` | ref → Config Total FTE | 5.5 |
| Capacity `Variance` | Budget − Sum | 0 |
| Config `Sig Assigned` | SUMIFS on Total FTE | 5.5 |
| Config `Balance` | `max_fte × months − Sig Assigned − Sig External` | 24.5 |
| Plan `Assigned PM` | sum of monthly planned FTE | 5.5 |

## Edge cases (handled for free)

| Case | Behavior |
|---|---|
| Window ∩ schedule = ∅ | `distribute(total, 0)` → `[]`, all cells grey, Sum = 0, Variance = Total FTE → existing red CF flags it |
| Window spans N ≥ 2 gaps | intersection generalizes; no special code path |
| WP has contiguous schedule (today's demo data) | intersection equals window → identical output, no regression |
| Assignment entirely inside a gap | same as "∅ overlap" — red Variance flag |

## Tests

### New: `tests/test_gapped_wp.py`

Convert `scripts/investigate_gapped_wp.py` into pytest form.

- Shared fixture: build + recalc the gapped-WP workbook once per module.
- One assertion per behavior (~5 tests, no ratio inflation per `~/.claude/rules/testing.md`):
  1. M18 and M19 are grey-filled and empty
  2. Per-month FTE in valid months ≈ `total_fte / len(valid_months)`
  3. Capacity `Sum` equals `total_fte`
  4. Capacity `Variance` equals 0
  5. Plan `Assigned PM` sum equals Capacity `Sum` (cross-sheet reconciliation)
- Delete `scripts/investigate_gapped_wp.py` once the test lands (single source of truth).

### Existing tests to re-verify

- `tests/test_capacity.py::test_capacity_inactive_months_empty` — parametrized over `DEMO_CAPACITY_ASSIGNMENTS`, all of which use contiguous WPs. Should stay green.
- `tests/test_capacity.py::test_capacity_wp_rows_are_blue_input` — picks `asgn["start"]` as "first active month". Safe only if `asgn.start` is not inside a gap. Demo data is gap-free; no change needed, but add a note.

## Rollout

1. Code change to `capacity.py`.
2. New test file `tests/test_gapped_wp.py`.
3. Delete `scripts/investigate_gapped_wp.py`.
4. `budget_planner/meta.py`: one-line update so `_META` documents the gap-aware distribution (Rule 6).
5. `docs/grant-domain.md`: short paragraph under the cost-model section describing gap handling.
6. Recalc `output/budget_plan.xlsx` via `python scripts/recalc.py`; commit the refreshed artifact (per dev-workflow rule: workbooks are tracked).

No migration needed. Existing workbooks with contiguous WP schedules are unaffected. Workbooks with gapped schedules and cross-gap assignments will auto-correct on first rebuild — M18/M19 (etc.) become grey, per-month FTE in valid months rises to absorb the redistributed total.
