# Rebuild correctness under Config-tab row-count changes

**Date:** 2026-04-12
**Status:** Design agreed, implementation plan pending

## Problem

Hand-edits to the Config sheet (inserting rows in the middle of tables) cause
`rebuild` to misbehave:

1. Persons are dropped — reader claims they're absent from the Team table when
   they exist.
2. Formulas in Capacity/Actuals point at pre-edit row positions after rebuild.

User confirmed: rows were inserted in the **middle** of ranges (not appended),
and Excel's auto-extension of named ranges still covers all rows.

## Root cause

`_META` is treated as a trusted source of truth on rebuild. Excel auto-extends
**named ranges** and **formula ranges** on row insertion, but it does **not**
update plain cell values on other sheets. `_META` is a plain data sheet whose
row-count cells are written once by the generator and never refreshed by
Excel-side edits. Rebuild reads stale META → boundaries lag reality → persons
past the old last row are dropped and formula rewrites use stale offsets.

Per CLAUDE.md Critical Rule 6, `_META` is supposed to be kept in sync — but
there is no runtime enforcement, and Excel hand-edits are a normal workflow
that bypasses any Python-side update path.

## Design

### Section 1 — make `_META` derived, not trusted

On rebuild, the reader discovers table shape **live** from the Config sheet.
`_META` becomes a derived artifact that is rewritten at the end of rebuild;
it is never read *before* it is re-derived.

- **Primary source of shape:** named ranges
  (`team_table`, `wp_table`, `assignments`, etc.). Excel maintains these
  correctly across middle-insertions.
  `wb.defined_names['<name>'].destinations` gives the current span.
- **Fallback for tables without named ranges:** header-row lookup + scan
  until first fully-empty row.
- **Safety net:** if named-range span and empty-row scan disagree, fail loud
  with a message naming the table and the two conflicting row counts. The
  user can then fix the sheet rather than silently lose data.
- **META rewrite:** at the end of rebuild, recompute all `_META` entries
  from the discovered live shape. Rule 6 still holds for the documentation
  value of `_META`, but the fragility is gone.

This eliminates the whole stale-META failure mode instead of patching around
it.

### Section 2 — testing strategy

Three layers, each catching a different bug class.

**(a) Shape-parameterized rebuild tests — lean matrix**

A new `test_rebuild_shapes.py` using `pytest.mark.parametrize`. **Scope: default
shape plus one variant per table**, not the full cross-product. Concretely:

| Case | n_wps | n_team | n_partners | n_grade_bands |
|------|-------|--------|------------|---------------|
| default | D | D | D | D |
| +wps | D+2 | D | D | D |
| +team | D | D+3 | D | D |
| +partners | D | D | D+1 | D |
| +gradebands | D | D | D | D+1 |

(Where `D` = current default sample shape. Exact deltas TBD in the plan.)

For each case: generate workbook → populate Assignments and Actuals with
deterministic values (e.g., `person_i × wp_j = i*10+j`) → run `rebuild` →
assert every person round-trips, every assignment value round-trips, every
Capacity/Actuals formula resolves to a cell inside the correct live table
range.

**(b) Excel-edit simulation test — pins the exact bug**

One dedicated test that mimics the user's hand-edit: open a generated
workbook, `ws.insert_rows(mid_row, amount=3)` in the middle of the Team
table, write values into the new rows, save, then `rebuild`. Assert the
three inserted persons survive and their formulas point at post-insert rows.
This test must fail on current code and pass after the Section 1 fix.

**(c) META-divergence invariant**

A fast helper `assert_meta_matches_live(wb)` called at the end of every
rebuild test. Catches future regressions where a META-trusting code path is
reintroduced.

### Section 3 — privacy workflow for the user's real file

The common case is already solved: Sections 1–2 reproduce the bug from
**shape alone**, no real values needed. The user's file never leaves their
machine for normal debugging.

**Escape hatch for shape-specific bugs:** `scripts/diagnose_shape.py`.

Emits only non-sensitive structural info:

- Per Config table: named-range span, first/last non-empty row, row count,
  whether named-range and empty-row scan agree.
- `_META` contents vs. live shape — diff only.
- Capacity/Actuals cross-sheet formula *coordinates*
  (e.g., `Capacity!D15 → Config!B23`) with no resolved values.
- Never emits: names, salaries, rates, partner names, or any cell value
  outside `_META` numeric counts.

The user runs it locally and pastes stdout. That's enough to hand-craft a
synthetic reproducer matching the exact shape.

**Explicit non-goals (YAGNI):**

- No anonymizer / value-mapping pipeline.
- No redacted copy of the real file committed to the repo.
- No changes to `.gitignore` beyond a one-time audit that the real file
  is not sitting under an allow-listed name
  (`budget_plan.xlsx`, `budget_plan_change.xlsx`,
  `budget_plan_demo.xlsx`, `budget_plan_init.xlsx`,
  `budget_plan_rebuild.xlsx`).

## Out of scope

- Rewriting the generator to stop emitting `_META` at all. Keep it as
  documentation; just stop trusting it as input.
- Cell protection re-enablement (tracked separately per Rule 7).
- Any change to the user-facing CLI surface.

## Next step

Produce an implementation plan under `docs/plans/` breaking this into
bite-sized tasks (reader changes, META rewrite, three test files,
diagnostic script) with the order-of-operations and a proof-of-reproduction
task at the front.
