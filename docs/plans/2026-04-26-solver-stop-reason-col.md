# Solver stop-reason column on PER-WP SUMMARY

**Date:** 2026-04-26
**Scope:** single-session, ~6 files touched
**Status:** design — ready for implementation

## Goal

Add a 1–2-word "Stop reason" string per WP at the right edge of the PER-WP
SUMMARY base columns, explaining why the solver finished filling that WP.
Distinct logical termination branches must produce distinct strings so a
human reader can tell at a glance whether the WP is filled, capped,
overshot via the lower-cap exception, capacity-bound, or has no candidates.

User-facing examples (from the prompt): "person max FTE running out" → one
string, "person-WP assignments month range not allowing more months" →
another, "+12% buffer" → another. The full reason set is fixed (closed
enum) — see § Reason set below.

## Design decisions

### D1. Column placement: end of base cols (col J), not absolute right edge

Insert "Stop reason" at col J — the new last base column, immediately
before the dynamic year-col band.

| Before | After |
|--------|-------|
| `A..I` base cols, `J..` year cols | `A..J` base cols (+ "Stop reason"), `K..` year cols |

Rationale: stop-reason is per-WP metadata, identical in shape to existing
base-col metrics like "Buffer %" or "Burn rate needed". The trailing year
band is a numeric breakdown; inserting text metadata after it would break
the existing visual band hierarchy. Stable col J position regardless of
project span. The user phrase "at the end of PER-WP SUMMARY" is read as
"at the end of the base-metric band" (where all the other per-WP metrics
already are).

Alternative (rejected): place at dynamic absolute right edge after year
cols. Trades stable col position for a more literal "at the end" reading;
yearly-rows below get an awkward empty trailing cell.

### D2. Reason classification: post-pass derived, not solver-internal

Add a new pure function `classify_wp_stop_reasons` in `timelog_solver.py`
that takes the final `Assignment` plus the same inputs `solve` consumes
and returns `dict[str, str]` (`wp_code -> reason_string`).

Rationale: the solver's existing public surface (`solve`, `greedy_pass`,
`switch_reduction_pass`) is consumed by 18 unit tests asserting only on
`Assignment`. Returning a tuple from `solve` would force every test +
`build_timelog_plan_sheet` call site to unpack. A separate classifier:

- preserves the existing solver signature (zero test breakage);
- isolates UI-string concerns from the solver hot loop;
- is independently testable with hand-crafted final states.

The trade-off accepted: classification is heuristic rather than
authoritative. The classifier reconstructs which break fired by examining
final state + remaining capacity. For every reason in § Reason set, final
state is sufficient to determine the cause unambiguously — see the
classifier algorithm below for proof-by-cases.

Alternative (rejected): instrument `greedy_pass` to record the exact break
that fired per WP. Cleanest mapping but couples solver to UI strings and
breaks every direct `greedy_pass` test.

### D3. Reason set (closed enum, 1–2 words each)

Mutually exclusive, exhaustive, classifier-priority order:

| String | Trigger condition |
|--------|-------------------|
| `"No candidates"` | No `PersonWPWindow` row references this WP for any person |
| `"No window"` | `PersonWPWindow` rows exist but none overlaps the WP's active range |
| `"-11% override"` | Final `placed_eur > ceiling × (1 + MAX_BUFFER_PCT)` — lower-cap tie-break placed beyond the upper cap |
| `"Filled"` | Final `placed_eur >= ceiling_eur` (and not overshooting beyond +12%) |
| `"+12% cap"` | Final state in `[ceiling × 0.89, ceiling)` AND at least one eligible candidate has remaining pool, a legal unplaced month, and a rate that would have pushed the next placement past the upper cap |
| `"Pool empty"` | Final state below ceiling AND no eligible candidate has remaining pool |
| `"Window full"` | Final state below ceiling AND eligible candidates have pool BUT no legal unplaced months remain in any (person, wp) intersection (vacation/frozen/already-placed exhaust the windows) |
| `"Underfilled"` | Catch-all fallback for any state the above don't cover (should be unreachable; surface as a signal of policy drift, not silently mask) |

Rationale for each-name choice:
- `"No candidates"` vs `"No window"`: distinguishes "user never entered
  any assignment row" from "user entered rows but none overlap the
  active period".
- `"-11% override"` mirrors the constant `MIN_BUFFER_PCT` and the
  policy term in `timelog_solver.py`'s docstring.
- `"+12% cap"` matches `MAX_BUFFER_PCT` semantics. The `+` is intentional
  — visually paired with `-11%`.
- `"Pool empty"` vs `"Window full"`: capacity-bound vs window-bound. Both
  are infeasibility cases; the user's debugging path differs (raise
  `max_fte` for one, widen `start`/`end` for the other).

### D4. Threading

`build_timelog_plan_sheet` already calls `solve(persons, wps, windows,
project_months)` at `budget_planner/timelog_plan.py:285`. Immediately after
the frozen-placement overlay (line ~291), call:

```python
stop_reasons = classify_wp_stop_reasons(
    assigned_map, persons, wps, windows, project_months,
)
```

Pass `stop_reasons` into `_write_wp_summary_row` as a new kwarg. That
function writes the per-WP base cells; add a 10th `style_cell` call
writing `stop_reasons.get(wp.code, "Underfilled")` at col J with
`BLACK_FORMULA` font and `CENTER` alignment.

The yearly-rows region (`_write_yearly_rows`) gets nothing in col J on
those rows — leave empty. Section-title merge automatically extends to
the new last col because `wp_summary_last_col = _WP_SUMMARY_BASE_LAST_COL
+ len(year_specs)` (line 403) recomputes from the bumped constant.

## Implementation steps

Each step is independent unless noted. Steps 2 and 3 must land before
step 4 (the integration test depends on both).

### Step 1 — Layout constant bump

**File:** `budget_planner/layout.py`

- `TIMELOG_WP_SUMMARY_BASE_LAST_COL = 9` → `10`. Update the comment to
  mention the new col J = "Stop reason" and that year cols now start at
  col K.
- `TIMELOG_YEARLY_LABEL_COL` stays at `9` (col I — yearly-row labels
  don't move). Update the surrounding comment to clarify that the
  label sits adjacent-left to the new "Stop reason" col.

**Verification:** the assertion at `timelog_plan.py:129`
(`assert _WP_SUMMARY_BASE_LAST_COL == TIMELOG_WP_SUMMARY_BASE_LAST_COL`)
will fail until step 3 lands the matching `_WP_SUMMARY_BASE_HEADERS`
extension. That's expected — this is a layout-shift step, validated only
when step 3 lands. Do not run pytest after step 1 alone.

### Step 2 — Classifier in solver module

**File:** `budget_planner/timelog_solver.py`

Add:

```python
def classify_wp_stop_reasons(
    assigned: Assignment,
    persons: list[Person],
    wps: list[WP],
    person_wp_windows: list[PersonWPWindow],
    project_months: int,
) -> dict[str, str]:
    """Per-WP termination reason. Pure; derives reasons from final state.

    Algorithm (per WP, in priority order — first match wins):
      1. No PersonWPWindow rows reference this WP → "No candidates"
      2. Rows exist but no row's window overlaps wp.active_start..end
         → "No window"
      3. final_placed_eur > ceiling × (1 + MAX_BUFFER_PCT) → "-11% override"
      4. final_placed_eur >= ceiling_eur → "Filled"
      5. else (under ceiling): inspect remaining capacity per candidate.
         - Any candidate has pool > 0 AND ≥1 legal unplaced month AND
           rate would push past upper cap → "+12% cap"
         - All candidates have pool == 0 → "Pool empty"
         - Some candidate has pool > 0 but no legal unplaced months
           → "Window full"
         - Otherwise (unreachable) → "Underfilled"
    """
```

Reason-derivation helpers (private to the module):

- `_remaining_pool_after(assigned, person_name, person.pool)` — pool minus
  count of non-None cells in `assigned[person_name]`.
- `_legal_unplaced_months(person, wp, windows_idx, assigned)` — months
  that satisfy active-window + person-WP-window + not-vacation +
  not-frozen + currently None in `assigned`.
- `_placed_eur_for_wp(assigned, persons, wp_code)` — sum of
  `rate_eur_per_pm` across all `(person, month)` cells assigned to
  `wp_code`.

Tolerance: use `>=` and `>` exactly as documented (no epsilon needed —
all values are integers).

### Step 3 — PER-WP SUMMARY column header + per-WP write

**File:** `budget_planner/timelog_plan.py`

- Append `"Stop reason"` to `_WP_SUMMARY_BASE_HEADERS` at line 121–125.
  This makes `_WP_SUMMARY_BASE_LAST_COL` = 10 automatically (it's
  `len(_WP_SUMMARY_BASE_HEADERS)`); the assertion at line 129 starts
  passing.
- Header style: plain `HEADER_FONT`, no `+i`/`-i` fill (it's not a
  monetary column).
- After line 285 (`assigned_map = solve(...)`), add the
  `classify_wp_stop_reasons` call. Pass `stop_reasons` to
  `_write_wp_summary_row` as a new kwarg.
- Inside `_write_wp_summary_row`, add a final cell write at col 10:
  `stop_reasons.get(wp.code, "Underfilled")` with `BLACK_FORMULA` font
  and `CENTER` alignment.
- Yearly rows: no col-10 write needed — `_write_yearly_rows` writes only
  cols 3 and the year-col band; col 10 stays empty.

**Verification:** confirm `_write_wp_summary_row` and `_write_yearly_rows`
year-col offsets at lines 1124 and 1323 use
`_WP_SUMMARY_BASE_LAST_COL + 1 + i` (already verified in design
research) so the constant bump propagates cleanly. No code change needed
to those offsets.

### Step 4 — Spec doc

**File:** `docs/spec.md`

Update the section describing PER-WP SUMMARY columns (search for
`"Buffer %"` and `"Any out-of-window?"`). Add a row for `Stop reason` at
col J explaining the reason set.

If the spec lists year cols at cols J..N, update to K..O (post-shift).

If the named-range table's row schema is parsed by
`tests/test_config_sheet.py::test_config_named_range_targets_match_spec_doc`
(per `.claude/rules/plan-hygiene.md`), DO NOT alter that table's pipe
count or column order — only edit unrelated prose.

### Step 5 — Solver classifier tests

**File:** `tests/test_timelog_solver.py`

Add 7 tests covering each enum value:

1. `test_stop_reason_filled` — single person, ample pool, small WP
   ceiling. Expect `"Filled"`.
2. `test_stop_reason_overshoot_override` — re-uses the
   `test_lower_cap_forces_overshoot_on_small_wp` scenario (rate=600,
   ceiling=1000). Expect `"-11% override"`.
3. `test_stop_reason_upper_cap` — rate that pushes from ~0% buffer to
   above +12% on the next placement, with current state already above
   -11% (so the override doesn't fire). Expect `"+12% cap"`.
4. `test_stop_reason_pool_empty` — ceiling huge, person pool small (e.g.
   pool=2, ceiling=10000, rate=1000). All pool consumed before reaching
   ceiling. Expect `"Pool empty"`.
5. `test_stop_reason_window_full` — pool ample, but PersonWPWindow only
   covers 1 month inside active range; vacation eats the rest. Expect
   `"Window full"`.
6. `test_stop_reason_no_window` — `PersonWPWindow("Alice", "WP1", 1, 5)`
   but `WP1.active_start=10, active_end=12`. No overlap. Expect
   `"No window"`.
7. `test_stop_reason_no_candidates` — WP1 exists but no
   `PersonWPWindow` row references it. Expect `"No candidates"`.

Each test calls `greedy_pass` (not `solve` — `switch_reduction_pass`
doesn't change WP totals) then `classify_wp_stop_reasons` and asserts
`reasons["WP1"] == "<expected>"`.

### Step 6 — Sheet integration test

**File:** `tests/test_timelog_plan.py`

Add one test that asserts the col J header value is `"Stop reason"` on
the WP-summary header row, and that for at least one WP the cell at
col J on its WP row is non-empty (any of the 7+1 reason strings).

If a sibling test asserts the year-col positions on PER-WP SUMMARY (e.g.
`Y2026 +i` at col J), find it and update to col K. Grep for
`"Y20"|first_year_col|year_col_start` inside `tests/test_timelog_plan.py`
to locate.

### Step 7 — Full validation

Run:

```bash
python -m budget_planner init
python scripts/recalc.py output/budget_plan.xlsx
pytest tests/ -q  # use timeout: 300000 from CLAUDE.md dev-workflow
```

All 607+8 = ~615 tests should pass.

## Out of scope

- **Conditional formatting on the new col.** Could colour-code reasons
  (red for "No candidates", amber for "+12% cap", green for "Filled").
  Defer until user feedback confirms the reason set is stable.
- **Reasons on PER-PERSON SUMMARY.** Same idea applied per-person ("why
  this person didn't fill more") would need a separate classifier and a
  separate column. Not requested.
- **Reason for `switch_reduction_pass` outcomes.** Pass-2 only swaps
  rate-equal cells and preserves WP totals; it can't change the stop
  reason. No reason string emitted for it.
- **Internationalisation.** Strings are English literals. No locale
  table.
- **Reason explanations / tooltips.** The 1–2 word string is the only
  surface; no Excel comment or hover text.

## Open questions

None — every decision above is committed. If the implementation reveals
that final state is *not* sufficient to distinguish "+12% cap" from
"Window full" in some edge case (e.g. a candidate with pool but a rate
that's both cap-blocked AND has no legal months left), prefer "+12% cap"
in the priority order — the cap is the more specific signal.

## Re-work audit

| Step | File | Concern |
|------|------|---------|
| 1 | `layout.py` | constant bump |
| 2 | `timelog_solver.py` | new function |
| 3 | `timelog_plan.py` | header + write |
| 4 | `docs/spec.md` | doc |
| 5 | `tests/test_timelog_solver.py` | new tests |
| 6 | `tests/test_timelog_plan.py` | new test |

No file appears in two steps. No write-then-rewrite risk. Single-pass
implementation; no phased roll-out.
