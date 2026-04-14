# Hours Timelog v2 — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Ship the three v2 workstreams from `2026-04-14-hours-timelog-v2-design.md`: (W1) row grouping on Capacity + Actuals-Hours, (W2) actual-side view redesign built on cumulative-to-date readouts, (W3) Forecast-vs-Actual EUR reconciliation section on Trends.

**Architecture:** Add one new blue input (`hours_actuals_through_month`) + three derived Config cells (`months_elapsed_hours`, `months_elapsed_eur`, `reconciled_through_month`). Wire those into existing builders. All actual-side derived readouts switch to cumulative-through-a-through-month semantics. No input-cell semantics change — skip-and-cumulate entry remains valid.

**Tech Stack:** openpyxl, python 3.11+, pytest. See `CLAUDE.md` for Critical Rules (formula-only, colour convention, inactive-month greying, named-range invariants).

**Reference:** Design doc — `docs/plans/2026-04-14-hours-timelog-v2-design.md`.

---

## Stage 1 — Foundation (T1–T4)

Stage boundary: all new Config params + named ranges exist, reader round-trips them, `docs/design-invariants.md` + `docs/internals.md` document them (post-_META removal), and row grouping renders on Capacity + Actuals-Hours. No consumer code yet depends on the new named ranges.

### Task 1: Add `hours_actuals_through_month` input + three derived named ranges

**Files:**
- Modify: `budget_planner/config_sheet.py` (add param row + derived cells + named-range registrations)
- Modify: `budget_planner/layout.py` (extend `ConfigParamRows` TypedDict)
- Test: `tests/test_config_sheet.py`

**Step 1: Extend `ConfigParamRows`.**

Add `"Hours Actuals Through Month": int` as the 8th key, after `"EUR Actuals Through Month"`.

**Step 2: Add param_def + param_rows entries in `config_sheet.py`.**

In `param_defs` (around line 648), append:
```python
("Hours Actuals Through Month", 1, "0", None),
```

In `param_rows`, add `"Hours Actuals Through Month": 8`.

**Step 3: Add derived-cells block immediately after the params block.**

After the `for i, (label, value, fmt, _) in enumerate(...)` loop, insert three derived cells in rows 9–11 (label col A, formula col B, BLACK_FORMULA font, no border):
```python
derived_params = [
    ("Months Elapsed (Hours)",  "=DATEDIF($B$2, $B$8, \"M\") + 1"),
    ("Months Elapsed (EUR)",    "=DATEDIF($B$2, $B$7, \"M\") + 1"),
    ("Reconciled Through Month", "=MIN($B$7, $B$8)"),
]
derived_first_row = len(param_defs) + 1  # = 8 + 1 = 9
for offset, (label, formula) in enumerate(derived_params):
    r = derived_first_row + offset
    style_cell(ws, r, 1, value=label, font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, value=formula, font=BLACK_FORMULA, fmt="0", border=False)
```

Note the `budget_section_row` computation (`len(param_defs) + 2`) must be updated to account for the 3 new derived rows: `budget_section_row = len(param_defs) + len(derived_params) + 2`.

**Step 4: Register four new named ranges.**

Alongside the existing `eur_actuals_through_month` registration (~line 745):
```python
hours_through_row = param_rows["Hours Actuals Through Month"]
wb.defined_names.add(DefinedName(
    "hours_actuals_through_month",
    attr_text=f"{SHEET_CONFIG}!$B${hours_through_row}",
))
# Derived cells — positions known from derived_first_row
wb.defined_names.add(DefinedName(
    "months_elapsed_hours",
    attr_text=f"{SHEET_CONFIG}!$B${derived_first_row}",
))
wb.defined_names.add(DefinedName(
    "months_elapsed_eur",
    attr_text=f"{SHEET_CONFIG}!$B${derived_first_row + 1}",
))
wb.defined_names.add(DefinedName(
    "reconciled_through_month",
    attr_text=f"{SHEET_CONFIG}!$B${derived_first_row + 2}",
))
```

**Step 5: Write tests.**

Add to `tests/test_config_sheet.py`:
- Test all four new named ranges exist, point at Config cells, and their formulas match.
- Test label col A carries the expected text for row 8, 9, 10, 11.
- Test `hours_actuals_through_month` cell value is `1` (seed) and font is `BLUE_INPUT`.
- Test derived cell fonts are `BLACK_FORMULA`.

**Step 6: Run tests.**

Run: `pytest tests/test_config_sheet.py -v`
Expected: all pass.

**Step 7: Run full pyright.**

Run: `pyright`
Expected: 0 errors.

---

### Task 2: Reader + rebuild round-trip for new Config rows

**Files:**
- Modify: `budget_planner/reader.py` (tolerate the 4 new rows in the Project Parameters section)
- Test: `tests/test_reader_v2.py` or `tests/test_rebuild_shape_drift.py`

**Step 1: Inspect current reader logic for the Project Parameters section.**

Grep `Project Parameters|param_rows|read_project_params` in `reader.py`. Confirm reader does not hardcode "stop at row 7" — it should key off the blank separator row before the WP BUDGET header.

**Step 2: Adjust if the reader misses new rows.**

If necessary, change bounding logic so params through row 11 are parsed. `hours_actuals_through_month` value must round-trip.

**Step 3: Write round-trip test.**

Init workbook → user edits B8 to `5` via openpyxl → save → `rebuild_workbook` → reopen → assert B8 is still `5` and all four named ranges still point at the correct cells.

**Step 4: Run tests + pyright.**

`pytest tests/test_reader_v2.py tests/test_rebuild_shape_drift.py -v && pyright`

---

### Task 3: ~~_META update~~ — OBSOLETE (2026-04-14)

The `_META` sheet was removed on 2026-04-14 (see `docs/plans/archived/2026-04-14-remove-meta-sheet-design.md`). Cross-sheet invariants now live in `docs/design-invariants.md` and named-range documentation lives in `docs/internals.md`. When resuming, add the two-cadence note + new named-range entries to those docs instead. Skip this task — no code change needed.

---

### Task 4: W1 row grouping on Capacity + Actuals-Hours

**Files:**
- Modify: `budget_planner/person_section.py` (shared geometry helper — add outline-level application)
- Modify: `budget_planner/capacity.py` (call helper)
- Modify: `budget_planner/actuals_hours.py` (call helper)
- Test: `tests/test_grouping.py` (new assertions) or new `tests/test_row_grouping.py`

**Step 1: Decide where to apply `outline_level=1`.**

Per the design: each person's WP rows + TOTAL + EU TOTAL rows get `outline_level=1`. Person-header row stays at `0`. Read `person_section.py` to find the single source of person-block geometry.

**Step 2: Add `summaryBelow = False` to both sheets.**

In `capacity.py` and `actuals_hours.py` build functions:
```python
ws.sheet_properties.outlinePr.summaryBelow = False
```

**Step 3: Apply `row_dimensions[r].outline_level = 1` to body rows.**

Extend the shared helper (or add a small helper alongside) that returns the body-row range per person and sets the outline level.

**Step 4: Tests.**

Per sheet (Capacity, Actuals-Hours):
- For each person: header row has `outline_level == 0`, every body row (WP + TOTAL + EU TOTAL) has `outline_level == 1`.
- `ws.sheet_properties.outlinePr.summaryBelow is False`.
- Rebuild round-trip: outline levels survive `rebuild_workbook`.

**Step 5: Run tests + pyright.**

`pytest tests/test_grouping.py tests/test_capacity.py tests/test_actuals_hours.py tests/test_rebuild.py -v && pyright`

---

**Stage 1 completion gate:**

- All Stage 1 tasks committed.
- `pytest tests/ -v` green (all ~420 tests).
- `pyright` clean.
- Named-range inventory + two-cadence note landed in `docs/design-invariants.md` + `docs/internals.md` (post-_META removal).
- Save checkpoint in `memory/project_state.md` noting Stage 1 done + next-up Stage 2 (T5–T9).

Run retrospective via `/retro` at the stage boundary.

---

## Stage 2 — W2 actual-side view redesign (T5–T9)

*Fresh `/impag` session after Stage 1 commit.*

### Task 5: Capacity to-date column block

Add far-right column group (own outline level, collapsible):
- Plan PM to date (SUMIFS of Plan FTE cells where month ≤ `hours_actuals_through_month`)
- Actual PM to date (SUM of Actuals-Hours WP cells where month ≤ `hours_actuals_through_month`, divided by `hours_per_fte_month`)
- Variance PM to date (Plan − Actual)
- Variance % with red CF at ±20% (both directions)

All black (same-sheet formulas) or green (cross-sheet to Actuals-Hours). Tests cover skip-and-cumulate, monthly-only, mixed entry patterns.

### Task 6: Capacity per-person summary rows + raw-monthly collapsed band

- Per-person: Actual PM to date, Plan PM to date, Utilisation to date (red ≥ 100%, amber ≥ 80%), EU Actual PM to date, EU Ceiling pro-rated.
- Wrap existing raw monthly Actual FTE / Variance FTE bands in `outline_level = 1` column group, default collapsed. Band header: "Raw monthly — reliable only when hours were entered per-month".

### Task 7: Actuals Breach-In cumulative-slope + As-of headers

- Breach In: `(ceiling − cum EUR) / (cum EUR / months_elapsed_eur)` — replaces recent-month-slope.
- Forecast Breach In parallel: uses `months_elapsed_hours`.
- Add two green "As of" header cells above Actual and Forecast blocks.

### Task 8: Summary per-column partial-year shading

- Forecast column shades by `hours_actuals_through_month`.
- Actual column shades by `eur_actuals_through_month`.
- Variance cell shades by stricter MIN.
- Variance-vs-Plan CF only fires on fully-covered years (gated by `year_coverage_pct`).

### Task 9: Actuals-Hours + Plan-Forecast cosmetic shading + hours-to-date rows

- Actuals-Hours: cosmetic grey shading past `hours_actuals_through_month` on forecast-EUR leaves; add "Cumulative hours to date" + "Actual PM to date" summary rows above existing TOTAL.
- Plan-Forecast: cosmetic grey shading past `hours_actuals_through_month`.

**Stage 2 completion gate:** same as Stage 1 (full tests + pyright + invariant/internals doc updates + checkpoint + /retro).

---

## Stage 3 — W3 reconciliation (T10–T13)

*Fresh `/impag` session after Stage 2 commit.*

### Task 10: Trends LineChart series clipping

Update existing cumulative-EUR LineChart series so:
- Actual cumulative points past `eur_actuals_through_month` return `NA()`.
- Forecast cumulative points past `hours_actuals_through_month` return `NA()`.
- Planned + Ceiling run through project end unchanged.

### Task 11: Trends Person Load table update

- Remove peak FTE column.
- Replace avg FTE with Avg FTE to date (`actual PM to date / months_elapsed_hours`).
- Add Plan avg FTE to date.
- Add Variance to-date % with same CF thresholds as Capacity's Variance %.

### Task 12: Trends Reconciliation section — health cells + summary table

Placement: between existing LineChart and Person Load.

Health cells (three rows):
- Actuals EUR through (green `eur_actuals_through_month`)
- Hours through (green `hours_actuals_through_month`)
- Reconciled through (black `reconciled_through_month`)

Summary table (one row per programme WP):
- WP, Cum Forecast EUR, Cum Actual EUR, Gap EUR, Gap %
- Bottom totals row
- Gap % CF: red ≥ ±10%, amber ≥ ±5% (absolute)

### Task 13: Trends Reconciliation gap chart

- Single LineChart, single series: cumulative Actual EUR − cumulative Forecast EUR across months.
- `NA()` past `reconciled_through_month`.
- Gridlines + zero reference line (dashed).

**Stage 3 completion gate:** same as prior stages.

---

## Stage 4 — Polish & integration (T14–T15)

### Task 14: Rebuild preservation

- Outline row levels must survive `rebuild_workbook`.
- Four new named ranges must round-trip.
- `eur_actuals_through_month` + `hours_actuals_through_month` user-entered values must be preserved.
- Add tests to `test_rebuild_shape_drift.py` covering shape-drift scenarios with new Config rows (insert/delete within the Project Parameters band).

### Task 15: Full e2e + shape fingerprint

- Extend `test_shape_fingerprint.py` if needed.
- Run full `pytest tests/ -v` + `pyright`.
- Dispatch `superpowers:code-reviewer` for BASE_SHA..HEAD_SHA review.
- Final `/retro`.

---

## Notes for executors

- **Always use Excel formulas, never hardcoded values** (Rule 1).
- **Colour convention** (Rule 3): blue only for `hours_actuals_through_month` input; green for cross-sheet refs; black for same-sheet derived. No new conventions.
- **After every task**: `python -m budget_planner init` + `python scripts/recalc.py output/budget_plan.xlsx` + `pytest tests/ -v` + `pyright`. Commit only when green.
- **Do not touch** Plan sheet builders (out of scope — Plan is forward-looking monthly and untouched).
- **Do not** reinterpret input cells as cumulative — explicitly rejected in the design.
- When a to-date formula references `Actuals-Hours`, prefer SUMIFS with the Config through-month as the comparison RHS — avoids per-month ISNUMBER gymnastics that fail cross-platform (see `.claude/rules/openpyxl-patterns.md`).
