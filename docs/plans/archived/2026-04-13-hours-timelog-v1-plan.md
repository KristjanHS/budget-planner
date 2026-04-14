# Hours-Timelog v1 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add hours-based month-end monitoring: new `Actuals-Hours` sheet captures per-person per-WP monthly hours; derived forecast EUR flows into a new `Plan-Forecast` sheet and into parallel monitoring columns on Actuals, Summary, Trends; Capacity gains an actual-FTE twin view.

**Architecture:** Stateless sheet builders mirroring existing patterns (`capacity.py`, `plan.py`). Hour-to-EUR chain anchored in a single new Config constant `hours_per_fte_month = 143` (named range). Forecast EUR leaves live on `Actuals-Hours` as an outline-grouped collapsible column band — every EUR traceable to one cell. `Plan-Forecast` aggregates leaves into WP × month. Existing Actuals/Summary/Trends columns duplicated with parallel "Forecast" variants; CF thresholds mirror Actual-EUR rules exactly. Reader round-trips blue input hours on rebuild (like Capacity FTE and Actuals EUR today).

**Tech Stack:** Python 3.11+, openpyxl, pytest, pyright. No new runtime dependencies.

**Design source:** `docs/plans/2026-04-13-hours-timelog-design.md` (v1 only; v2–v4 out of scope).

**Critical invariants to uphold:**

- Rule 1: Excel formulas, never Python-computed values.
- Rule 2: `pytest tests/ -v` passes after every task.
- Rule 3: Font colors — Blue=input, Black=formula, Green=cross-sheet.
- Rule 4: Inactive months greyed and empty.
- Rule 5: `project_start` + `fte_rate` named ranges intact (never rebind).
- Rule 6: `_META` updated when structure changes.

---

## Staging

16 tasks. **Stage 1: tasks 1–7** (Config constant + Actuals-Hours builder + forecast leaves + Plan-Forecast + Actuals monitoring columns). Stage boundary: `Actuals-Hours` and `Plan-Forecast` exist, forecast monitoring shows on `Actuals`, all tests pass. **Stage 2: tasks 8–16** (reconciliation named range, Capacity twin view, Summary/Trends extensions, Reader/rebuild, _META, end-to-end tests).

---

## Pre-flight (all tasks)

Before every task: `git status` clean, on `main`. After every task: `pytest tests/ -v`, `pyright`, commit.

---

## Task 1: Add `hours_per_fte_month` Config constant and named range

**Files:**
- Modify: `budget_planner/config_sheet.py` (around line 639, `param_defs` list; around line 709, named-range registration)
- Test: `tests/test_config_sheet.py` (new test function)

**Why first:** every downstream formula references this named range. Nothing else compiles without it.

**Step 1: Write the failing test**

Add to `tests/test_config_sheet.py`:

```python
def test_hours_per_fte_month_constant_and_named_range(tmp_path):
    from budget_planner.workbook import init_workbook
    from openpyxl import load_workbook

    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))
    wb = load_workbook(str(path))

    assert "hours_per_fte_month" in wb.defined_names
    target = wb.defined_names["hours_per_fte_month"].attr_text
    assert target.startswith("Config!")

    sheet, cell_ref = target.split("!")
    cfg = wb["Config"]
    assert cfg[cell_ref.replace("$", "")].value == 143

    # Label in col A on the same row
    row = int(cell_ref.replace("$", "").lstrip("B"))
    assert cfg.cell(row=row, column=1).value == "Hours per FTE Month"
```

**Step 2: Run — expect fail**

`pytest tests/test_config_sheet.py::test_hours_per_fte_month_constant_and_named_range -v`

Expected: `KeyError: 'hours_per_fte_month'` or similar.

**Step 3: Implement**

In `budget_planner/config_sheet.py`, locate `param_defs` list (~line 639). Add after the EU ceiling entry:

```python
("Hours per FTE Month", 143, "0", None),
```

Locate the named-range registration block (~line 702–710). After the existing `DefinedName` adds for `project_start` and `fte_rate`, add:

```python
hours_row = param_rows["Hours per FTE Month"]
dn_hours = DefinedName(
    "hours_per_fte_month",
    attr_text=f"{SHEET_CONFIG}!$B${hours_row}",
)
wb.defined_names.add(dn_hours)
```

**Step 4: Run — expect pass**

`pytest tests/test_config_sheet.py::test_hours_per_fte_month_constant_and_named_range -v`

**Step 5: Sanity check existing tests**

`pytest tests/ -v` — expect all green. If any meta-test hard-coded the Config row count, update it to match.

**Step 6: Commit**

```bash
git add budget_planner/config_sheet.py tests/test_config_sheet.py
git commit -m "feat(config): add hours_per_fte_month constant and named range"
```

---

## Task 2: Extract shared person-section geometry helpers

**Files:**
- Modify: `budget_planner/capacity.py` (extract helpers from `_write_person_grid`)
- Create: `budget_planner/person_section.py`
- Test: `tests/test_person_section.py`

**Why:** `Actuals-Hours` mirrors `Capacity` geometry 1:1 (per design §1). Copy-paste would violate DRY and rot independently. Extract once before writing the second consumer.

**Minimum surface to extract** (identify in `capacity.py::_write_person_grid`):

- A function returning per-person row layout: `(header_row, wp_rows: dict[wp_id, row], external_rows: dict[name, row], total_row, remaining_row, eu_total_row)` given a starting row, a person record, their assignments, and externals.
- A function to apply active-month greying across a row given a WP schedule.

**Step 1: Write tests**

`tests/test_person_section.py`:

```python
from budget_planner.person_section import plan_person_layout

def test_plan_person_layout_assigns_sequential_rows():
    layout = plan_person_layout(
        start_row=10,
        wp_ids=["WP1", "WP2"],
        external_names=["TeachingX"],
        include_remaining=True,
        include_eu_total=True,
    )
    assert layout.header_row == 10
    assert layout.wp_rows == {"WP1": 11, "WP2": 12}
    assert layout.external_rows == {"TeachingX": 13}
    assert layout.total_row == 14
    assert layout.remaining_row == 15
    assert layout.eu_total_row == 16
    assert layout.next_row == 17


def test_plan_person_layout_without_optional_rows():
    layout = plan_person_layout(
        start_row=5,
        wp_ids=["WP1"],
        external_names=[],
        include_remaining=False,
        include_eu_total=False,
    )
    assert layout.total_row == 7
    assert layout.remaining_row is None
    assert layout.eu_total_row is None
    assert layout.next_row == 8
```

**Step 2: Run — expect fail**

`pytest tests/test_person_section.py -v` — `ModuleNotFoundError`.

**Step 3: Implement `budget_planner/person_section.py`**

```python
from dataclasses import dataclass


@dataclass
class PersonLayout:
    header_row: int
    wp_rows: dict[str, int]
    external_rows: dict[str, int]
    total_row: int
    remaining_row: int | None
    eu_total_row: int | None
    next_row: int


def plan_person_layout(
    *,
    start_row: int,
    wp_ids: list[str],
    external_names: list[str],
    include_remaining: bool,
    include_eu_total: bool,
) -> PersonLayout:
    r = start_row
    header_row = r
    r += 1
    wp_rows = {wp: r + i for i, wp in enumerate(wp_ids)}
    r += len(wp_ids)
    external_rows = {name: r + i for i, name in enumerate(external_names)}
    r += len(external_names)
    total_row = r
    r += 1
    remaining_row = r if include_remaining else None
    if include_remaining:
        r += 1
    eu_total_row = r if include_eu_total else None
    if include_eu_total:
        r += 1
    return PersonLayout(
        header_row=header_row,
        wp_rows=wp_rows,
        external_rows=external_rows,
        total_row=total_row,
        remaining_row=remaining_row,
        eu_total_row=eu_total_row,
        next_row=r,
    )
```

**Step 4: Refactor `capacity.py` to use `plan_person_layout`**

Inside `_write_person_grid`, replace inline row-counting with a call to `plan_person_layout(include_remaining=True, include_eu_total=True, ...)` and drive all subsequent cell writes from the returned dataclass fields. No behavior change — same rows, same formulas.

**Step 5: Run — expect pass**

`pytest tests/ -v` — all tests still green (Capacity round-trip proves the refactor preserved geometry).

**Step 6: Commit**

```bash
git add budget_planner/person_section.py budget_planner/capacity.py tests/test_person_section.py
git commit -m "refactor(capacity): extract plan_person_layout for reuse by Actuals-Hours"
```

---

## Task 3: Build `Actuals-Hours` sheet skeleton (hours input, no forecast EUR yet)

**Files:**
- Create: `budget_planner/actuals_hours.py`
- Create: `tests/test_actuals_hours.py`
- Modify: `budget_planner/workbook.py` (`SHEET_*` constants; `_build_all_sheets` insertion after Capacity, before Plan stub on line ~228 — actually after Actuals per design §7: Plan → Capacity → Actuals-Hours → Actuals → Summary. But Actuals-Hours depends on nothing that Actuals provides, so **insert right after Actuals (current line ~241)** to preserve tab-order design §7 recommendation)

**Scope of this task:** blue input hour cells only. Per-person sections mirror Capacity (same WPs, same externals, same months). Inactive months greyed. Per-person TOTAL hours (programme WPs + all) and EU TOTAL hours rows. **No forecast EUR columns yet** (Task 4).

**Step 1: Write the failing tests**

`tests/test_actuals_hours.py`:

```python
import pytest
from openpyxl import load_workbook
from budget_planner.workbook import init_workbook


@pytest.fixture
def wb(tmp_path):
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), demo=True)
    return load_workbook(str(path))


def test_actuals_hours_sheet_exists(wb):
    assert "Actuals-Hours" in wb.sheetnames


def test_actuals_hours_tab_order_after_actuals(wb):
    names = wb.sheetnames
    assert names.index("Actuals-Hours") == names.index("Actuals") + 1


def test_actuals_hours_person_sections_match_capacity(wb):
    cap = wb["Capacity"]
    hrs = wb["Actuals-Hours"]
    # Same first-person-group row and same total months
    # (Placeholder: replace with _META-driven assertion or exported layout dict)
    assert hrs.max_row >= cap.max_row - 5  # similar shape; tighten after implementation


def test_actuals_hours_inactive_months_are_empty_and_grey(wb):
    from budget_planner.styles import GREY_FILL
    hrs = wb["Actuals-Hours"]
    # Find a WP row with a known inactive month (demo has WP1 active M1-M12, WP2 M7-M24)
    # Pick any cell that must be inactive and assert empty + grey fill
    # Detailed assertion deferred to integration test after layout dict is exposed


def test_actuals_hours_hour_input_cells_are_blue(wb):
    from budget_planner.styles import BLUE_INPUT
    hrs = wb["Actuals-Hours"]
    # Sample a known active cell (first WP row, first month column) — assert blue font
    # Detailed assertion: pick first person group header, first WP row, first active month
```

Flesh out cell-level assertions once the builder exists (Step 3 produces a layout dict you can drive the tests from).

**Step 2: Run — expect fail**

`pytest tests/test_actuals_hours.py -v` — sheet not found.

**Step 3: Implement `budget_planner/actuals_hours.py`**

Mirror `capacity.py::build_capacity_sheet` signature. Key differences:

- Unit is **hours** per cell, not FTE. Cell format: `0` (integer hours) or `#,##0` — no decimals for input.
- No distribution curves — cells start **blank** (blue input), user fills at month-end.
- `TOTAL` row: SUM of all data rows (hours per month).
- `EU TOTAL` row: SUM of programme-WP rows + EU-flagged external rows (hours per month). Same EU ceiling CF pattern as Capacity, but comparing **annual hours totals** against `eu_annual_ceiling_pm × hours_per_fte_month`.
- Do **not** add a `REMAINING` row (hours remaining vs. Max FTE × hours_per_fte_month is a Capacity concern; Actuals-Hours is a record, not a budget).

Signature:

```python
def build_actuals_hours_sheet(
    ws,
    team_layout,
    *,
    project_params,
    work_packages,
    wp_schedule,
    team,
    external_commitments,
    assignments,
    config_first_wp_row,
    config_layout,
    preserved: dict[str, dict[str, list[float | None]]] | None = None,
) -> dict:
    ...
```

Returns an `actuals_hours_layout` dict analogous to `capacity_layout`: `{person_groups: [...], month_header_row, year_header_row, ...}`. Each person group entry: `{person_name, header_row, wp_rows: {wp_id: row}, external_rows: {name: row}, total_row, eu_total_row}`.

Implementation pattern:

1. Title row + year header + month header (copy from `capacity.py`).
2. For each person in `team`, call `plan_person_layout(include_remaining=False, include_eu_total=True, ...)` to allocate rows.
3. Write row headers (person header as merged `=Config!A<roster_row>`, WP ID/Name as green cross-refs to Config).
4. For each person × WP × month: if month in `wp_active_months(wp_id, schedule)`, write a blue-input empty cell (format `#,##0`); else GREY_FILL + empty + protected.
5. Write TOTAL row: `=SUM(<data cells this column>)` black formula.
6. Write EU TOTAL row: `=SUM(<programme-WP cells + EU-flagged externals this column>)` black formula.
7. Apply EU ceiling CF per calendar year on EU TOTAL row: red if `annual_hours > eu_annual_ceiling_pm × hours_per_fte_month`, amber if within one hours-per-FTE unit.
8. Apply `preserved` data if provided (blue-input cells overwritten with saved hours).

**Step 4: Wire into workbook build order**

In `budget_planner/workbook.py::_build_all_sheets`, after Actuals is built (~line 241), add:

```python
from .actuals_hours import build_actuals_hours_sheet

actuals_hours_ws = wb.create_sheet(SHEET_ACTUALS_HOURS)
actuals_hours_layout = build_actuals_hours_sheet(
    actuals_hours_ws, team_layout,
    project_params=project_params,
    work_packages=work_packages,
    wp_schedule=wp_schedule,
    team=team,
    external_commitments=external_commitments,
    assignments=assignments,
    config_first_wp_row=config_layout["wp_first_row"],
    config_layout=config_layout,
    preserved=preserved_actuals_hours,  # None on init; populated on rebuild (Task 13)
)
```

Add `SHEET_ACTUALS_HOURS = "Actuals-Hours"` near other `SHEET_*` constants (top of `workbook.py`). Thread `preserved_actuals_hours=None` through `init_workbook`; real plumbing comes in Task 13.

**Step 5: Run tests — expect pass**

`pytest tests/test_actuals_hours.py tests/test_workbook.py -v`. Then full suite: `pytest tests/ -v`.

**Step 6: Commit**

```bash
git add budget_planner/actuals_hours.py budget_planner/workbook.py tests/test_actuals_hours.py
git commit -m "feat(actuals-hours): add sheet with per-person hour inputs and EU ceiling CF"
```

---

## Task 4: Add collapsible forecast-EUR leaf columns on `Actuals-Hours`

**Files:**
- Modify: `budget_planner/actuals_hours.py`
- Modify: `tests/test_actuals_hours.py`

**Design:** for each data cell (person-WP-month, hour input), a sibling EUR cell sits in a parallel column band to the right of the monthly-hours block. Band is outline-grouped (`ws.column_dimensions[col].outline_level = 1`), collapsed by default via `ws.sheet_properties.outlinePr.summaryRight = True` and `ws.column_dimensions.group(start, end, outline_level=1, hidden=True)` per [openpyxl docs].

Per CLAUDE.md rule noted in recent commit `9b04111`: **prefer outline groups over `hidden`**. Use `column_dimensions.group(..., hidden=False, outline_level=1)` so the user sees the outline toggle and can expand.

**Leaf formula** (black text — derived):

```
= IF(<hour_cell>="", "", <hour_cell> / hours_per_fte_month * VLOOKUP(<person_name_cell>, roster_lookup, <rate_col>, FALSE))
```

`<rate_col>` is the column index in `roster_lookup` holding the per-person rate. Per the design map: `roster_lookup` is `Config!$A$<first>:$B$<last>` today — **only 2 cols** (name, Max FTE). Rate per person is not yet in `roster_lookup`. Two options:

- **Option A — extend `roster_lookup` to include rate:** add a Grade column and a Rate column to the roster, widening `roster_lookup` from B to D (e.g.). Rate per person = `VLOOKUP(grade, tbl_GradeRates, 2, FALSE)`. Update existing Assignments Balance formula (currently `VLOOKUP(name, roster_lookup, 2, FALSE)` for Max FTE — col 2 must stay Max FTE; put rate at col 4).
- **Option B — separate `person_rate_lookup` named range:** new 2-col range (person → rate), built next to roster. Keeps `roster_lookup` unchanged.

**Choose Option A.** Reason: one lookup table per entity; roster is the authoritative person directory. Wider is fine; it's flat data.

**Step 1: Write the failing tests**

Extend `tests/test_actuals_hours.py`:

```python
def test_actuals_hours_forecast_eur_leaves_exist_and_grouped(wb):
    hrs = wb["Actuals-Hours"]
    # For one known person-WP-month cell, sibling EUR cell exists
    # EUR cell has a formula starting with "=IF("
    # EUR column has outline_level >= 1

def test_forecast_eur_leaf_formula_uses_hours_per_fte_month_and_roster_lookup(wb):
    hrs = wb["Actuals-Hours"]
    # Find any active hour input cell; locate its sibling EUR cell
    # Assert formula contains "hours_per_fte_month" and "roster_lookup"

def test_forecast_eur_leaf_blank_when_hour_blank(tmp_path):
    from budget_planner.workbook import init_workbook
    import subprocess
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), demo=False)  # no demo hours → all blank
    # Open, run recalc, verify EUR leaves are "" (blank)
```

**Step 2: Run — expect fail**

**Step 3: Implement**

In `build_actuals_hours_sheet`:

- After the main hour-input column block ends, compute `forecast_col_start = month_col_end(months) + N` where N is the number of validation cols on Actuals-Hours (sum/total only — no budget/variance on this sheet). Reuse `val_columns(months)` for consistency.
- Actually simpler: put forecast EUR in a parallel **right-side band** starting at `forecast_col_start = val_col_last + 2` (one blank gutter column). Band spans `months` columns.
- Each month column in the band: outline_level=1, `ws.column_dimensions[col].outline_level = 1`. Set `ws.sheet_properties.outlinePr.summaryRight = True`.
- For each person × WP × month active cell, write EUR formula (see above) into the corresponding forecast column.
- For inactive months: GREY_FILL, empty.
- Per-person TOTAL EUR row and EU TOTAL EUR row in the band: `=SUM(<wp-eur-cells>)` per month.
- Header band: a single merged row above the forecast month columns reading "FORECAST EUR (from hours)" with light green-blue fill to signal derived-from-hours.

**Step 4: Extend `roster_lookup` to include rate (Option A)**

In `budget_planner/config_sheet.py::_write_roster_section`:

- Current roster cols: A=Name, B=Max FTE. Add C=Grade (blue input), D=Rate EUR/PM (black formula: `=VLOOKUP(C<row>, tbl_GradeRates, 2, FALSE)`).
- Grade default: first grade in `GRADE_RATES` (seed value for new persons). User can edit.
- Update `roster_lookup` named-range `attr_text` from `$A$<first>:$B$<last>` to `$A$<first>:$D$<last>`.
- Existing Assignments Balance formula (`VLOOKUP(name, roster_lookup, 2, FALSE)`) still works (col 2 = Max FTE).

Update `budget_planner/reader.py::read_team_from_config` to read the new Grade column so rebuild preserves it.

Update any test that asserts the roster table width (grep for `roster_lookup` in `tests/`).

**Step 5: Run tests — expect pass**

`pytest tests/ -v`. Be prepared for `test_config_sheet.py` and `test_reader_v2.py` failures — fix roster width assertions.

**Step 6: Commit**

```bash
git add budget_planner/actuals_hours.py budget_planner/config_sheet.py budget_planner/reader.py tests/test_actuals_hours.py tests/test_config_sheet.py tests/test_reader_v2.py
git commit -m "feat(actuals-hours): add collapsible forecast EUR leaf columns; extend roster_lookup with grade+rate"
```

---

## Task 5: Build `Plan-Forecast` sheet

**Files:**
- Create: `budget_planner/plan_forecast.py`
- Create: `tests/test_plan_forecast.py`
- Modify: `budget_planner/workbook.py` (build order — after Actuals-Hours, before Summary)

**Structure:** mirror `Plan` exactly — WP × month grid, compact header identical to Plan. Per-cell formula:

```
= SUM(<forecast-EUR-leaf-cells on Actuals-Hours for this WP, this month, across all persons assigned to this WP>)
```

Aggregation source: `actuals_hours_layout["person_groups"]` gives each person's `wp_rows` and the forecast column band offsets.

**Step 1: Write failing tests**

`tests/test_plan_forecast.py`:

```python
def test_plan_forecast_sheet_exists(wb):
    assert "Plan-Forecast" in wb.sheetnames

def test_plan_forecast_mirrors_plan_shape(wb):
    plan = wb["Plan"]
    pf = wb["Plan-Forecast"]
    assert pf.max_column >= plan.max_column - 1  # allow small delta

def test_plan_forecast_cell_sums_actuals_hours_leaves(wb):
    pf = wb["Plan-Forecast"]
    # Pick WP1 × M1 cell; assert formula starts with "=SUM(" and contains "Actuals-Hours!"

def test_plan_forecast_inactive_months_empty_and_grey(wb):
    # Same contract as Plan
```

**Step 2: Run — expect fail**

**Step 3: Implement `plan_forecast.py`**

Copy `plan.py` structure:

- Compact header (Title / Start / Ceiling).
- "FORECAST EUR (from hours)" section header.
- Year + month header rows.
- One row per WP, green cross-ref to Config for ID/Name.
- Monthly cell formulas: green `=SUM(Actuals-Hours!<leaf1>, Actuals-Hours!<leaf2>, ...)` built from `actuals_hours_layout["person_groups"]` by iterating persons assigned to the WP and picking their forecast-EUR column for this month.
- Validation cols: Sum (total forecast EUR per WP), Ceiling (green from Config), Variance (`=Ceiling − Sum`, red CF if negative).
- TOTAL row at bottom.

**Step 4: Wire into build order**

In `workbook.py`, after `build_actuals_hours_sheet`:

```python
from .plan_forecast import build_plan_forecast_sheet
plan_forecast_ws = wb.create_sheet(SHEET_PLAN_FORECAST)
plan_forecast_layout = build_plan_forecast_sheet(
    plan_forecast_ws,
    project_params=project_params,
    work_packages=work_packages,
    wp_schedule=wp_schedule,
    team=team,
    assignments=assignments,
    config_first_wp_row=config_layout["wp_first_row"],
    actuals_hours_layout=actuals_hours_layout,
)
```

Add `SHEET_PLAN_FORECAST = "Plan-Forecast"`. Tab order: after Actuals-Hours, before Summary.

**Step 5: Run — expect pass**

`pytest tests/ -v`.

**Step 6: Commit**

```bash
git add budget_planner/plan_forecast.py budget_planner/workbook.py tests/test_plan_forecast.py
git commit -m "feat(plan-forecast): aggregate forecast EUR leaves into WP × month grid"
```

---

## Task 6: Add parallel forecast monitoring columns to `Actuals` sheet

**Files:**
- Modify: `budget_planner/actuals.py`
- Modify: `tests/test_conditional_formatting.py`
- Add: `tests/test_actuals_forecast_block.py`

**Scope:** add a grouped column block to the right of existing Actual-EUR monitoring columns. Columns:

- Cumulative Forecast EUR: `=SUMPRODUCT((COLUMN(<pf_range>) − COLUMN(<pf_first>) + 1 ≤ current_month) * <pf_range>)` — same cumulative shape as Actual EUR but reading from `Plan-Forecast` for this WP row.
- Forecast Variance vs Ceiling: `=Ceiling − Cum Forecast EUR`.
- Forecast % Ceiling: `=Cum Forecast EUR / Ceiling`.
- Forecast Breach In: same formula shape as Actual Breach In, fed by Forecast EUR.

**CF rules (mirror Actual EUR exactly):**

- Red: `% Ceiling > 100%` or `Breach In < 6 months`.
- Amber: `% Ceiling >= 85%`.

Put the forecast block under an outline group (level 1), with a header band "FORECAST MONITORING" above it.

**Step 1: Write tests** — assert 4 new columns exist per WP row with formulas referencing `Plan-Forecast!`, outline level, and CF rules present.

**Step 2: Run — expect fail**

**Step 3: Implement** — extend `actuals.py::build_actuals_sheet` to add the forecast block after the existing val_columns. Reuse `val_columns()` pattern; introduce a sibling `forecast_val_columns()` helper in `layout.py` that returns the same offsets shifted by `+ len(val_columns)` to keep the right-edge layout tidy.

**Step 4: Run — expect pass**

**Step 5: Commit**

```bash
git add budget_planner/actuals.py budget_planner/layout.py tests/test_actuals_forecast_block.py tests/test_conditional_formatting.py
git commit -m "feat(actuals): add parallel forecast monitoring columns with mirrored CF"
```

---

## Task 7: Stage 1 checkpoint — full regression + _META partial update

**Files:**
- Modify: `budget_planner/meta.py`
- Run: full test suite

**Step 1:** Run `pytest tests/ -v`. All green.

**Step 2:** Run `pyright budget_planner/`. Zero errors.

**Step 3:** Open `output/budget_plan.xlsx` (after `python -m budget_planner init --demo && python scripts/recalc.py output/budget_plan.xlsx`) and spot-check in LibreOffice:

- `Actuals-Hours` tab exists, tab color matches family (actuals).
- Entering `143` hours in a person × WP × month cell produces EUR equal to that person's rate in the sibling forecast-EUR column.
- `Plan-Forecast` cell for that WP × month shows the same EUR.
- `Actuals!<WP row>` forecast block shows the EUR as part of Cumulative Forecast EUR.

**Step 4:** Update `_META` sheet text (`budget_planner/meta.py`) with minimum entries for `Actuals-Hours` and `Plan-Forecast`: row positions, purpose, font/color conventions, cross-sheet dependencies. Full `_META` rewrite deferred to Task 15 after Stage 2 lands the remaining extensions.

**Step 5: Commit**

```bash
git add budget_planner/meta.py
git commit -m "docs(meta): add partial entries for Actuals-Hours and Plan-Forecast (stage 1 checkpoint)"
```

**Step 6: Stage 1 complete.** Save `memory/project_state.md` note: "hours-timelog Stage 1 landed — tasks 1–7; resume with Task 8 in fresh session."

---

## Task 8: Add `eur_actuals_through_month` reconciliation named range

**Files:**
- Modify: `budget_planner/config_sheet.py` (new parameter)
- Modify: `budget_planner/actuals.py` (visual emphasis)
- Modify: `budget_planner/actuals_hours.py` (visual emphasis)
- Test: `tests/test_config_sheet.py`, `tests/test_actuals_forecast_block.py`

**Scope:** one Config param `EUR Actuals Through Month` (integer, default = `current_month`; user bumps it when finance closes a month). Named range `eur_actuals_through_month` → `Config!$B$<row>`. Visual behavior on Actuals:

- Months ≤ `eur_actuals_through_month`: Actual EUR normal styling; Forecast EUR cells light-grey fill (de-emphasized — reference only).
- Months > `eur_actuals_through_month`: Forecast EUR normal styling; Actual EUR cells light-grey fill (blank-or-stale — soon to be filled).

Implement via CF on the monthly Actual EUR cells and on the leaf Forecast EUR cells: fill light-grey if the month column index is on the wrong side of `eur_actuals_through_month`.

**Step 1–6: TDD cycle.** Test asserts named range exists, attr_text points to Config, and CF rules exist on both sheets. Implement; run; commit.

```bash
git commit -m "feat(reconciliation): add eur_actuals_through_month named range and visual emphasis"
```

---

## Task 9: Capacity twin view — Actual FTE, Variance FTE, Utilisation %, Actual EU TOTAL ceiling check

**Files:**
- Modify: `budget_planner/capacity.py`
- Modify: `budget_planner/layout.py` (new constants)
- Add: `tests/test_capacity_actual_view.py`

**Scope:** add a right-side grouped band on `Capacity` — per cell, three derived rows:

- **Actual FTE**: `=Actuals-Hours!<same_cell> / hours_per_fte_month` (green cross-sheet).
- **Variance FTE**: `=Planned FTE − Actual FTE`. Red CF if `ABS > 0.2`.
- **Utilisation %** per person per month: `=Actual TOTAL / <Max FTE>`. Red if ≥100%, amber if ≥80%. Mirrors planned-side CF.

Also add per-person **Actual TOTAL** and **Actual EU TOTAL** rows. EU annual ceiling validation (per-calendar-year SUMPRODUCT) runs on actual hours converted to PM: red if `annual_actual_PM > eu_annual_ceiling_pm`.

Threshold `CAPACITY_VARIANCE_FTE_RED = 0.2` in `layout.py`.

**Step 1–6: TDD cycle.** 

```bash
git commit -m "feat(capacity): add actual-FTE twin view with variance, utilisation, and EU ceiling check on actuals"
```

---

## Task 10: Summary — Forecast EUR column per year + PROJECT TOTAL

**Files:**
- Modify: `budget_planner/summary.py`
- Modify: `tests/test_conditional_formatting.py`, `tests/test_summary.py` (or create if absent)

**Scope:** add a 5th column per calendar-year group: **Forecast EUR** (green SUMPRODUCT over `Plan-Forecast` monthly cells falling in the year). Same CF as Actual EUR (red if Forecast < 0.8 × Plan EUR for the year, etc. — match the existing Actual EUR CF rule text exactly). Replicate in PROJECT TOTAL group.

Update `COLS_PER_YEAR` (currently 4 → 5) and `TOTAL_COLS` consistently. Grep for these constants first; propagate.

**Step 1–6: TDD cycle.**

```bash
git commit -m "feat(summary): add Forecast EUR column per year with variance CF"
```

---

## Task 11: Trends — Cumulative Forecast EUR line

**Files:**
- Modify: `budget_planner/trends.py`
- Test: `tests/test_trends.py`

**Scope:** add a 5th series to the Cumulative EUR LineChart: "Cumulative Forecast" (amber dashed). Data source: helper row summing `Plan-Forecast` monthly totals cumulatively. Insert series after "Cumulative Actual", before "Ceiling".

**Step 1–6: TDD cycle.**

```bash
git commit -m "feat(trends): add Cumulative Forecast EUR series to LineChart"
```

---

## Task 12: Trends — Person Load actual-side columns

**Files:**
- Modify: `budget_planner/trends.py`
- Test: `tests/test_trends.py`

**Scope:** Person Load table gains three columns:

- **Actual Avg FTE (to date)**: `= AVERAGEIFS` over Capacity Actual-FTE row (new, from Task 9) up to `current_month`.
- **Actual Peak FTE**: `= MAX` over same range.
- **Planned − Actual Variance**: text status (OVER/WATCH/OK) via nested IF, matching the planned-side status column.

**Step 1–6: TDD cycle.**

```bash
git commit -m "feat(trends): add actual-side columns to Person Load"
```

---

## Task 13: Reader — round-trip hours on rebuild

**Files:**
- Modify: `budget_planner/reader.py` (new function `read_actuals_hours_sheet`)
- Modify: `budget_planner/workbook.py::rebuild_workbook` (plumb preserved hours)
- Add: `tests/test_rebuild_actuals_hours.py`

**Scope:** extract the blue-input hours per person × WP × month on rebuild, pass through to `build_actuals_hours_sheet(preserved=...)`. Same pattern as `read_actuals_sheet` and `read_capacity_sheet`.

The reader must tolerate:
- The collapsible Forecast EUR band (skip columns with `outline_level >= 1` — those are derived, recalculated on rebuild).
- Blank cells (user hasn't logged that month yet).
- Internal blank rows inside a person section (design `docs/plans/archived/2026-04-12-rebuild-shape-drift-design.md`: use named ranges / header anchors, tolerate gaps).

**Step 1:** Failing test — init workbook, write demo hours into several cells, rebuild, assert hours preserved.

**Step 2–5:** Implement `read_actuals_hours_sheet(ws) → dict[person_name, dict[wp_id, list[float | None]]]`. Anchor on person header rows (merged A:B with `=Config!A<roster_row>`). Walk WP rows under each person until next person header or TOTAL.

**Step 6:**

```bash
git commit -m "feat(reader): round-trip Actuals-Hours data on rebuild"
```

---

## Task 14: End-to-end rebuild test for hours → forecast EUR chain

**Files:**
- Add: `tests/test_rebuild_hours_forecast_e2e.py`

**Scope:** integration test closing the loop:

1. `init_workbook(path, demo=True)`.
2. Open, write hours into 3 person-WP-month cells.
3. Write a grade rate edit in Config.
4. Run `rebuild_workbook(path)`.
5. Re-open; run `scripts/recalc.py path`.
6. Assert:
   - Hours preserved.
   - Forecast EUR leaf = hours / hours_per_fte_month × new rate.
   - `Plan-Forecast` WP × month cell sums the leaves correctly.
   - `Actuals` forecast block shows cumulative Forecast EUR equal to the sum of the three leaves (assuming all in months ≤ current_month).
   - `Summary` Forecast EUR column for the year includes them.

**Step 1: Write the test.**

**Step 2: Run — expect pass** (all prior tasks have made this possible).

**Step 3: Commit**

```bash
git add tests/test_rebuild_hours_forecast_e2e.py
git commit -m "test(e2e): rebuild preserves hours and propagates through forecast chain"
```

---

## Task 15: Update `_META` sheet with full v1 documentation

**Files:**
- Modify: `budget_planner/meta.py`

**Scope:** per Critical Rule 6, `_META` is primary documentation for future AI agents. Rewrite sections:

- **Sheet index:** add `Actuals-Hours` and `Plan-Forecast` with purposes, tab colors, data-flow arrows.
- **Actuals-Hours section:** row positions (year_header, month_header, first/last person group), person-section layout, hour-input convention, forecast-EUR band convention (outline level 1, collapsible, recalculated on rebuild).
- **Plan-Forecast section:** row positions, aggregation formula shape, source cells.
- **Actuals section (extend):** forecast monitoring columns — positions, formulas, CF thresholds, reconciliation via `eur_actuals_through_month`.
- **Capacity section (extend):** Actual FTE / Variance FTE / Utilisation % band — positions, formulas, CF thresholds.
- **Summary section (extend):** Forecast EUR column per year.
- **Trends section (extend):** Cumulative Forecast series; Person Load actual-side columns.
- **Named ranges list:** add `hours_per_fte_month`, `eur_actuals_through_month`.
- **Build order:** update to show `Config → Capacity → Plan → Actuals → Actuals-Hours → Plan-Forecast → Summary → Trends → _META`.
- **Data flow diagram:** update the ASCII/prose diagram to include the hours → forecast EUR chain.

**Step 1:** Read `budget_planner/meta.py` fully (~500 lines) to understand the existing text-builder pattern.

**Step 2:** Extend the text lists with new sections in the right order.

**Step 3:** Run `pytest tests/test_meta.py -v` — adjust any hard-coded row-offset assertions.

**Step 4: Commit**

```bash
git add budget_planner/meta.py tests/test_meta.py
git commit -m "docs(meta): full v1 documentation for Actuals-Hours, Plan-Forecast, and all extensions"
```

---

## Task 16: Final verification + checklist close-out

**Files:** none; verification only.

**Step 1:** Full test suite: `pytest tests/ -v`. Expect all green, ~350+ tests (rose from ~328 baseline).

**Step 2:** `pyright budget_planner/`. Zero errors.

**Step 3:** Fresh demo workbook: `python -m budget_planner init --demo && python scripts/recalc.py output/budget_plan.xlsx`. Open in LibreOffice. Smoke tests (critical-rules verification):

- Rule 1: every derived cell is a formula (check Summary Forecast EUR — not a value).
- Rule 3: Actuals-Hours hour cells are blue; forecast-EUR leaves are black; Plan-Forecast cells are green.
- Rule 4: Actuals-Hours inactive months are greyed and empty on both hours band and forecast-EUR band.
- Rule 5: `project_start` and `fte_rate` still bound.
- Rule 6: `_META` shows all new sheets and columns.

**Step 4:** Rebuild cycle: edit a hours cell, `python -m budget_planner rebuild`, verify the hour preserved, forecast EUR updated, Plan-Forecast aggregates correctly.

**Step 5:** Snapshot/restore works across new sheets: `python -m budget_planner --snapshot output/budget_plan.xlsx && python -m budget_planner --restore output/budget_plan.xlsx`. (If the snapshot logic is shape-agnostic, no code change needed; otherwise file a follow-up — v1 can skip scenario comparison for hours.)

**Step 6: Commit (if any fix needed)**

```bash
git commit -m "chore(v1): final verification and fixes"
```

**Done.** v1 complete.

---

## Waste analysis checkpoints

Before dispatching Stage 2, run `/kaizen:analyse --muda` against tasks 8–16 to catch:

- Unnecessary column bands (do we really need **both** Actual FTE and Variance FTE on Capacity, or does variance alone suffice?).
- Duplicate CF rules (Actuals forecast CF + Summary forecast CF + Trends forecast alarm — consolidate messaging where possible).
- Over-processing in _META (full rewrite vs. delta).

Apply simplifications before Stage 2 starts.

---

## Execution handoff

Plan complete and saved to `docs/plans/2026-04-13-hours-timelog-v1-plan.md`. Two execution options:

1. **Subagent-Driven (this session)** — dispatch fresh subagent per task, review between tasks, fast iteration. Stage 1 only (tasks 1–7).
2. **Parallel Session (separate)** — open new session with executing-plans; full v1 run in batches with checkpoints.

Which approach?
