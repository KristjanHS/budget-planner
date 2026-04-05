# P1a: 215-Day EU Ceiling Check — Implementation (Session 1 of 2)

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add per-person EU capacity tracking with annual 215-day ceiling validation (≈12 PM/year) across all EU grants.

**Architecture:** Extend Config with `eu_annual_ceiling_pm` parameter. Tag external commitments as EU/non-EU via `is_eu` flag. Add "EU TOTAL" summary row per person in Capacity (after REMAINING) that sums programme WP + EU-external monthly FTE. Apply per-calendar-year CF rules on EU TOTAL: Red >ceiling, Amber >ceiling−1.

**Tech Stack:** Python, openpyxl, pytest

**Key files:**
- `budget_planner/config.py` — seed data
- `budget_planner/config_sheet.py` — Config sheet builder
- `budget_planner/capacity.py` — Capacity sheet builder
- `budget_planner/reader.py` — round-trip reader
- `budget_planner/workbook.py` — orchestrator
- `budget_planner/layout.py` — constants (read-only, use `_calendar_year_groups`)

**Risk notes:**
- `reader.py:82` hardcodes `r = 8` for WP start row — MUST fix when Config gains a row
- Named ranges `project_start=$B$2` and `fte_rate=$B$4` are hardcoded — don't shift rows 1-4
- CF rule count in `test_conditional_formatting.py:55` will need updating after Task 6
- `test_config_sheet.py:73` asserts relative WP row positions — should survive the shift
- `workbook.py:238-243` builds project_params without `default_project_params()` — must add `eu_ceiling_pm` there too

**Session 2:** After this session is merged, run `docs/plans/2026-04-05-p1b-eu-ceiling-tests-docs.md` for rebuild safety-net tests, _META docs, and final verification.

---

## Task 1: Add EU ceiling constant to config seed data

**Files:**
- Modify: `budget_planner/config.py:17` (after DEFAULT_RATE)
- Modify: `budget_planner/config.py:106-113` (default_project_params)
- Modify: `budget_planner/workbook.py:238-243` (init_workbook project_params)

**Step 1: Add constant and update both project_params builders**

In `budget_planner/config.py`, add after line 17 (`DEFAULT_RATE = 7500`):
```python
EU_ANNUAL_CEILING_PM = 12.0              # Max PM per calendar year across all EU grants (≈215 days)
```

In `default_project_params()` (line 108), add the key:
```python
"eu_ceiling_pm": EU_ANNUAL_CEILING_PM,
```

In `budget_planner/workbook.py` `init_workbook()` (line 238), add:
```python
"eu_ceiling_pm": config.EU_ANNUAL_CEILING_PM,
```

**Step 2: Run tests to verify nothing breaks**

Run: `pytest tests/ -q`
Expected: All pass (no test depends on project_params keys yet)

**Step 3: Commit**

```bash
git add budget_planner/config.py budget_planner/workbook.py
git commit -m "feat(config): add EU_ANNUAL_CEILING_PM constant"
```

---

## Task 2: Add EU ceiling param to Config sheet builder

**Files:**
- Modify: `budget_planner/config_sheet.py:49-54` (param_defs)
- Modify: `budget_planner/config_sheet.py:64-68` (section positioning — make dynamic)
- Modify: `budget_planner/config_sheet.py:167-176` (named range + freeze_panes)
- Test: `tests/test_config_sheet.py`

**Step 1: Write the failing tests**

Add to `tests/test_config_sheet.py`:
```python
def test_config_eu_ceiling_param():
    """EU Annual Ceiling (PM) parameter row exists with default value."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    found = False
    for r in range(1, 10):
        if ws.cell(r, 1).value and "EU Annual Ceiling" in str(ws.cell(r, 1).value):
            assert ws.cell(r, 2).value == 12.0
            found = True
            break
    assert found, "EU Annual Ceiling param not found"


def test_config_eu_ceiling_named_range():
    """Named range eu_annual_ceiling_pm exists."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    names = set(wb.defined_names.keys())
    assert "eu_annual_ceiling_pm" in names
```

**Step 2: Run tests to verify they fail**

Run: `pytest tests/test_config_sheet.py::test_config_eu_ceiling_param tests/test_config_sheet.py::test_config_eu_ceiling_named_range -v`
Expected: FAIL (param and named range don't exist yet)

**Step 3: Implement — extend param_defs, dynamicize section positions, add named range**

In `budget_planner/config_sheet.py`, extend `param_defs` (line 49):
```python
param_defs = [
    ("Project Title", project_params["title"], None, None),
    ("Project Start", project_params["start"], "MMM YYYY", None),
    ("Project Months", project_params["months"], "0", None),
    ("FTE Rate (EUR)", project_params["rate"], "€#,##0", None),
    ("EU Annual Ceiling (PM)", project_params.get("eu_ceiling_pm", 12.0), "0.00", None),
]
```

Replace the hardcoded section positioning (lines 64-68):
```python
# Old:
# # Row 5 is empty separator (no action needed)
# section_row = 6
# header_row = 7
# first_data_row = 8

# New — dynamic:
section_row = len(param_defs) + 2   # gap row after last param
header_row = section_row + 1
first_data_row = header_row + 1
```

Add named range (after existing named ranges, ~line 172):
```python
ceiling_row = param_rows["EU Annual Ceiling (PM)"]
dn_ceiling = DefinedName("eu_annual_ceiling_pm",
                          attr_text=f"{SHEET_CONFIG}!$B${ceiling_row}")
wb.defined_names.add(dn_ceiling)
```

Update freeze_panes (~line 176):
```python
ws.freeze_panes = f"A{section_row}"
```

**Step 4: Run all tests**

Run: `pytest tests/ -q`
Expected: All pass. Watch for failures in test_config_sheet.py and test_reader (see Task 3).

**Step 5: Commit**

```bash
git add budget_planner/config_sheet.py tests/test_config_sheet.py
git commit -m "feat(config-sheet): add EU Annual Ceiling param row and named range"
```

---

## Task 3: Fix Config reader for new row layout

**Files:**
- Modify: `budget_planner/reader.py:59-98` (read_config_sheet)
- Test: `tests/test_config_sheet.py`

**Step 1: Write the failing test**

Add to `tests/test_config_sheet.py`:
```python
def test_config_round_trip_eu_ceiling():
    """EU ceiling param survives write → read round-trip."""
    from budget_planner.reader import read_config_sheet
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {**_sample_params(), "eu_ceiling_pm": 10.0}
    build_config_sheet(ws, wb, params, _sample_wps())
    read_params, read_wps = read_config_sheet(ws)
    assert read_params["eu_ceiling_pm"] == 10.0
    assert len(read_wps) == 2  # WPs still read correctly after row shift
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_config_sheet.py::test_config_round_trip_eu_ceiling -v`
Expected: FAIL — reader doesn't read eu_ceiling_pm and WP start row is wrong

**Step 3: Fix reader.py read_config_sheet**

In `budget_planner/reader.py`, update `read_config_sheet()` (lines 59-98):

Add ceiling to the validation labels:
```python
_expected_labels = [
    (1, "title", "Project Title"),
    (2, "start", "Project Start"),
    (3, "months", "Project Months"),
    (4, "rate", "FTE Rate"),
    (5, "ceiling", "EU Annual Ceiling"),
]
```

Add to project_params dict:
```python
project_params = {
    "title": ws.cell(1, 2).value,
    "start": ws.cell(2, 2).value,
    "months": int(ws.cell(3, 2).value or 60),
    "rate": float(ws.cell(4, 2).value or 7500),
    "eu_ceiling_pm": float(ws.cell(5, 2).value or 12.0),
}
```

Replace hardcoded WP start row (lines 81-83):
```python
# Old:
# # Work packages start at row 8 (row 6 = section header, row 7 = headers)
# r = 8

# New — dynamic:
wp_section = _find_header_row(ws, "WORK PACKAGES")
r = wp_section + 2  # skip section header + column headers
```

**Step 4: Run all tests**

Run: `pytest tests/ -q`
Expected: All pass (rebuild path also works since read_config_sheet returns eu_ceiling_pm)

**Step 5: Commit**

```bash
git add budget_planner/reader.py tests/test_config_sheet.py
git commit -m "fix(reader): read EU ceiling param, dynamicize WP start row"
```

---

## Task 4: Add is_eu flag to external commitments

**Files:**
- Modify: `budget_planner/config.py:48-51` (EXTERNAL_COMMITMENTS seed data)
- Modify: `budget_planner/capacity.py:169-188` (config block 2 columns)
- Modify: `budget_planner/reader.py:242-265` (read external commitments)
- Test: `tests/test_capacity.py`

**Step 1: Write the failing test**

Add to `tests/test_capacity.py`:
```python
def test_capacity_external_config_has_eu_column(cap_ws):
    """External commitments config block has EU Grant column."""
    for r in range(1, cap_ws.max_row + 1):
        val = cap_ws.cell(r, 1).value
        if val and "EXTERNAL COMMITMENTS" in str(val):
            # Column headers on next row
            header_row = r + 1
            assert cap_ws.cell(header_row, 6).value == "EU Grant"
            break
    else:
        pytest.fail("EXTERNAL COMMITMENTS section not found")
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_capacity.py::test_capacity_external_config_has_eu_column -v`
Expected: FAIL

**Step 3: Update seed data**

In `budget_planner/config.py`, update EXTERNAL_COMMITMENTS (line 48) and add a demo EU entry:
```python
EXTERNAL_COMMITMENTS = [
    {"person": "Alice", "project": "ProjectX",        "fte": 2.0, "start": 1, "end": 11, "is_eu": False},
    {"person": "Bob",   "project": "ProjectX",        "fte": 2.0, "start": 1, "end": 14, "is_eu": False},
    {"person": "Bob",   "project": "HorizonEU-Alpha", "fte": 2.0, "start": 1, "end": 7, "is_eu": True},
]
```

**Step 4: Update capacity.py config block 2**

In `budget_planner/capacity.py`, update the external commitments config block (~line 169):

Add "EU Grant" to column headers:
```python
ext_col_headers = ["Person", "Project", "Total FTE", "Start Month", "End Month", "EU Grant"]
```

Update the merge_end for the section header:
```python
write_section_header(ws, r, "EXTERNAL COMMITMENTS", merge_end=6)
```

In the data row loop (~line 181), add EU Grant value after end month:
```python
style_cell(ws, r, 6, "Yes" if entry.get("is_eu", False) else "No",
           font=BLUE_INPUT)
```

**Step 5: Update reader.py to read is_eu**

In `budget_planner/reader.py`, in the external commitments reading loop (~line 258), add:
```python
is_eu_val = ws.cell(r, 6).value
external_commitments.append({
    "person": str(person),
    "project": str(project or ""),
    "fte": float(fte or 0),
    "start": int(start or 1),
    "end": int(end or 1),
    "is_eu": str(is_eu_val).strip().upper() in ("YES", "TRUE", "1"),
})
```

**Step 6: Run all tests**

Run: `pytest tests/ -q`
Expected: All pass

**Step 7: Commit**

```bash
git add budget_planner/config.py budget_planner/capacity.py budget_planner/reader.py tests/test_capacity.py
git commit -m "feat(external): add is_eu flag to external commitments"
```

---

## Task 5: Add EU TOTAL row per person in Capacity

**Files:**
- Modify: `budget_planner/capacity.py:396-454` (after REMAINING row, before person_groups dict)
- Test: `tests/test_capacity.py`

**Step 1: Write the failing tests**

Add to `tests/test_capacity.py`:
```python
def test_capacity_person_has_eu_total_row(cap_ws, capacity_layout):
    """EU TOTAL row exists per person."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        eu_total_row = group.get("eu_total_row")
        assert eu_total_row is not None, f"{member['name']} missing eu_total_row"
        assert cap_ws.cell(row=eu_total_row, column=2).value == "EU TOTAL"


def test_capacity_eu_total_row_formula(cap_ws, capacity_layout):
    """EU TOTAL monthly cells have SUM formulas referencing WP rows."""
    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        eu_total_row = group["eu_total_row"]
        val = cap_ws.cell(row=eu_total_row, column=MONTH_COL_START).value
        # Should be a SUM formula (or =0 if no WP rows)
        assert isinstance(val, str) and val.startswith("="), (
            f"{name} EU TOTAL M1: expected formula, got {val}"
        )
```

**Step 2: Run tests to verify they fail**

Run: `pytest tests/test_capacity.py::test_capacity_person_has_eu_total_row tests/test_capacity.py::test_capacity_eu_total_row_formula -v`
Expected: FAIL

**Step 3: Implement EU TOTAL row**

In `budget_planner/capacity.py`, after the REMAINING row block (~line 415) and before the CF section (~line 417), add:

```python
        # EU TOTAL row — sum of programme WP rows + EU-flagged external rows
        eu_total_row = r
        style_cell(ws, r, 1, fill=LIGHT_BLUE_FILL)
        style_cell(ws, r, 2, "EU TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)

        # Build list of rows that count toward EU total
        eu_row_refs = list(wp_rows.values())  # All programme WPs are EU
        for proj, ext_row in external_rows.items():
            # Check if this external commitment is EU
            entry = next((e for e in person_externals
                          if e["project"] == proj), None)
            if entry and entry.get("is_eu", False):
                eu_row_refs.append(ext_row)

        for mi in range(project_params["months"]):
            c = month_col_start + mi
            cl = col_letter(c)
            if eu_row_refs:
                cell_refs = ",".join(f"{cl}{row}" for row in eu_row_refs)
                formula = f"=SUM({cell_refs})"
            else:
                formula = "=0"
            style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
                       fmt="0.00").value = formula

        # Sum validation column for EU TOTAL
        if eu_row_refs:
            sum_refs = ",".join(f"{col_letter(val_col_sum)}{row}"
                                for row in eu_row_refs)
            style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                       fill=LIGHT_BLUE_FILL, fmt="0.00").value = f"=SUM({sum_refs})"
        else:
            style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                       fill=LIGHT_BLUE_FILL, fmt="0.00").value = "=0"
        # Budget and Variance columns left empty for EU TOTAL
        style_cell(ws, r, val_col_budget, fill=LIGHT_BLUE_FILL, fmt="0.00")
        style_cell(ws, r, val_col_diff, fill=LIGHT_BLUE_FILL, fmt="0.00")
        r += 1
```

Add `"eu_total_row": eu_total_row` to the person_groups dict (~line 452):
```python
person_groups[name] = {
    "header_row": header_row,
    "wp_rows": wp_rows,
    "external_rows": external_rows,
    "total_row": total_row,
    "remaining_row": remaining_row,
    "eu_total_row": eu_total_row,
}
```

**Step 4: Run all tests**

Run: `pytest tests/ -q`
Expected: All pass (existing CF rule count test will need fixing in Task 6)

**Step 5: Commit**

```bash
git add budget_planner/capacity.py tests/test_capacity.py
git commit -m "feat(capacity): add EU TOTAL row per person"
```

---

## Task 6: Add annual EU ceiling CF rules

**Files:**
- Modify: `budget_planner/capacity.py` (after EU TOTAL row, before person_groups dict)
- Modify: `budget_planner/layout.py` (import _calendar_year_groups — already exported)
- Test: `tests/test_conditional_formatting.py`

**Step 1: Write the failing test**

Add to `tests/test_conditional_formatting.py`:
```python
def test_capacity_eu_ceiling_red_rules(self, workbook):
    """Each person's EU TOTAL row has red CF for annual EU total > ceiling."""
    ws = workbook["Capacity"]
    rules = _collect_cf_rules(ws)
    # Red rules referencing SUM(...) > Config!$B$ (annual ceiling)
    # These have FFC7CE fill and SUM formula with Config ref
    eu_red_rules = [
        r for r in rules
        if r[2] is not None and r[2].endswith("FFC7CE")
        and any("SUM(" in f and "Config!$B$" in f for f in r[1])
    ]
    # At least one per person per calendar year
    assert len(eu_red_rules) >= len(TEAM), (
        f"Expected at least {len(TEAM)} EU ceiling red rules, got {len(eu_red_rules)}"
    )

def test_capacity_eu_ceiling_amber_rules(self, workbook):
    """Each person's EU TOTAL row has amber CF for annual EU total > ceiling-1."""
    ws = workbook["Capacity"]
    rules = _collect_cf_rules(ws)
    eu_amber_rules = [
        r for r in rules
        if r[2] is not None and r[2].endswith("FFD966")
        and any("SUM(" in f and "Config!$B$" in f for f in r[1])
    ]
    assert len(eu_amber_rules) >= len(TEAM), (
        f"Expected at least {len(TEAM)} EU ceiling amber rules, got {len(eu_amber_rules)}"
    )
```

**Step 2: Run tests to verify they fail**

Run: `pytest tests/test_conditional_formatting.py::TestCapacityConditionalFormatting::test_capacity_eu_ceiling_red_rules -v`
Expected: FAIL

**Step 3: Implement annual ceiling CF**

In `budget_planner/capacity.py`, add import at top:
```python
from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG, CAPACITY_AMBER_LOAD,
    build_year_month_headers, month_col_end as _month_col_end, val_columns,
    _calendar_year_groups,  # NEW
)
```

After the EU TOTAL row block (before the existing TOTAL/REMAINING CF), add:

```python
        # Conditional formatting for EU TOTAL — per calendar year
        ceiling_row = team_layout["param_rows"]["EU Annual Ceiling (PM)"]
        for group_start, span in _calendar_year_groups(
            project_params["start"], project_params["months"]
        ):
            yr_first_col = month_col_start + group_start
            yr_last_col = yr_first_col + span - 1
            yr_range = (f"{col_letter(yr_first_col)}{eu_total_row}:"
                        f"{col_letter(yr_last_col)}{eu_total_row}")
            # Absolute refs so formula doesn't shift across the range
            sum_expr = (f"SUM(${col_letter(yr_first_col)}${eu_total_row}"
                        f":${col_letter(yr_last_col)}${eu_total_row})")
            # Red if annual EU total > ceiling (added first = higher priority)
            ws.conditional_formatting.add(yr_range, FormulaRule(
                formula=[f"{sum_expr}>{SHEET_CONFIG}!$B${ceiling_row}"],
                fill=CF_RED_FILL,
            ))
            # Amber if annual EU total > ceiling - 1
            ws.conditional_formatting.add(yr_range, FormulaRule(
                formula=[f"{sum_expr}>{SHEET_CONFIG}!$B${ceiling_row}-1"],
                fill=CF_AMBER_FILL,
            ))
```

**Step 4: Update the CF rule count test**

In `tests/test_conditional_formatting.py`, update `test_capacity_rule_count`:

The new rule count adds 2 CF rules (red + amber) per person per calendar year:
```python
def test_capacity_rule_count(self, workbook):
    ws = workbook["Capacity"]
    rules = _collect_cf_rules(ws)
    total_wp_rows = len(DEMO_CAPACITY_ASSIGNMENTS)
    num_years = len(list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS)))
    # yellow per WP + 4 per person (TOTAL/REMAINING) + 2*years per person (EU ceiling)
    expected = total_wp_rows + (4 + 2 * num_years) * len(TEAM)
    assert len(rules) == expected, (
        f"Expected {expected} CF rules on Capacity, got {len(rules)}"
    )
```

**Step 5: Run all tests**

Run: `pytest tests/ -q`
Expected: All pass

**Step 6: Commit**

```bash
git add budget_planner/capacity.py tests/test_conditional_formatting.py
git commit -m "feat(capacity): add per-calendar-year EU ceiling CF rules"
```

---

## Task 7: Update reader to skip EU TOTAL row

**Files:**
- Modify: `budget_planner/reader.py:352` (skip list)
- Test: `tests/test_integration.py` (round-trip)

**Step 1: Write the failing test**

Add to `tests/test_integration.py` (or `tests/test_rebuild.py`):
```python
def test_rebuild_preserves_eu_total_structure(tmp_path):
    """Rebuild preserves person groups despite EU TOTAL row."""
    from budget_planner.workbook import init_workbook, rebuild_workbook
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS, TEAM
    import openpyxl

    out = str(tmp_path / "test.xlsx")
    init_workbook(out, assignments=DEMO_CAPACITY_ASSIGNMENTS)
    rebuild_workbook(out)
    wb = openpyxl.load_workbook(out, data_only=False)
    cap = wb["Capacity"]
    # Verify EU TOTAL rows exist after rebuild
    eu_total_count = sum(
        1 for r in range(1, cap.max_row + 1)
        if cap.cell(r, 2).value == "EU TOTAL"
    )
    assert eu_total_count == len(TEAM), (
        f"Expected {len(TEAM)} EU TOTAL rows after rebuild, got {eu_total_count}"
    )
```

**Step 2: Run test to verify it fails (or passes — reader may already skip it)**

Run: `pytest tests/test_integration.py::test_rebuild_preserves_eu_total_structure -v` (adjust path as needed)

**Step 3: Update reader skip list**

In `budget_planner/reader.py`, update the TOTAL/REMAINING skip check (~line 352):
```python
# Old:
if col_b and str(col_b).upper() in ("TOTAL", "REMAINING"):

# New:
if col_b and str(col_b).upper() in ("TOTAL", "REMAINING", "EU TOTAL"):
```

**Step 4: Run all tests**

Run: `pytest tests/ -q`
Expected: All pass

**Step 5: Commit**

```bash
git add budget_planner/reader.py tests/test_integration.py
git commit -m "fix(reader): skip EU TOTAL row during capacity read"
```

---

## Task 8: Make row constants dynamic in app logic and tests

**Why:** Several places hardcode Config param row numbers (1-5) and the fallback
`cfg_wp_first = 8` in plan.py. Adding or reordering params breaks these silently.
The `param_rows` dict built in `config_sheet.py` already tracks positions — use it
(via `config_layout`) everywhere instead of magic numbers.

**Files:**
- Modify: `budget_planner/config_sheet.py` (named ranges → use `param_rows` for all three)
- Modify: `budget_planner/plan.py` (compact header formulas → use `config_layout["param_rows"]`)
- Modify: `budget_planner/reader.py` (label-scan loop instead of hardcoded row numbers)
- Modify: `budget_planner/plan.py:92-94` (remove hardcoded fallback `cfg_wp_first = 8`)
- Modify: `tests/test_rebuild_e2e.py` (derive `CONFIG_FIRST_WP_ROW` from generated workbook)
- Modify: `tests/test_config_sheet.py` (assert params by label scan, not `ws.cell(N, 2)`)
- Modify: `tests/test_rebuild.py` (replace `wb["Config"].cell(1, 2)` / `cell(3, 2)` with label scan)
- Modify: `tests/test_protection.py` (replace `wp_first_row = 8` with layout lookup)

**Step 1: Expose `param_rows` from config_sheet builder**

`build_config_sheet` already returns a layout dict. Add `"param_rows": param_rows` to it
so callers can look up `param_rows["Project Start"]` → row number.

**Step 2: Use `param_rows` for named ranges in config_sheet.py**

Replace the two hardcoded named ranges:
```python
# Old:
dn_start = DefinedName("project_start", attr_text=f"{SHEET_CONFIG}!$B$2")
dn_rate = DefinedName("fte_rate", attr_text=f"{SHEET_CONFIG}!$B$4")

# New:
start_row = param_rows["Project Start"]
dn_start = DefinedName("project_start", attr_text=f"{SHEET_CONFIG}!$B${start_row}")
rate_row = param_rows["FTE Rate (EUR)"]
dn_rate = DefinedName("fte_rate", attr_text=f"{SHEET_CONFIG}!$B${rate_row}")
```

**Step 3: Use `config_layout["param_rows"]` in plan.py compact header**

Pass `param_rows` through `config_layout` to plan builder. Replace:
```python
# Old:
).value = f"={SHEET_CONFIG}!B1"   # title
).value = f"={SHEET_CONFIG}!B2"   # start
).value = f"={SHEET_CONFIG}!B4"   # rate

# New:
).value = f"={SHEET_CONFIG}!B{config_layout['param_rows']['Project Title']}"
).value = f"={SHEET_CONFIG}!B{config_layout['param_rows']['Project Start']}"
).value = f"={SHEET_CONFIG}!B{config_layout['param_rows']['FTE Rate (EUR)']}"
```

**Step 4: Remove hardcoded fallback in plan.py**

Delete the `else` branch with `cfg_wp_first = 8`:
```python
# Old:
if config_layout is not None:
    cfg_wp_first = config_layout["wp_first_row"]
    cfg_wp_last = config_layout["wp_last_row"]
else:
    cfg_wp_first = 8
    cfg_wp_last = 8 + len(work_packages) - 1

# New — config_layout is always provided:
cfg_wp_first = config_layout["wp_first_row"]
cfg_wp_last = config_layout["wp_last_row"]
```

Verify all callers of `build_plan_sheet` pass `config_layout`. If any don't,
make the parameter required (remove default `None`).

**Step 5: Make reader.py label-scan dynamic**

Replace hardcoded row-number reads with a label scan:
```python
# Old:
project_params = {
    "title": ws.cell(1, 2).value,
    "start": ws.cell(2, 2).value,
    "months": int(ws.cell(3, 2).value or 60),
    "rate": float(ws.cell(4, 2).value or 7500),
    "eu_ceiling_pm": float(ws.cell(5, 2).value or 12.0),
}

# New — scan column A for known labels:
_param_map = {
    "Project Title": ("title", None),
    "Project Start": ("start", None),
    "Project Months": ("months", lambda v: int(v or 60)),
    "FTE Rate": ("rate", lambda v: float(v or 7500)),
    "EU Annual Ceiling": ("eu_ceiling_pm", lambda v: float(v or 12.0)),
}
project_params: dict = {}
for r in range(1, ws.max_row + 1):
    label = str(ws.cell(r, 1).value or "")
    for prefix, (key, converter) in _param_map.items():
        if prefix.lower() in label.lower():
            raw = ws.cell(r, 2).value
            project_params[key] = converter(raw) if converter else raw
            break
    if len(project_params) == len(_param_map):
        break
# Validate all params found
for prefix, (key, _) in _param_map.items():
    if key not in project_params:
        raise ValueError(f"Config sheet missing param row matching '{prefix}'")
```

**Step 6: Make test constants dynamic**

In `tests/test_rebuild_e2e.py`, replace the hardcoded constant:
```python
# Old:
CONFIG_FIRST_WP_ROW = 9

# New — derive from a generated workbook:
def _get_config_first_wp_row():
    from budget_planner.reader import _find_header_row
    from budget_planner.workbook import init_workbook
    import openpyxl, tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        init_workbook(path)
        wb = openpyxl.load_workbook(path)
        return _find_header_row(wb["Config"], "WORK PACKAGES") + 2
    finally:
        os.unlink(path)

CONFIG_FIRST_WP_ROW = _get_config_first_wp_row()
```

Similarly update `tests/test_config_sheet.py` tests to find param rows by label,
`tests/test_rebuild.py` to use `read_config_sheet()` instead of raw cell reads,
and `tests/test_protection.py` to derive `wp_first_row` from layout.

**Step 7: Run all tests**

Run: `pytest tests/ -v`
Expected: All pass

**Step 8: Commit**

```bash
git add budget_planner/config_sheet.py budget_planner/plan.py budget_planner/reader.py \
        tests/test_rebuild_e2e.py tests/test_config_sheet.py tests/test_rebuild.py \
        tests/test_protection.py
git commit -m "refactor: make Config row constants dynamic in app and tests"
```

---

## Summary of changes per file

| File | Change |
|------|--------|
| `config.py` | Add `EU_ANNUAL_CEILING_PM`, `is_eu` on externals, demo EU entry |
| `config_sheet.py` | 5th param row, dynamic section positioning, named range |
| `capacity.py` | EU Grant column in config block, EU TOTAL row, ceiling CF |
| `reader.py` | Read `eu_ceiling_pm`, dynamic WP start, read `is_eu`, skip EU TOTAL |
| `workbook.py` | Add `eu_ceiling_pm` to init project_params |
| `layout.py` | No changes (only import `_calendar_year_groups`) |
| `test_config_sheet.py` | Ceiling param, named range, round-trip tests |
| `test_capacity.py` | EU Grant column, EU TOTAL row, formula tests |
| `test_conditional_formatting.py` | CF rule count update, ceiling red/amber tests |
| `test_rebuild_e2e.py` | Dynamic CONFIG_FIRST_WP_ROW |
| `test_rebuild.py` | Replace hardcoded cell reads with reader |
| `test_protection.py` | Derive wp_first_row from layout |
| `test_integration.py` | Rebuild round-trip with EU TOTAL |
