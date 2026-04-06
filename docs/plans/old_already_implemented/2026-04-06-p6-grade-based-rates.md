# P6: Grade-Based Personnel Rates — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Replace the single flat FTE rate with a per-grade rate table (Senior/Mid/Junior) so fte_rate becomes a weighted-average formula, matching EU auditor expectations for personnel cost justification.

**Architecture:** Add a RATE BY GRADE section to Config sheet between params and WP table. Add Grade column to the team roster. Change fte_rate (Config!B4) from blue input to black formula (weighted average). All downstream formulas referencing `fte_rate` named range remain unchanged — the named range still points to Config!$B$4.

**Tech Stack:** openpyxl, Python 3.12, pytest, pyright

**Design doc:** `docs/plans/2026-04-05-grade-based-rates-design.md`

---

## Task 1: config.py — Grade Seed Data

**Files:**
- Modify: `budget_planner/config.py`

**What:** Add grade rate constants and add "grade" field to TEAM seed data.

**Step 1: Add constants after EU_ANNUAL_CEILING_PM (line 20 area)**

```python
# After DEFAULT_RP_LENGTH line:

# ── GRADE RATES ──
# EUR per person-month by job grade (bookkeeping-provided bands)
GRADE_RATES = [
    {"grade": "Senior", "rate": 10000},
    {"grade": "Mid",    "rate": 7500},
    {"grade": "Junior", "rate": 5500},
]
DEFAULT_GRADE = "Mid"
```

**Step 2: Add "grade" key to each TEAM member**

```python
TEAM = [
    {"name": "Alice", "max_fte": 1.0, "grade": "Mid"},
    {"name": "Bob",   "max_fte": 1.0, "grade": "Mid"},
    {"name": "Carol", "max_fte": 1.0, "grade": "Mid"},
    {"name": "Dave",  "max_fte": 1.0, "grade": "Mid"},
]
```

All default to "Mid" — the user changes grades in Excel post-generation.

**Step 3: Run tests**

Run: `pytest tests/ -v`
Expected: All 245 pass (no consumer of grade data yet).

---

## Task 2: config_sheet.py — Grade Rate Table + Roster Grade Column + fte_rate Formula

**Files:**
- Modify: `budget_planner/config_sheet.py`
- Modify: `tests/test_config_sheet.py`
- Modify: `tests/test_team.py`

This is the core builder change. Three sub-changes in one task because they're tightly coupled (grade table row positions affect roster, which affects fte_rate formula).

### 2a: Add RATE BY GRADE section between params and WP table

Insert a new section after the param rows (row 5) and before the WP section header. This shifts the WP table, roster, and reporting periods down by ~6 rows. All positions are computed dynamically, so only the builder code changes.

In `config_sheet.py`, add a `grade_rates=None` parameter to `build_config_sheet()`. When None, lazy-import `config.GRADE_RATES`.

After the param_defs loop (around line 69), add:

```python
# ── Section A.5: Rate by Grade ──
if grade_rates is None:
    from budget_planner.config import GRADE_RATES
    grade_rates = GRADE_RATES

grade_section_row = len(param_defs) + 2  # gap row after params, then section
ws.merge_cells(start_row=grade_section_row, start_column=1,
               end_row=grade_section_row, end_column=2)
style_cell(ws, grade_section_row, 1, value="RATE BY GRADE",
           font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
style_cell(ws, grade_section_row, 2, fill=SECTION_FILL)

grade_header_row = grade_section_row + 1
style_cell(ws, grade_header_row, 1, "Grade", font=HEADER_FONT,
           fill=LIGHT_BLUE_FILL, alignment=CENTER)
style_cell(ws, grade_header_row, 2, "EUR/PM", font=HEADER_FONT,
           fill=LIGHT_BLUE_FILL, alignment=CENTER)

grade_first_row = grade_header_row + 1
for gi, gr in enumerate(grade_rates):
    gr_row = grade_first_row + gi
    style_cell(ws, gr_row, 1, gr["grade"], font=BLUE_INPUT, border=False)
    style_cell(ws, gr_row, 2, gr["rate"], font=BLUE_INPUT, fmt="€#,##0",
               border=False)
grade_last_row = grade_first_row + len(grade_rates) - 1

# tbl_GradeRates named range
dn_grades = DefinedName(
    "tbl_GradeRates",
    attr_text=f"{SHEET_CONFIG}!$A${grade_first_row}:$B${grade_last_row}",
)
wb.defined_names.add(dn_grades)
```

Update the WP section_row calculation to follow the grade table:

```python
# ── Section B: Work Packages Table ──
section_row = grade_last_row + 2  # gap row after grade table
```

(Previously: `section_row = len(param_defs) + 2`)

### 2b: Add Grade column (col C) to Team Roster

In the roster section, widen the section header merge from A:B to A:C. Add "Grade" column header. Write grade values from team data. Add data validation dropdown referencing the grade table.

```python
# Section header — merged A:C (was A:B)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
style_cell(ws, r, 1, value="TEAM ROSTER",
           font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
style_cell(ws, r, 2, fill=SECTION_FILL)
style_cell(ws, r, 3, fill=SECTION_FILL)
r += 1

# Column headers — add Grade
roster_header_row = r
style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
           alignment=CENTER)
style_cell(ws, r, 2, "Max FTE", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
           alignment=CENTER)
style_cell(ws, r, 3, "Grade", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
           alignment=CENTER)
r += 1

# Team member data rows — add grade column
roster_first_row = r
person_rows: dict[str, int] = {}
for member in team:
    style_cell(ws, r, 1, member["name"], font=BLUE_INPUT)
    style_cell(ws, r, 2, member["max_fte"], font=BLUE_INPUT, fmt="0.00")
    grade = member.get("grade", "Mid")
    style_cell(ws, r, 3, grade, font=BLUE_INPUT)
    person_rows[member["name"]] = r
    r += 1
roster_last_row = r - 1

# Data validation: dropdown from grade table column A
from openpyxl.worksheet.datavalidation import DataValidation
dv = DataValidation(
    type="list",
    formula1=f"=$A${grade_first_row}:$A${grade_last_row}",
    allow_blank=False,
)
dv.error = "Grade must match a value in the RATE BY GRADE table"
dv.errorTitle = "Invalid Grade"
ws.add_data_validation(dv)
for pr in range(roster_first_row, roster_last_row + 1):
    dv.add(ws.cell(pr, 3))
```

### 2c: Change fte_rate from blue input to black formula

After the roster is built (so we know roster row positions), overwrite B4 with the weighted-average formula. The IFERROR handles unmatched grade values gracefully.

```python
# ── Overwrite FTE Rate with weighted-average formula ──
rate_row = param_rows["FTE Rate (EUR)"]
fte_formula = (
    f"=SUMPRODUCT(IFERROR(VLOOKUP(C{roster_first_row}:C{roster_last_row},"
    f"tbl_GradeRates,2,FALSE),0)"
    f"*B{roster_first_row}:B{roster_last_row})"
    f"/SUM(B{roster_first_row}:B{roster_last_row})"
)
cell = style_cell(ws, rate_row, 2, font=BLACK_FORMULA, fmt="€#,##0", border=False)
cell.value = fte_formula
```

Add `grade_first_row` and `grade_last_row` to the return dict:

```python
return {
    # ... existing keys ...
    "grade_first_row": grade_first_row,
    "grade_last_row": grade_last_row,
}
```

### 2d: Update tests

**test_config_sheet.py:**

Update `test_config_sheet_params` — FTE Rate is now a formula, not the numeric value:

```python
def test_config_sheet_params():
    """Project parameters are written correctly."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    assert _find_param_value(ws, "Project Title") == "Test Project"
    assert _find_param_value(ws, "Project Start") == datetime(2026, 7, 1)
    assert _find_param_value(ws, "Project Months") == 60
    # FTE Rate is now a formula (weighted average) — check it's a formula string
    fte_val = _find_param_value(ws, "FTE Rate")
    assert isinstance(fte_val, str) and fte_val.startswith("="), (
        f"Expected formula, got {fte_val!r}"
    )
```

Add new tests:

```python
def test_config_grade_rate_table():
    """RATE BY GRADE section exists with grade rows."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    gf = layout["grade_first_row"]
    gl = layout["grade_last_row"]
    assert gl - gf + 1 == 3  # Senior, Mid, Junior
    assert ws.cell(gf, 1).value == "Senior"
    assert ws.cell(gf + 1, 1).value == "Mid"
    assert ws.cell(gf + 2, 1).value == "Junior"
    # Rates are blue input
    assert "0000FF" in str(ws.cell(gf, 2).font.color.rgb)


def test_config_grade_rates_named_range():
    """tbl_GradeRates named range exists and covers the grade table."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    assert "tbl_GradeRates" in set(wb.defined_names.keys())


def test_config_fte_rate_is_formula():
    """FTE Rate cell is a black formula (weighted average), not blue input."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    rate_row = layout["param_rows"]["FTE Rate (EUR)"]
    cell = ws.cell(rate_row, 2)
    assert isinstance(cell.value, str) and cell.value.startswith("=SUMPRODUCT")
    # Font should be black (formula), not blue (input)
    rgb = str(cell.font.color.rgb) if cell.font.color else ""
    assert "0000FF" not in rgb, f"FTE Rate should be black formula, got {rgb}"
```

**test_team.py:**

Update `test_config_roster_has_all_persons` — no change needed (reads col A only).

Update `test_config_roster_cells_are_blue_input` to also check Grade column (col C):

```python
def test_config_roster_cells_are_blue_input(config_wb):
    """Person name, max_fte, and grade cells have blue font color."""
    _, ws, layout = config_wb
    for r in range(layout["roster_first_row"], layout["roster_last_row"] + 1):
        for col in (1, 2, 3):  # Person, Max FTE, Grade
            font = ws.cell(row=r, column=col).font
            assert font.color and font.color.rgb in ("0000FF", "000000FF"), (
                f"Row {r} col {col} font color: {font.color}"
            )
```

Add grade-specific test:

```python
def test_config_roster_grade_column(config_wb):
    """Grade column (C) has values for each team member."""
    _, ws, layout = config_wb
    for r in range(layout["roster_first_row"], layout["roster_last_row"] + 1):
        grade = ws.cell(row=r, column=3).value
        assert grade in ("Senior", "Mid", "Junior"), (
            f"Row {r}: unexpected grade {grade!r}"
        )
```

**Step: Run tests**

Run: `pytest tests/ -v`
Expected: All pass (some existing tests updated, new tests added).

**Pyright:** Run `pyright` and fix any type errors.

---

## Task 3: reader.py + workbook.py — Round-Trip Grade Data

**Files:**
- Modify: `budget_planner/reader.py`
- Modify: `budget_planner/workbook.py`
- Modify: `tests/test_reader_new.py`
- Modify: `tests/test_rebuild.py`

### 3a: reader.py — Handle formula-based FTE Rate

In `read_config_sheet()`, the FTE Rate converter must handle formula strings:

```python
# Update the _param_map entry for FTE Rate:
"FTE Rate": ("rate", lambda v: float(v) if isinstance(v, (int, float)) else 7500),
```

This returns 7500 (default) when B4 contains a formula string, avoiding `float("=SUMPRODUCT(...)")` crash.

### 3b: reader.py — Add read_grade_rates_from_config()

New function to read the RATE BY GRADE section:

```python
def read_grade_rates_from_config(ws) -> list[dict]:
    """Read grade rate table from Config sheet.

    Returns:
        List of dicts with keys 'grade' (str) and 'rate' (float).
        Empty list if section not found.
    """
    section = _find_header_row(ws, "RATE BY GRADE", required=False)
    if section is None:
        return []

    # Column headers row is section + 1, data starts at section + 2
    r = section + 2
    rates: list[dict] = []
    while r <= ws.max_row:
        grade = ws.cell(r, 1).value
        if grade is None or str(grade).strip() == "":
            break
        rate_val = ws.cell(r, 2).value
        rates.append({
            "grade": str(grade),
            "rate": float(rate_val or 0),
        })
        r += 1
    return rates
```

### 3c: reader.py — Update read_team_from_config() for Grade column

Add Grade column (col C) to the returned dicts:

```python
def read_team_from_config(ws) -> list[dict]:
    # ... existing code to find roster section and Person header ...

    roster: list[dict] = []
    r = person_header + 1
    while r <= ws.max_row:
        name = ws.cell(r, 1).value
        if name is None or str(name).strip() == "":
            break
        max_fte = ws.cell(r, 2).value
        grade = ws.cell(r, 3).value
        entry: dict = {
            "name": str(name),
            "max_fte": float(max_fte or 0),
        }
        if grade is not None and str(grade).strip():
            entry["grade"] = str(grade)
        roster.append(entry)
        r += 1

    return roster
```

Grade is optional in the returned dict — old workbooks without a Grade column still parse correctly.

### 3d: workbook.py — Wire grade_rates through init and rebuild

**init_workbook** — pass grade_rates from config:

In `init_workbook()` (around line 236), add:

```python
grade_rates = config.GRADE_RATES
```

And pass to `_build_all_sheets`:

```python
_build_all_sheets(wb, project_params, work_packages, team,
                  external_commitments, assignments=assignments,
                  grade_rates=grade_rates)
```

**_build_all_sheets** — accept and pass grade_rates:

```python
def _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      *, assignments=None, preserved_actuals=None,
                      rebuild_warnings=None, grade_rates=None):
```

Pass to build_config_sheet:

```python
config_layout = build_config_sheet(config_ws, wb, project_params, work_packages,
                                   team=team, reporting_periods=reporting_periods,
                                   grade_rates=grade_rates)
```

**rebuild_workbook** — read and pass grade_rates:

After reading team (around line 315):

```python
from budget_planner.reader import read_grade_rates_from_config
grade_rates = read_grade_rates_from_config(existing_wb[SHEET_CONFIG])
```

Pass to `_build_all_sheets`:

```python
_build_all_sheets(wb, project_params, work_packages, team,
                  external_commitments,
                  assignments=assignments,
                  preserved_actuals=preserved_actuals,
                  rebuild_warnings=warnings,
                  grade_rates=grade_rates or None)
```

If `grade_rates` is empty (old workbook without grades), passing None lets the builder fall back to config.GRADE_RATES defaults.

### 3e: Tests

**test_reader_new.py** — Add grade round-trip tests:

```python
def test_read_grade_rates_from_config():
    """Grade rate table is read correctly from Config sheet."""
    from budget_planner.reader import read_grade_rates_from_config
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    rates = read_grade_rates_from_config(ws)
    assert len(rates) == 3
    assert rates[0]["grade"] == "Senior"
    assert rates[0]["rate"] == 10000
    assert rates[1]["grade"] == "Mid"
    assert rates[2]["grade"] == "Junior"


def test_read_team_includes_grade():
    """Team roster reader includes grade when present."""
    from budget_planner.reader import read_team_from_config
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    team = read_team_from_config(ws)
    assert all("grade" in m for m in team), "Grade field missing from team members"
    assert all(m["grade"] in ("Senior", "Mid", "Junior") for m in team)


def test_read_config_formula_rate_fallback():
    """FTE Rate reader handles formula value gracefully."""
    from budget_planner.reader import read_config_sheet
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    params, _ = read_config_sheet(ws)
    # Rate should be the default fallback (7500) since B4 is now a formula
    assert params["rate"] == 7500
```

**test_rebuild.py** — Add grade preservation test:

```python
def test_rebuild_preserves_grade_rates(self, init_wb):
    """Grade rate table and person grades survive rebuild."""
    from budget_planner.reader import read_grade_rates_from_config, read_team_from_config
    wb = load_workbook(init_wb)
    original_rates = read_grade_rates_from_config(wb["Config"])
    original_team = read_team_from_config(wb["Config"])
    wb.close()

    rebuild_workbook(input_path=init_wb)

    wb = load_workbook(init_wb)
    rebuilt_rates = read_grade_rates_from_config(wb["Config"])
    rebuilt_team = read_team_from_config(wb["Config"])
    assert len(rebuilt_rates) == len(original_rates)
    for got, expected in zip(rebuilt_rates, original_rates):
        assert got["grade"] == expected["grade"]
        assert got["rate"] == expected["rate"]
    for got, expected in zip(rebuilt_team, original_team):
        assert got.get("grade") == expected.get("grade")
    wb.close()
```

**Step: Run tests**

Run: `pytest tests/ -v`
Expected: All pass.

**Pyright:** Run `pyright` and fix any type errors.

---

## Task 4: meta.py — Document Grade Rate Section

**Files:**
- Modify: `budget_planner/meta.py`

Update the CONFIG SHEET STRUCTURE section in the meta text to document:

1. New SECTION A.5: RATE BY GRADE (between params and WP table)
   - Section header, column headers (Grade, EUR/PM), 3 data rows
   - Named range tbl_GradeRates
   - All blue input

2. Updated SECTION C: TEAM ROSTER
   - Now 3 columns: Person | Max FTE | Grade
   - Grade column has data validation dropdown referencing grade table
   - Blue input

3. Updated Row 4 description
   - FTE Rate (EUR) is now a BLACK FORMULA (weighted average from grade table × roster FTE)
   - Not blue input anymore

4. Updated NAMED RANGES section
   - Add `tbl_GradeRates` = Config!$A${grade_first}:$B${grade_last}

5. Updated MODIFICATION GUIDE
   - TO CHANGE FTE RATE: now says "Edit grade rates in RATE BY GRADE table or change person grades in roster"

**Important:** Verify each text update against the actual builder output. Check row numbers match the generated workbook.

**Step: Run tests**

Run: `pytest tests/ -v`
Expected: All pass.

**Pyright:** Run `pyright` and fix any type errors.

---

## Notes for Implementer

**SUMPRODUCT formula safety:** The dev-workflow rule warns about SUMPRODUCT with mixed text/number arrays. In our formula, VLOOKUP(text_range, ...) returns a numeric array, so the multiplication is number×number — safe. The IFERROR wrapper handles #N/A from unmatched grades.

**Row position cascade:** Adding the grade table section shifts WP table, roster, and RP sections down by ~6 rows. All positions are computed dynamically in the builder, so no hardcoded row fixups needed. BUT: grep test files for any hardcoded row numbers that assume the old layout (dev-workflow rule).

**Protection:** Cell protection is currently disabled (CLAUDE.md rule 7). The fte_rate cell changes from unlocked (blue input) to locked (black formula). This is correct behavior — no manual override needed.

**Backward compatibility on rebuild:** Old workbooks without RATE BY GRADE section: `read_grade_rates_from_config()` returns empty list → builder falls back to `config.GRADE_RATES` defaults. Old workbooks without Grade column: `read_team_from_config()` returns dicts without "grade" key → builder uses `member.get("grade", "Mid")` default.
