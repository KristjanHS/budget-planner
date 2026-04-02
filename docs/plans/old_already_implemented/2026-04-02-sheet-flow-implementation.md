# Sheet Flow Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Extend workbook from 5 to 6 sheets — add Actuals (EUR input from accountant), remove Plan's Actual FTE block, add Personnel Ceiling tracking, extend Summary with budget overview.

**Architecture:** Actuals come as EUR from the accountant, not as FTE per person — the Plan Actual FTE block is dead weight causing source-of-truth confusion. Replace it with a dedicated Actuals sheet (EUR per WP per month). Plan gets a Personnel Ceiling column and global ceiling row. Summary gains Actual EUR columns (from Actuals sheet) and a budget overview section for ceiling compliance. `--update` mode is removed — users transfer legacy data manually.

**Tech Stack:** Python 3.12+, openpyxl, pytest, pyright

**Design doc:** `docs/plans/2026-04-01-sheet-flow-design.md`

---

## Dependency Graph

```
Batch 1 (parallel): Task 1 ║ Task 2 ║ Task 3
Batch 2 (parallel): Task 4 ║ Task 6
Batch 3:            Task 5
Batch 4:            Task 7
Batch 5:            Task 8
```

```
Task 1 (plan.py) ──────────┐
Task 2 (actuals.py — new) ─┼─→ Task 4 (summary.py) ──→ Task 5 (workbook.py) ──→ Task 7 (tests)
Task 3 (reader+CLI) ───────┘    Task 6 (meta.py) ─────┘                         Task 8 (rules)
```

**File ownership per task (no overlaps within a batch):**

| Task | Files modified |
|------|---------------|
| 1 | `budget_planner/plan.py` |
| 2 | `budget_planner/actuals.py` (create) |
| 3 | `budget_planner/reader.py`, `budget_planner/__main__.py` |
| 4 | `budget_planner/summary.py` |
| 5 | `budget_planner/workbook.py` |
| 6 | `budget_planner/meta.py` |
| 7 | `tests/test_workbook.py`, `tests/test_plan_v2.py`, `tests/test_actuals.py` (create), `tests/test_integration.py`, `tests/test_reader_v2.py`, `tests/test_cli.py`, `tests/test_e2e.py` |
| 8 | `.claude/rules/architecture.md`, `.claude/rules/common-tasks.md`, `.claude/rules/dev-workflow.md` |

---

## Task 1: Plan Sheet Redesign

**Files:** Modify `budget_planner/plan.py`

### Changes

#### 1a. Compact global params to 4 rows (A–B only)

Replace the current 2-row globals (lines 69–89) with 4 rows in columns A–B only:

```python
r = 1

# ── GLOBAL PARAMETERS (rows 1–4, columns A–B only) ──
style_cell(ws, r, 1, "Project Title:", font=HEADER_FONT, border=False)
style_cell(ws, r, 2, PROJECT_TITLE,
           font=Font(name="Arial", bold=True, size=14, color="0000FF"),
           border=False)
r += 1
style_cell(ws, r, 1, "Project Start:", font=HEADER_FONT, border=False)
style_cell(ws, r, 2, PROJECT_START, font=BLUE_INPUT, border=False,
           fmt="MMM YYYY")
r += 1
style_cell(ws, r, 1, "FTE Rate (\u20ac):", font=HEADER_FONT, border=False)
style_cell(ws, r, 2, DEFAULT_RATE, font=BLUE_INPUT, border=False,
           fmt="\u20ac#,##0")
r += 1
# Personnel Ceiling — formula filled after config table rows are known
ceiling_row = r
style_cell(ws, r, 1, "Personnel Ceiling:", font=HEADER_FONT, border=False)
r += 1
```

Key differences from current code:
- Remove merged cells on title row (was `merge_cells(B1:G1)`)
- FTE Rate moves from D2:E2 to A3:B3
- New row 4 for Personnel Ceiling (formula added after config table)

#### 1b. Update named ranges

```python
wb.defined_names.add(DefinedName("project_start", attr_text="Plan!$B$2"))
wb.defined_names.add(DefinedName("fte_rate", attr_text="Plan!$B$3"))
# partner_ceiling added after config table is built
```

#### 1c. Config table — rename column G, use named range

Extend merge and headers from 7 to 7 columns (unchanged count — just rename):

```python
config_headers = [
    "WP ID", "WP Name", "Total FTE (PM)", "Start Month",
    "End Month", "Active Months", "Personnel Ceiling (\u20ac)",
]
```

Change Budget EUR formula to use `fte_rate` named range:

```python
# Column G: Personnel Ceiling EUR (was "Budget EUR")
style_cell(ws, row, 7, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
    f"=C{row}*fte_rate"
)
```

Config total row — column G SUM unchanged. After config total, fill in the ceiling at row 4:

```python
# Fill Personnel Ceiling global param
style_cell(ws, ceiling_row, 2, font=BLACK_FORMULA, border=False,
           fmt="\u20ac#,##0").value = (
    f"=SUM(G{config_first_wp_row}:G{config_last_wp_row})"
)
wb.defined_names.add(DefinedName("partner_ceiling", attr_text="Plan!$B$4"))
```

#### 1d. Remove `preserved_data` from `build_plan_sheet`

New signature: `build_plan_sheet(ws, wb)` — no `preserved_data` parameter.

#### 1e. Remove Actual FTE block

Delete these lines (currently ~152–158):
```python
# ── ACTUAL FTE BLOCK ──   <-- DELETE THIS ENTIRE BLOCK
actuals = _build_month_block(
    ws, planned["next_row"], "ACTUAL FTE (PERSON-MONTHS)",
    config_first_wp_row=config_first_wp_row,
    is_actuals=True,
    preserved_data=preserved_data,
)
```

#### 1f. Update layout dict return

Remove all `actual_*` keys:

```python
return {
    "config_header_row": config_header_row,
    "config_col_headers_row": config_col_headers_row,
    "config_first_wp_row": config_first_wp_row,
    "config_last_wp_row": config_last_wp_row,
    "config_total_row": config_total_row,
    "planned_header_row": planned["header_row"],
    "planned_first_wp_row": planned["first_wp_row"],
    "planned_last_wp_row": planned["last_wp_row"],
    "planned_total_row": planned["total_row"],
}
```

#### 1g. Simplify `_build_month_block`

New signature:
```python
def _build_month_block(ws, start_row, block_title, config_first_wp_row):
```

Remove all `is_actuals` and `preserved_data` logic:
- Delete `block_key`, `block_preserved` variables
- Validation diff label: always `"Variance"` (remove `"Remaining" if is_actuals else "Variance"`)
- Seed value: always `even_fte` (remove `even_fte if not is_actuals else 0`)
- Remove `wp_preserved` / preserved value check
- Conditional formatting: always yellow for `ABS(...)>0.01` (remove red-for-actuals branch)

#### 1h. Update `_apply_planned_formulas` — no changes needed

This function only references `layout["planned_first_wp_row"]` and `capacity_layout`, both unchanged.

### Row position changes (for reference)

| Element | Old row | New row | Shift |
|---------|---------|---------|-------|
| Project Title | 1 | 1 | — |
| Project Start | 2 | 2 | — |
| FTE Rate | 2 (col E) | 3 (col B) | moved |
| Personnel Ceiling | — | 4 | new |
| Config header | 4 | 6 | +2 |
| Config col headers | 5 | 7 | +2 |
| Config first WP | 6 | 8 | +2 |
| Config last WP | 13 | 15 | +2 |
| Config total | 14 | 16 | +2 |
| Planned header | 16 | 18 | +2 |
| Planned first WP | 19 | 21 | +2 |

### Verification

```bash
python -c "from budget_planner.plan import build_plan_sheet; print('import ok')"
```

---

## Task 2: Actuals Sheet Builder

**Files:** Create `budget_planner/actuals.py`

### Complete implementation

```python
"""Actuals sheet builder — EUR actuals per WP per month from accountant."""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    RED_FILL, CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_START, PROJECT_MONTHS, WORK_PACKAGES, NUM_WPS,
    MONTH_COL_START, MONTH_COL_END,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
)


def build_actuals_sheet(wb, layout):
    """Build the Actuals sheet for EUR actuals per WP per month.

    Args:
        wb: The parent Workbook.
        layout: Layout dict from build_plan_sheet(), needs config_first_wp_row.

    Returns:
        Layout dict with keys: header_row, year_row, month_row,
        first_wp_row, last_wp_row, total_row.
    """
    config_first_wp_row = layout["config_first_wp_row"]

    ws = wb.create_sheet("Actuals")
    ws.sheet_properties.tabColor = "0000FF"

    r = 1

    # ── Section header ──
    merge_end = VAL_COL_DIFF
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_end)
    style_cell(ws, r, 1, "ACTUAL EUR", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, merge_end + 1):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1
    year_row = r

    # ── Year header row ──
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "", fill=LIGHT_BLUE_FILL)
    m = 0
    while m < PROJECT_MONTHS:
        year_start_col = MONTH_COL_START + m
        months_left = PROJECT_MONTHS - m
        span = min(12, months_left)
        year_end_col = year_start_col + span - 1
        cal_date = PROJECT_START + relativedelta(months=m)
        year_label = f"Year {m // 12 + 1} ({cal_date.year})"
        if span > 1:
            ws.merge_cells(start_row=r, start_column=year_start_col,
                           end_row=r, end_column=year_end_col)
        style_cell(ws, r, year_start_col, year_label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(year_start_col + 1, year_end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)
        m += span
    # Validation headers
    style_cell(ws, r, VAL_COL_SUM, "Sum", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_BUDGET, "Ceiling", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_DIFF, "Variance", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1
    month_row = r

    # ── Month header row ──
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    for m_idx in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m_idx
        cal_date = PROJECT_START + relativedelta(months=m_idx)
        mlabel = f"M{m_idx + 1}\n{cal_date.strftime('%b %y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8),
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_SUM, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_BUDGET, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_DIFF, "", fill=LIGHT_BLUE_FILL)
    r += 1
    first_wp_row = r

    # ── WP rows (blue input EUR cells) ──
    for i, wp in enumerate(WORK_PACKAGES):
        row = r + i
        config_row = config_first_wp_row + i
        # Cross-sheet refs to Plan config
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"=Plan!A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"=Plan!B{config_row}"

        for m_idx in range(PROJECT_MONTHS):
            c = MONTH_COL_START + m_idx
            month_num = m_idx + 1
            if wp["start"] <= month_num <= wp["end"]:
                style_cell(ws, row, c, 0, font=BLUE_INPUT, fmt="\u20ac#,##0")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="\u20ac#,##0")

        # Validation: Sum
        sum_range = (f"{col_letter(MONTH_COL_START)}{row}:"
                     f"{col_letter(MONTH_COL_END)}{row}")
        style_cell(ws, row, VAL_COL_SUM, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
        # Validation: Ceiling (from Plan Personnel Ceiling column G)
        style_cell(ws, row, VAL_COL_BUDGET, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = f"=Plan!G{config_row}"
        # Validation: Variance = Ceiling − Sum
        style_cell(ws, row, VAL_COL_DIFF, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(VAL_COL_BUDGET)}{row}"
            f"-{col_letter(VAL_COL_SUM)}{row}"
        )

    last_wp_row = first_wp_row + NUM_WPS - 1
    r = last_wp_row + 1

    # ── TOTAL row ──
    total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for m_idx in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m_idx
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="\u20ac#,##0").value = f"=SUM({col_range})"
    # TOTAL Sum
    sum_range = (f"{col_letter(VAL_COL_SUM)}{first_wp_row}:"
                 f"{col_letter(VAL_COL_SUM)}{last_wp_row}")
    style_cell(ws, r, VAL_COL_SUM, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
    # TOTAL Ceiling = partner_ceiling named range
    style_cell(ws, r, VAL_COL_BUDGET, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = "=partner_ceiling"
    # TOTAL Variance = Ceiling − Sum
    style_cell(ws, r, VAL_COL_DIFF, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"={col_letter(VAL_COL_BUDGET)}{r}"
        f"-{col_letter(VAL_COL_SUM)}{r}"
    )

    # ── Conditional formatting: red on negative variance ──
    diff_col_l = col_letter(VAL_COL_DIFF)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=RED_FILL,
    ))

    # Freeze panes
    ws.freeze_panes = f"C{first_wp_row}"

    return {
        "header_row": 1,
        "year_row": year_row,
        "month_row": month_row,
        "first_wp_row": first_wp_row,
        "last_wp_row": last_wp_row,
        "total_row": total_row,
    }
```

### Verification

```bash
python -c "from budget_planner.actuals import build_actuals_sheet; print('import ok')"
```

---

## Task 3: Reader + CLI Cleanup

**Files:** Modify `budget_planner/reader.py`, `budget_planner/__main__.py`

### 3a. reader.py — Update `discover_layout` (lines 299–366)

Make ACTUAL block optional. Remove the requirement to find `"ACTUAL"` header:

```python
def discover_layout(ws) -> dict:
    """Scan existing Plan worksheet, reconstruct layout dict.

    Discovers config table and PLANNED block. ACTUAL block is not expected
    (removed in sheet flow redesign).
    """
    # Find config table — same as before (scan for "WP ID" in col A)
    config_header_row = None
    for r in range(1, 21):
        val = ws.cell(r, 1).value
        if val == "WP ID":
            config_header_row = r
            break
    if config_header_row is None:
        raise ValueError("Cannot find config table header row")

    # Count WP rows — same as before
    config_first_wp_row = config_header_row + 1
    config_last_wp_row = config_first_wp_row
    while True:
        next_val = ws.cell(config_last_wp_row + 1, 1).value
        next_name = str(ws.cell(config_last_wp_row + 1, 2).value or "")
        if not next_val or next_name.upper().startswith("TOTAL"):
            break
        config_last_wp_row += 1
    config_total_row = config_last_wp_row + 1

    # Find PLANNED block header
    planned_header_row = None
    for r in range(config_total_row + 1, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or "")
        if "PLANNED" in val.upper():
            planned_header_row = r
            break
    if planned_header_row is None:
        raise ValueError("Cannot find PLANNED FTE block header")

    num_wps = config_last_wp_row - config_first_wp_row + 1
    planned_first_wp_row = planned_header_row + 3
    planned_last_wp_row = planned_first_wp_row + num_wps - 1
    planned_total_row = planned_last_wp_row + 1

    return {
        "config_header_row": config_header_row,
        "config_first_wp_row": config_first_wp_row,
        "config_last_wp_row": config_last_wp_row,
        "config_total_row": config_total_row,
        "planned_header_row": planned_header_row,
        "planned_first_wp_row": planned_first_wp_row,
        "planned_last_wp_row": planned_last_wp_row,
        "planned_total_row": planned_total_row,
    }
```

### 3b. reader.py — Remove `read_workbook` function

Delete the `read_workbook` function (lines 369–420) and `_read_monthly_values` (lines 281–296). Keep `discover_layout`, `read_plan_config`, `read_team_sheet`, `read_capacity_sheet` — they're used by tests and may be useful later.

### 3c. __main__.py — Remove `--update` flag

Remove these lines from `main()`:
```python
parser.add_argument("--update", metavar="PATH", ...)
```

And the handling block:
```python
if args.update:
    ...
```

The `generate_workbook` call simplifies (no `existing_path`):
```python
path = generate_workbook(output_path, mode=mode)
```

### Verification

```bash
python -m budget_planner --help  # No --update shown
python -c "from budget_planner.reader import discover_layout; print('ok')"
```

---

## Task 4: Summary Sheet Redesign

**Files:** Modify `budget_planner/summary.py`

### Changes

#### 4a. New function signature

```python
def build_summary_sheet(wb, layout, actuals_layout):
    """Create the Summary sheet with annual budget view and budget overview.

    Parameters
    ----------
    wb : openpyxl.Workbook
    layout : dict
        From build_plan_sheet(): config_first_wp_row, planned_first_wp_row.
    actuals_layout : dict
        From build_actuals_sheet(): first_wp_row, total_row.
    """
    config_first_wp_row = layout["config_first_wp_row"]
    planned_first_wp_row = layout["planned_first_wp_row"]
    actuals_first_wp_row = actuals_layout["first_wp_row"]
    actuals_total_row = actuals_layout["total_row"]
```

#### 4b. Year columns: 4 per year (was 5)

```python
COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Actual EUR, Variance EUR
```

Sub-headers:
```python
year_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
```

Per-WP year formulas:
```python
for y in range(num_years):
    m_start = y * 12
    m_end = min((y + 1) * 12, PROJECT_MONTHS) - 1
    first_col_l = col_letter(MONTH_COL_START + m_start)
    last_col_l = col_letter(MONTH_COL_START + m_end)

    sc = YEAR_GROUP_START + y * COLS_PER_YEAR
    plan_wp_row = planned_first_wp_row + i
    actuals_wp_row = actuals_first_wp_row + i

    # Plan FTE (cross-sheet ref to Plan Planned block)
    style_cell(summary, row, sc, font=GREEN_XREF, fmt="0.00").value = (
        f"=SUM(Plan!{first_col_l}{plan_wp_row}:{last_col_l}{plan_wp_row})"
    )
    # Plan EUR = Plan FTE × fte_rate
    style_cell(summary, row, sc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
        f"={col_letter(sc)}{row}*fte_rate"
    )
    # Actual EUR (cross-sheet ref to Actuals sheet)
    style_cell(summary, row, sc + 2, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
        f"=SUM(Actuals!{first_col_l}{actuals_wp_row}:{last_col_l}{actuals_wp_row})"
    )
    # Variance EUR = Plan EUR − Actual EUR
    style_cell(summary, row, sc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
        f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
    )
```

#### 4c. Project total columns: 4 (was 7)

```python
TOTAL_COLS = 4  # Plan FTE, Plan EUR, Actual EUR, Variance EUR

total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
```

Per-WP project totals: SUM of year columns for each metric.

Totals row: SUM of WP rows per column, except `% Used` (removed — moved to budget overview).

#### 4d. Conditional formatting on Variance EUR

Red fill if negative (actuals exceed plan) — both year columns and project totals:

```python
for y in range(num_years):
    var_col = YEAR_GROUP_START + y * COLS_PER_YEAR + 3
    var_col_l = col_letter(var_col)
    var_range = f"{var_col_l}{SUMMARY_FIRST_WP_ROW}:{var_col_l}{SUMMARY_LAST_WP_ROW}"
    summary.conditional_formatting.add(var_range, FormulaRule(
        formula=[f"{var_col_l}{SUMMARY_FIRST_WP_ROW}<0"], fill=RED_FILL,
    ))
```

#### 4e. Budget overview section (NEW — below totals row)

```python
# ── BUDGET OVERVIEW ──
r = totals_row + 2
summary.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
style_cell(summary, r, 1, "BUDGET OVERVIEW", font=SECTION_FONT,
           fill=SECTION_FILL, alignment=CENTER)
style_cell(summary, r, 2, fill=SECTION_FILL)
r += 2

# Personnel Actual — total EUR from Actuals sheet
actual_row = r
style_cell(summary, r, 1, "Personnel Actual", font=HEADER_FONT, border=False)
style_cell(summary, r, 2, font=GREEN_XREF, fmt="\u20ac#,##0",
           border=False).value = (
    f"=Actuals!{col_letter(VAL_COL_SUM)}{actuals_total_row}"
)
r += 1

# Total Claimable (+25%)
style_cell(summary, r, 1, "Total Claimable (+25%)", font=HEADER_FONT, border=False)
style_cell(summary, r, 2, font=BLACK_FORMULA, fmt="\u20ac#,##0",
           border=False).value = f"=B{actual_row}*1.25"
r += 1

# Personnel Ceiling
ceiling_row = r
style_cell(summary, r, 1, "Personnel Ceiling", font=HEADER_FONT, border=False)
style_cell(summary, r, 2, font=GREEN_XREF, fmt="\u20ac#,##0",
           border=False).value = "=partner_ceiling"
r += 1

# Remaining = Ceiling − Personnel Actual
remaining_row = r
style_cell(summary, r, 1, "Remaining", font=HEADER_FONT, border=False)
style_cell(summary, r, 2, font=BLACK_FORMULA, fmt="\u20ac#,##0",
           border=False).value = f"=B{ceiling_row}-B{actual_row}"
r += 1

# % Consumed
style_cell(summary, r, 1, "% Consumed", font=HEADER_FONT, border=False)
style_cell(summary, r, 2, font=BLACK_FORMULA, fmt="0.0%",
           border=False).value = (
    f"=IF(B{ceiling_row}=0,0,B{actual_row}/B{ceiling_row})"
)
```

#### 4f. Conditional formatting on Remaining

```python
from budget_planner.styles import AMBER_FILL  # add to imports

# Red if negative
summary.conditional_formatting.add(
    f"B{remaining_row}", FormulaRule(
        formula=[f"B{remaining_row}<0"], fill=RED_FILL,
    ))
# Amber if < 10% of ceiling (add lower-priority first per openpyxl convention)
summary.conditional_formatting.add(
    f"B{remaining_row}", FormulaRule(
        formula=[f"B{remaining_row}<B{ceiling_row}*0.1"], fill=AMBER_FILL,
    ))
```

#### 4g. Import changes

Add `AMBER_FILL` and `VAL_COL_SUM, VAL_COL_DIFF` to imports from styles/config:

```python
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT, TITLE_FONT,
    SECTION_FONT, SECTION_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    AMBER_FILL, RED_FILL, CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_START, PROJECT_MONTHS, WORK_PACKAGES, NUM_WPS,
    MONTH_COL_START, VAL_COL_SUM,
)
```

### Verification

```bash
python -c "from budget_planner.summary import build_summary_sheet; print('ok')"
```

---

## Task 5: Workbook Orchestrator

**Files:** Modify `budget_planner/workbook.py`

### Changes

#### 5a. Update imports

```python
from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.team import build_team_sheet
from budget_planner.capacity import build_capacity_sheet
from budget_planner.actuals import build_actuals_sheet  # NEW
from budget_planner.summary import build_summary_sheet
from budget_planner.meta import build_meta_sheet
# REMOVE: from budget_planner.reader import read_workbook
```

#### 5b. Simplify `generate_workbook` signature

```python
def generate_workbook(output_path: str = "output/budget_plan.xlsx",
                      mode: str = "demo") -> str:
    """Generate workbook."""
```

Remove `existing_path` parameter entirely.

#### 5c. Remove all --update/reader logic

Delete:
- `preserved = read_workbook(existing_path) if existing_path else {}`
- `team_data` preparation block (lines 65–75)
- All `preserved` data passing

Simplify builder calls:
```python
# Plan sheet
plan_ws = wb.create_sheet("Plan")
layout = build_plan_sheet(plan_ws, wb)

# Team sheet
team_ws = wb.create_sheet("Team")
team_layout = build_team_sheet(team_ws, wb)

# Capacity sheet
capacity_ws = wb.create_sheet("Capacity")
capacity_layout = build_capacity_sheet(capacity_ws, wb, team_layout, mode=mode)

# Overwrite planned FTE with cross-sheet formulas from Capacity
_apply_planned_formulas(plan_ws, layout, capacity_layout)

# Actuals sheet (NEW)
actuals_layout = build_actuals_sheet(wb, layout)

# Summary sheet (updated signature)
build_summary_sheet(wb, layout, actuals_layout)

# _META sheet
build_meta_sheet(meta_ws, layout)
```

#### 5d. Sheet order

Sheets created in order: _META, Plan, Team, Capacity, Actuals, Summary.
The `wb.move_sheet("_META", offset=0)` line keeps _META first. Verify Actuals lands between Capacity and Summary.

### Verification

```bash
python -m budget_planner && echo "Generation succeeded"
python -c "
import openpyxl
wb = openpyxl.load_workbook('output/budget_plan.xlsx')
print(wb.sheetnames)
assert wb.sheetnames == ['_META', 'Plan', 'Team', 'Capacity', 'Actuals', 'Summary']
print('Sheet order OK')
"
```

---

## Task 6: Meta Sheet Update

**Files:** Modify `budget_planner/meta.py`

### Changes

Rewrite the large text block in `build_meta_sheet` to document the v3 workbook structure. Key changes to the text:

1. **Sheet index:** 6 sheets (add Actuals between Capacity and Summary, tab color Blue)
2. **Plan section:** 4-row globals (A–B only), Personnel Ceiling at B4, column G renamed to "Personnel Ceiling (€)", no Actual FTE block
3. **Actuals section (NEW):** EUR per WP per month, TOTAL row, Ceiling/Variance validation, blue input cells, inactive months greyed
4. **Summary section:** 4 columns per year (Plan FTE, Plan EUR, Actual EUR, Variance EUR), budget overview section below annual table
5. **Named ranges:** `project_start` → B2, `fte_rate` → B3, `partner_ceiling` → B4
6. **Data flow:** `Team → Capacity → Plan ← Actuals → Summary` (no circular refs)
7. **Build order:** Plan → Team → Capacity → patch Plan → Actuals → Summary
8. **Remove mentions of:** `--update`, round-trip, actual FTE block, `read_workbook`
9. **Modification guide:** Update for new structure (how to add WPs now affects Actuals sheet too)

### Verification

```bash
python -c "from budget_planner.meta import build_meta_sheet; print('ok')"
```

---

## Task 7: Test Suite Update

**Files:** Modify 6 test files, create 1 new test file

### 7a. `tests/test_workbook.py`

**Changes to existing tests:**

| Test | Change |
|------|--------|
| `test_sheets_exist` | Expect 6 sheets: `["_META", "Plan", "Team", "Capacity", "Actuals", "Summary"]` |
| `test_global_params` | Check `ws["B2"]` (start) and `ws["B3"]` (rate, was `ws["E2"]`); remove `ws["E2"]` check |
| `test_named_ranges` | `fte_rate` → `Plan!$B$3` (was `$E$2`); ADD `partner_ceiling` → `Plan!$B$4` |
| `test_wp_config_valid` | Use dynamic `layout["config_first_wp_row"]` — should auto-adapt |
| `test_actual_fte_zeros` | **DELETE** — no Actual FTE block |
| `test_inactive_months_empty` | **Remove actual block check** — keep only planned block check. Update to use `layout["planned_first_wp_row"]` and `layout["planned_last_wp_row"]` only |
| `test_summary_wp_count` | May need adjustment if Summary row structure changed |
| `test_meta_has_modification_guide` | No change expected |

**Remove imports** of `actual_*` related items if any.

### 7b. `tests/test_plan_v2.py`

**Minimal changes expected.** Both tests use `discover_layout()` which is updated (Task 3) to return layout without `actual_*` keys. The tests only use `planned_first_wp_row` — should auto-adapt to new row positions.

Verify both `test_planned_fte_formula_refs_match_capacity_layout` and `test_planned_fte_cells_have_black_font` still pass with the new row offsets.

### 7c. `tests/test_actuals.py` (CREATE)

```python
"""Tests for the Actuals sheet builder."""
import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.reader import discover_layout
from budget_planner.config import (
    WORK_PACKAGES, NUM_WPS, MONTH_COL_START, PROJECT_MONTHS,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
)
from budget_planner.styles import col_letter


@pytest.fixture(scope="module")
def actuals_workbook(tmp_path_factory):
    out = tmp_path_factory.mktemp("actuals") / "test.xlsx"
    generate_workbook(str(out))
    return openpyxl.load_workbook(str(out), data_only=False)


def test_actuals_sheet_exists(actuals_workbook):
    assert "Actuals" in actuals_workbook.sheetnames


def test_actuals_tab_color(actuals_workbook):
    ws = actuals_workbook["Actuals"]
    assert ws.sheet_properties.tabColor and "0000FF" in str(
        ws.sheet_properties.tabColor.rgb or ws.sheet_properties.tabColor.theme
        if ws.sheet_properties.tabColor else ""
    ).upper() or ws.sheet_properties.tabColor.rgb == "000000FF"


def test_actuals_wp_count(actuals_workbook):
    """Actuals has one row per WP between header and TOTAL."""
    ws = actuals_workbook["Actuals"]
    # Find TOTAL row
    total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None
    # Find first WP row (row after month headers — look for "WP ID" in col A)
    month_header_row = None
    for r in range(1, total_row):
        if ws.cell(r, 1).value == "WP ID":
            month_header_row = r
            break
    assert month_header_row is not None
    first_wp_row = month_header_row + 1
    wp_count = total_row - first_wp_row
    assert wp_count == NUM_WPS


def test_actuals_values_are_zero(actuals_workbook):
    """Fresh workbook has 0 in active month cells."""
    ws = actuals_workbook["Actuals"]
    # Find first WP row
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    wp = WORK_PACKAGES[0]
    row = first_wp_row
    c = MONTH_COL_START + wp["start"] - 1
    assert ws.cell(row, c).value == 0


def test_actuals_inactive_months_empty(actuals_workbook):
    """Inactive month cells are empty with grey fill."""
    ws = actuals_workbook["Actuals"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    # WP3 starts at month 7 — months 1-6 should be inactive
    wp3_idx = next(i for i, w in enumerate(WORK_PACKAGES) if w["id"] == "WP3")
    row = first_wp_row + wp3_idx
    for m in range(6):  # months 1-6
        c = MONTH_COL_START + m
        val = ws.cell(row, c).value
        assert val is None or val == "", f"M{m+1} should be inactive, got {val}"


def test_actuals_sum_formula(actuals_workbook):
    """Sum validation column has SUM formula."""
    ws = actuals_workbook["Actuals"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    cell = ws.cell(first_wp_row, VAL_COL_SUM)
    assert cell.value and str(cell.value).startswith("=SUM(")


def test_actuals_ceiling_refs_plan(actuals_workbook):
    """Ceiling validation column references Plan config table."""
    ws = actuals_workbook["Actuals"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    cell = ws.cell(first_wp_row, VAL_COL_BUDGET)
    assert cell.value and "Plan!G" in str(cell.value)


def test_actuals_variance_formula(actuals_workbook):
    """Variance = Ceiling − Sum."""
    ws = actuals_workbook["Actuals"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    cell = ws.cell(first_wp_row, VAL_COL_DIFF)
    formula = str(cell.value)
    budget_col = col_letter(VAL_COL_BUDGET)
    sum_col = col_letter(VAL_COL_SUM)
    assert budget_col in formula and sum_col in formula


def test_actuals_total_ceiling_uses_named_range(actuals_workbook):
    """TOTAL row ceiling cell uses partner_ceiling named range."""
    ws = actuals_workbook["Actuals"]
    total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None
    cell = ws.cell(total_row, VAL_COL_BUDGET)
    assert "partner_ceiling" in str(cell.value)
```

### 7d. `tests/test_integration.py`

**Delete these round-trip tests** (no --update mode):
- `test_roundtrip_capacity_edits_preserved`
- `test_roundtrip_actuals_still_preserved`
- `test_roundtrip_capacity_external_edits_preserved`
- `test_update_refuses_empty_config_table`
- `test_roundtrip_team_roster_preserved`

**Keep:**
- `test_capacity_to_plan_formula_linkage` — still valid
- `test_team_shares_to_capacity_seed_values` — still valid

**Remove imports** of `read_workbook`, `read_capacity_sheet`, `read_plan_config` if no longer used.

**Update `generate_workbook` calls** — remove `existing_path` parameter usage.

### 7e. `tests/test_reader_v2.py`

All tests use `discover_layout`, `read_plan_config`, `read_team_sheet`, `read_capacity_sheet`. These functions still exist. Tests should still pass since they test reading from generated workbooks.

**Verify** `discover_layout` returns without `actual_*` keys — update any assertions that check for those keys.

### 7f. `tests/test_cli.py`

**Remove:** Any test for `--update` flag (there aren't any currently — the CLI tests only cover demo, skeleton, snapshot, restore).

**Update `generate_workbook` calls** if any pass `existing_path`.

Check: the current test_cli.py calls `generate_workbook(str(out))` and `generate_workbook(str(out), mode="skeleton")` — no `existing_path`. Should work without changes.

### 7g. `tests/test_e2e.py`

**Update `test_demo_mode_formula_count`:** The formula count will change (removed actual block formulas, added Actuals sheet formulas). Adjust threshold as needed — `total_formulas > 200` should still hold or increase.

**Other e2e tests** should pass without changes (they test recalc zero errors and planned-capacity linkage).

### Verification

```bash
pytest tests/ -v
pyright
```

---

## Task 8: Project Rules Update

**Files:** Modify `.claude/rules/architecture.md`, `.claude/rules/common-tasks.md`, `.claude/rules/dev-workflow.md`

### 8a. `.claude/rules/architecture.md`

- Add `actuals.py` to file tree: `actuals.py          # Actuals sheet: EUR per WP per month from accountant`
- Update `plan.py` description: remove "both FTE blocks", say "globals + config table + planned FTE block"
- Update `summary.py` description: add "budget overview"
- Update `reader.py` description: remove "read_workbook()"
- Update `__main__.py` description: remove `--update`
- Add to test coverage list: "Actuals: WP rows, EUR input, ceiling refs, variance formulas"
- Remove from test coverage: "Actual FTE zeros in fresh workbook", round-trip tests

### 8b. `.claude/rules/common-tasks.md`

- Update "Import actuals from timesheet" entry — clarify that Actuals sheet is for EUR from accountant, not FTE
- Remove any mention of `--update` mode
- Add: "Add actual costs" → Enter EUR values in Actuals sheet blue input cells

### 8c. `.claude/rules/dev-workflow.md`

- Remove "Update mode" section about injecting actuals into test fixtures
- Keep all other rules unchanged

### Verification

```bash
# Rules are documentation — no runtime verification needed
```
