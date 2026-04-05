# P3: Reporting Period Aggregation — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add configurable reporting period boundaries to Config, a Reporting Period Summary section in the Summary sheet that aggregates per-WP Plan/Actual EUR by period — matching the data needed for EU Form C preparation, including CFS threshold flagging — and RP boundary markers on the Trends cumulative chart.

**Architecture:** Add a REPORTING PERIODS config section to the Config sheet (below team roster) with user-editable period boundaries. Extend the Summary sheet with a new section below the Budget Overview that mirrors the existing calendar-year table but groups by reporting period. Add Form C preparation rows (Personnel, Indirect 25%, Total Eligible, Cumulative, CFS flag). Add RP boundary markers on the Trends cumulative EUR chart (green diamond markers at each RP end month). The reader gets a new `read_reporting_periods()` function. The workbook orchestrator passes `reporting_periods` through the build chain.

**Tech Stack:** Python, openpyxl, pytest

**Key files:**
- `budget_planner/config.py` — RP seed data, constants, helper function
- `budget_planner/config_sheet.py` — RP section builder
- `budget_planner/reader.py` — RP reader
- `budget_planner/workbook.py` — orchestrator wiring
- `budget_planner/summary.py` — RP summary section + Form C prep + CFS
- `budget_planner/trends.py` — RP boundary markers on cumulative chart
- `budget_planner/meta.py` — documentation update
- `budget_planner/layout.py` — read-only reference for constants

**Risk notes:**
- Adding a Config section below team roster: no row-shift cascade — team roster is the last section, and the RP section starts after a gap row. Existing readers (`read_config_sheet`, `read_team_from_config`) are position-independent and stop at empty rows.
- `tbl_WorkPackages` Excel Table is not affected (RP section is separate).
- Summary sheet: RP section adds rows below existing BUDGET OVERVIEW. Existing formulas and CF rules in the calendar-year section are row-absolute and won't shift.
- The `test_config_sheet.py` may have assertions about total merged cell count — check after Task 2 and fix any broken assertions.
- The `test_merged_cells.py` (176 lines) may also assert merge counts — check after Task 2.
- CFS threshold (EUR 430k) is hardcoded in formulas. If the programme changes thresholds, this needs a Config param row (separate future enhancement).

**Domain reference:** Read these before implementing:
- `docs/research/research-b-eu-pm-tools-workflows.md` section 5 (reporting periods, Form C, CFS)
- `docs/grant-domain.md` section 6 (indirect costs = 25% flat rate on eligible direct costs excl. subcontracting)

---

## Task 1: Add RP constants, seed data, and helper to config.py

**Files:**
- Modify: `budget_planner/config.py:18` (after EU_ANNUAL_CEILING_PM)

**Step 1: Add constants, seed data, and helper function**

In `budget_planner/config.py`, add after line 18 (`EU_ANNUAL_CEILING_PM = 12.0`):

```python
CFS_THRESHOLD = 430_000                 # EUR — CFS audit certificate required above this
DEFAULT_RP_LENGTH = 6                   # Months per reporting period (programme typical: 12-18)
```

Add after the `EXTERNAL_COMMITMENTS` list (after line 53), before the DEMO MODE section comment:

```python
# ── REPORTING PERIODS ──
# Seed data only — Config sheet is primary after first generation
# Boundaries define when Form C / financial statements are due
REPORTING_PERIODS = [
    {"id": "RP1", "start": 1, "end": 6},
    {"id": "RP2", "start": 7, "end": 12},
    {"id": "RP3", "start": 13, "end": 14},
]
```

Add helper function at the end of the file (after `default_project_params()`):

```python
def default_reporting_periods(project_months, rp_length=DEFAULT_RP_LENGTH):
    """Generate default reporting periods from project duration.

    Returns list of dicts with keys: id, start, end.
    """
    periods = []
    m = 1
    rp_num = 1
    while m <= project_months:
        end = min(m + rp_length - 1, project_months)
        periods.append({"id": f"RP{rp_num}", "start": m, "end": end})
        m = end + 1
        rp_num += 1
    return periods
```

**Step 2: Run tests**

Run: `pytest tests/ -q`
Expected: All pass (no test depends on these new constants yet)

**Step 3: Commit**

```bash
git add budget_planner/config.py
git commit -m "feat(config): add reporting period seed data and helper"
```

---

## Task 2: Add RP section to Config sheet builder

**Files:**
- Modify: `budget_planner/config_sheet.py:140-197` (add RP section after team roster, before named ranges)
- Test: `tests/test_config_sheet.py`

**Step 1: Write the failing test**

Add to `tests/test_config_sheet.py`:

```python
def test_reporting_periods_section(fresh_wb_path):
    """Config sheet has REPORTING PERIODS section with correct data."""
    from openpyxl import load_workbook
    from budget_planner.config import REPORTING_PERIODS

    wb = load_workbook(fresh_wb_path)
    ws = wb["Config"]

    # Find REPORTING PERIODS section header
    rp_header_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIODS" in str(val).upper():
            rp_header_row = r
            break
    assert rp_header_row is not None, "REPORTING PERIODS section not found"

    # Column headers row
    col_header_row = rp_header_row + 1
    assert ws.cell(col_header_row, 1).value == "Period"
    assert ws.cell(col_header_row, 2).value == "Start Month"
    assert ws.cell(col_header_row, 3).value == "End Month"

    # Data rows — should match seed data count
    first_data_row = col_header_row + 1
    for i, rp in enumerate(REPORTING_PERIODS):
        row = first_data_row + i
        assert ws.cell(row, 1).value == rp["id"]
        assert ws.cell(row, 2).value == rp["start"]
        assert ws.cell(row, 3).value == rp["end"]
        # Blue input font
        assert ws.cell(row, 1).font.color.rgb == "000000FF"

    wb.close()
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_config_sheet.py::test_reporting_periods_section -v`
Expected: FAIL — "REPORTING PERIODS section not found"

**Step 3: Implement RP section in config_sheet.py**

In `budget_planner/config_sheet.py`, update the function signature to accept `reporting_periods`:

```python
def build_config_sheet(ws, wb, project_params, work_packages, team=None,
                       reporting_periods=None):
```

Add after the team roster loop (after line 166, after `roster_last_row = r - 1`), before the Named Ranges section:

```python
    # ── Section D: Reporting Periods ──
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS

    r = roster_last_row + 2  # gap row after roster

    # Section header — merged A:C
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    style_cell(ws, r, 1, value="REPORTING PERIODS",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    for c in range(2, 4):
        style_cell(ws, r, c, fill=SECTION_FILL)
    r += 1

    # Column headers
    rp_col_header_row = r
    for col, hdr in enumerate(["Period", "Start Month", "End Month"], start=1):
        style_cell(ws, r, col, hdr, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    # RP data rows
    rp_first_row = r
    for rp in reporting_periods:
        style_cell(ws, r, 1, rp["id"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, 2, rp["start"], font=BLUE_INPUT, fmt="0", border=False)
        style_cell(ws, r, 3, rp["end"], font=BLUE_INPUT, fmt="0", border=False)
        r += 1
    rp_last_row = r - 1
```

Update the return dict (at end of function) to include RP positions:

```python
    return {
        "param_rows": param_rows,
        "wp_header_row": header_row,
        "wp_first_row": first_data_row,
        "wp_last_row": last_data_row,
        "wp_total_row": wp_total_row,
        "roster_header_row": roster_header_row,
        "roster_first_row": roster_first_row,
        "roster_last_row": roster_last_row,
        "person_rows": person_rows,
        "rp_first_row": rp_first_row,
        "rp_last_row": rp_last_row,
    }
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_config_sheet.py::test_reporting_periods_section -v`
Expected: PASS

**Step 5: Run full test suite to check for regressions**

Run: `pytest tests/ -q`
Expected: All pass. If `test_merged_cells.py` or `test_config_sheet.py` fail on merge count assertions, update those assertions to account for the new REPORTING PERIODS merged header.

**Step 6: Commit**

```bash
git add budget_planner/config_sheet.py tests/test_config_sheet.py
git commit -m "feat(config-sheet): add REPORTING PERIODS section"
```

---

## Task 3: Add RP reader to reader.py

**Files:**
- Modify: `budget_planner/reader.py` (add `read_reporting_periods` function)
- Test: `tests/test_reader_new.py`

**Step 1: Write the failing test**

Add to `tests/test_reader_new.py`:

```python
def test_read_reporting_periods(fresh_wb_path):
    """Reader extracts reporting periods from Config sheet."""
    from openpyxl import load_workbook
    from budget_planner.reader import read_reporting_periods
    from budget_planner.config import REPORTING_PERIODS

    wb = load_workbook(fresh_wb_path)
    ws = wb["Config"]
    rps = read_reporting_periods(ws)
    wb.close()

    assert len(rps) == len(REPORTING_PERIODS)
    for actual, expected in zip(rps, REPORTING_PERIODS):
        assert actual["id"] == expected["id"]
        assert actual["start"] == expected["start"]
        assert actual["end"] == expected["end"]


def test_read_reporting_periods_fallback():
    """Reader returns default periods when section is missing."""
    from openpyxl import Workbook
    from budget_planner.reader import read_reporting_periods

    wb = Workbook()
    ws = wb.active
    ws.title = "Config"
    # Empty sheet — no REPORTING PERIODS section
    ws.cell(1, 1).value = "Project Title"
    ws.cell(1, 2).value = "Test"

    rps = read_reporting_periods(ws, fallback_months=14)
    wb.close()

    # Should get default 6-month periods for 14-month project
    assert len(rps) == 3
    assert rps[0] == {"id": "RP1", "start": 1, "end": 6}
    assert rps[1] == {"id": "RP2", "start": 7, "end": 12}
    assert rps[2] == {"id": "RP3", "start": 13, "end": 14}
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_reader_new.py::test_read_reporting_periods -v`
Expected: FAIL — ImportError (function doesn't exist yet)

**Step 3: Implement read_reporting_periods**

Add to `budget_planner/reader.py` (after `read_team_from_config`):

```python
def read_reporting_periods(ws, fallback_months=None):
    """Read reporting period definitions from Config sheet.

    Args:
        ws: Config worksheet.
        fallback_months: If section not found, generate default periods
            for this many months. If None and section missing, returns [].

    Returns:
        List of dicts with keys: id, start, end.
    """
    rp_section = _find_header_row(ws, "REPORTING PERIODS", required=False)
    if rp_section is None:
        if fallback_months is not None:
            from budget_planner.config import default_reporting_periods
            return default_reporting_periods(fallback_months)
        return []

    # Skip section header + column headers
    r = rp_section + 2
    periods: list[dict] = []
    while r <= ws.max_row:
        period_id = ws.cell(r, 1).value
        if period_id is None or str(period_id).strip() == "":
            break
        start = ws.cell(r, 2).value
        end = ws.cell(r, 3).value
        periods.append({
            "id": str(period_id),
            "start": int(start or 1),
            "end": int(end or 1),
        })
        r += 1

    return periods
```

**Step 4: Run tests to verify they pass**

Run: `pytest tests/test_reader_new.py::test_read_reporting_periods tests/test_reader_new.py::test_read_reporting_periods_fallback -v`
Expected: PASS

**Step 5: Run full suite**

Run: `pytest tests/ -q`
Expected: All pass

**Step 6: Commit**

```bash
git add budget_planner/reader.py tests/test_reader_new.py
git commit -m "feat(reader): add read_reporting_periods with fallback"
```

---

## Task 4: Wire RP through workbook.py orchestrator

**Files:**
- Modify: `budget_planner/workbook.py:99-101` (`_build_all_sheets` signature)
- Modify: `budget_planner/workbook.py:110-113` (pass RP to `build_config_sheet`)
- Modify: `budget_planner/workbook.py:165-167` (pass RP to `build_summary_sheet`)
- Modify: `budget_planner/workbook.py:170-174` (pass RP to `build_trends_sheet`)
- Modify: `budget_planner/workbook.py:236-254` (`init_workbook` — pass seed RPs)
- Modify: `budget_planner/workbook.py:287-344` (`rebuild_workbook` — read RPs from Config)

**Step 1: Update `_build_all_sheets` to accept and pass `reporting_periods`**

In `_build_all_sheets` signature (line 99), add `reporting_periods=None`:

```python
def _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      *, assignments=None, preserved_actuals=None,
                      rebuild_warnings=None, reporting_periods=None):
```

In the Config sheet build call (line 112), pass `reporting_periods`:

```python
    config_layout = build_config_sheet(config_ws, wb, project_params, work_packages,
                                       team=team,
                                       reporting_periods=reporting_periods)
```

In the Summary sheet build call (line 165), pass `reporting_periods`:

```python
    build_summary_sheet(wb, layout, actuals_layout,
                        project_params=project_params,
                        work_packages=work_packages,
                        reporting_periods=reporting_periods)
```

In the Trends sheet build call (line 170), pass `reporting_periods`:

```python
    build_trends_sheet(trends_ws, wb, layout, actuals_layout,
                       capacity_layout, config_layout,
                       project_params=project_params,
                       work_packages=work_packages,
                       team=team,
                       reporting_periods=reporting_periods)
```

**Step 2: Update `init_workbook` to pass seed RPs**

In `init_workbook` (around line 245), add after `external_commitments = config.EXTERNAL_COMMITMENTS`:

```python
    reporting_periods = config.REPORTING_PERIODS
```

In the `_build_all_sheets` call (line 254), add `reporting_periods=reporting_periods`:

```python
    _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      assignments=assignments,
                      reporting_periods=reporting_periods)
```

**Step 3: Update `rebuild_workbook` to read RPs from Config**

In `rebuild_workbook`, add the import at the top of the function (around line 289):

```python
    from budget_planner.reader import (
        read_config_sheet, read_team_from_config, read_capacity_sheet,
        read_actuals_values, read_reporting_periods,
    )
```

After reading team (around line 315), add:

```python
    # Reporting periods from Config sheet
    reporting_periods = read_reporting_periods(
        existing_wb[SHEET_CONFIG],
        fallback_months=project_params["months"],
    )
```

In the `_build_all_sheets` call (around line 391), add `reporting_periods=reporting_periods`:

```python
    _build_all_sheets(
        wb, project_params, work_packages, team, external_commitments,
        assignments=assignments,
        preserved_actuals=preserved_actuals,
        rebuild_warnings=warnings if warnings else None,
        reporting_periods=reporting_periods,
    )
```

**Step 4: Run full test suite**

Run: `pytest tests/ -q`
Expected: All pass (Summary builder accepts but ignores `reporting_periods` until Task 5)

**Step 5: Commit**

```bash
git add budget_planner/workbook.py
git commit -m "feat(workbook): wire reporting_periods through init and rebuild"
```

---

## Task 5: Add RP summary section to Summary sheet

**Files:**
- Modify: `budget_planner/summary.py:19` (add `reporting_periods` parameter)
- Modify: `budget_planner/summary.py:245` (add RP section after freeze_panes — move freeze_panes to end)
- Test: `tests/test_integration.py` (or new test file)

**Step 1: Write the failing test**

Add to `tests/test_integration.py` (or create a new `tests/test_summary.py` if preferred):

```python
def test_rp_summary_section_exists(fresh_wb_path):
    """Summary sheet has REPORTING PERIOD SUMMARY section."""
    from openpyxl import load_workbook

    wb = load_workbook(fresh_wb_path)
    ws = wb["Summary"]

    # Find REPORTING PERIOD SUMMARY section header
    rp_header_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIOD" in str(val).upper():
            rp_header_row = r
            break
    assert rp_header_row is not None, "REPORTING PERIOD SUMMARY section not found"

    wb.close()


def test_rp_summary_has_wp_rows(fresh_wb_path):
    """RP summary has one row per WP with cross-sheet formula refs."""
    from openpyxl import load_workbook
    from budget_planner.config import WORK_PACKAGES

    wb = load_workbook(fresh_wb_path)
    ws = wb["Summary"]

    # Find section header
    rp_header_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIOD" in str(val).upper():
            rp_header_row = r
            break
    assert rp_header_row is not None

    # Find first WP data row (skip header rows — look for Config!A cross-ref)
    first_wp_row = None
    for r in range(rp_header_row + 1, rp_header_row + 10):
        val = str(ws.cell(r, 1).value or "")
        if "Config!" in val:
            first_wp_row = r
            break
    assert first_wp_row is not None, "No WP data rows found in RP summary"

    # Count WP rows
    wp_count = 0
    for r in range(first_wp_row, first_wp_row + len(WORK_PACKAGES) + 5):
        val = str(ws.cell(r, 1).value or "")
        if "Config!" in val:
            wp_count += 1
        else:
            break
    assert wp_count == len(WORK_PACKAGES)

    # First WP row col C should be a Plan formula referencing Plan sheet
    val_c = str(ws.cell(first_wp_row, 3).value or "")
    assert "Plan!" in val_c, f"Expected Plan! cross-ref, got: {val_c}"

    wb.close()
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_integration.py::test_rp_summary_section_exists -v`
Expected: FAIL — "REPORTING PERIOD SUMMARY section not found"

**Step 3: Implement RP summary section in summary.py**

Update `build_summary_sheet` signature (line 19) to accept `reporting_periods`:

```python
def build_summary_sheet(wb, layout, actuals_layout, *, project_params=None,
                        work_packages=None, reporting_periods=None):
```

Add lazy import fallback (after existing fallbacks, around line 39):

```python
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS
```

Move `summary.freeze_panes = f"C{SUMMARY_FIRST_WP_ROW}"` to the very end of the function (it currently is the last line).

Add the RP summary section after the existing Budget Overview CF rules (after the `CF_AMBER_FILL` rule, before freeze_panes). Here `r` continues from the `% Consumed` row:

```python
    # ── REPORTING PERIOD SUMMARY ──
    r += 2
    write_section_header(summary, r, "REPORTING PERIOD SUMMARY", merge_end=2)
    r += 2

    num_rps = len(reporting_periods)
    RP_GROUP_START = 3
    COLS_PER_RP = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR

    RP_TOTAL_GROUP_START = RP_GROUP_START + num_rps * COLS_PER_RP
    RP_TOTAL_COLS = 4

    # RP headers row
    style_cell(summary, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(summary, r, 2, "", fill=LIGHT_BLUE_FILL)
    for rp_idx, rp in enumerate(reporting_periods):
        start_col = RP_GROUP_START + rp_idx * COLS_PER_RP
        end_col = start_col + COLS_PER_RP - 1
        label = f"{rp['id']} (M{rp['start']}\u2013M{rp['end']})"
        summary.merge_cells(start_row=r, start_column=start_col,
                            end_row=r, end_column=end_col)
        style_cell(summary, r, start_col, label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(start_col + 1, end_col + 1):
            style_cell(summary, r, c2, fill=LIGHT_BLUE_FILL)

    summary.merge_cells(start_row=r, start_column=RP_TOTAL_GROUP_START,
                        end_row=r, end_column=RP_TOTAL_GROUP_START + RP_TOTAL_COLS - 1)
    style_cell(summary, r, RP_TOTAL_GROUP_START,
               font=Font(name="Arial", bold=True, size=10, color="FFFFFF"),
               fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(RP_TOTAL_GROUP_START + 1, RP_TOTAL_GROUP_START + RP_TOTAL_COLS):
        style_cell(summary, r, c2, fill=SECTION_FILL)
    r += 1

    # RP sub-headers row
    style_cell(summary, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(summary, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    rp_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for rp_idx in range(num_rps):
        for si, sh in enumerate(rp_sub):
            c = RP_GROUP_START + rp_idx * COLS_PER_RP + si
            style_cell(summary, r, c, sh, font=HEADER_FONT,
                       fill=LIGHT_BLUE_FILL, alignment=CENTER)
    rp_total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for si, sh in enumerate(rp_total_sub):
        c = RP_TOTAL_GROUP_START + si
        style_cell(summary, r, c, sh, font=HEADER_FONT,
                   fill=LIGHT_GREEN_FILL, alignment=CENTER)
    r += 1

    RP_FIRST_WP_ROW = r

    # RP WP data rows
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = config_first_wp_row + i
        style_cell(summary, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(summary, row, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        for rp_idx, rp in enumerate(reporting_periods):
            first_col_l = col_letter(month_col_start + rp["start"] - 1)
            last_col_l = col_letter(month_col_start + rp["end"] - 1)

            sc = RP_GROUP_START + rp_idx * COLS_PER_RP
            plan_wp_row = planned_first_wp_row + i
            actuals_wp_row = actuals_first_wp_row + i

            # Plan FTE
            style_cell(summary, row, sc, font=GREEN_XREF, fmt="0.00").value = (
                f"=SUM({SHEET_PLAN}!{first_col_l}{plan_wp_row}:"
                f"{last_col_l}{plan_wp_row})"
            )
            # Plan EUR = Plan FTE * fte_rate
            style_cell(summary, row, sc + 1, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0").value = (
                f"={col_letter(sc)}{row}*fte_rate"
            )
            # Actual EUR
            style_cell(summary, row, sc + 2, font=GREEN_XREF,
                       fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_ACTUALS}!{first_col_l}{actuals_wp_row}:"
                f"{last_col_l}{actuals_wp_row})"
            )
            # Variance EUR = Plan EUR - Actual EUR
            style_cell(summary, row, sc + 3, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0").value = (
                f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
            )

        # RP Project totals (sum across RP groups)
        tc = RP_TOTAL_GROUP_START
        rp_plan_cols = [col_letter(RP_GROUP_START + rp_idx * COLS_PER_RP)
                        for rp_idx in range(num_rps)]
        style_cell(summary, row, tc, font=BLACK_FORMULA, fmt="0.00").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_plan_cols)
        )
        rp_plan_eur_cols = [col_letter(RP_GROUP_START + rp_idx * COLS_PER_RP + 1)
                            for rp_idx in range(num_rps)]
        style_cell(summary, row, tc + 1, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_plan_eur_cols)
        )
        rp_act_eur_cols = [col_letter(RP_GROUP_START + rp_idx * COLS_PER_RP + 2)
                           for rp_idx in range(num_rps)]
        style_cell(summary, row, tc + 2, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_act_eur_cols)
        )
        style_cell(summary, row, tc + 3, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(tc + 1)}{row}-{col_letter(tc + 2)}{row}"
        )

    RP_LAST_WP_ROW = RP_FIRST_WP_ROW + num_wps - 1
    rp_totals_row = RP_LAST_WP_ROW + 1

    # RP TOTAL row
    style_cell(summary, rp_totals_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(summary, rp_totals_row, 2, "TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    for rp_idx in range(num_rps):
        for si in range(COLS_PER_RP):
            c = RP_GROUP_START + rp_idx * COLS_PER_RP + si
            fmt = "0.00" if si == 0 else "\u20ac#,##0"
            col_range = (f"{col_letter(c)}{RP_FIRST_WP_ROW}:"
                         f"{col_letter(c)}{RP_LAST_WP_ROW}")
            style_cell(summary, rp_totals_row, c, font=BLACK_FORMULA,
                       fill=LIGHT_GREEN_FILL, fmt=fmt).value = (
                f"=SUM({col_range})"
            )
    for si in range(RP_TOTAL_COLS):
        c = RP_TOTAL_GROUP_START + si
        fmt = "0.00" if si == 0 else "\u20ac#,##0"
        col_range = (f"{col_letter(c)}{RP_FIRST_WP_ROW}:"
                     f"{col_letter(c)}{RP_LAST_WP_ROW}")
        style_cell(summary, rp_totals_row, c, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL, fmt=fmt).value = (
            f"=SUM({col_range})"
        )

    # CF: red on negative RP Variance EUR
    for rp_idx in range(num_rps):
        var_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 3
        var_col_l = col_letter(var_col)
        var_range = (f"{var_col_l}{RP_FIRST_WP_ROW}:"
                     f"{var_col_l}{RP_LAST_WP_ROW}")
        summary.conditional_formatting.add(var_range, FormulaRule(
            formula=[f"{var_col_l}{RP_FIRST_WP_ROW}<0"], fill=CF_RED_FILL,
        ))
    # CF: red on negative PROJECT TOTAL Variance EUR
    rp_total_var_col = RP_TOTAL_GROUP_START + 3
    rp_total_var_col_l = col_letter(rp_total_var_col)
    rp_total_var_range = (f"{rp_total_var_col_l}{RP_FIRST_WP_ROW}:"
                          f"{rp_total_var_col_l}{RP_LAST_WP_ROW}")
    summary.conditional_formatting.add(rp_total_var_range, FormulaRule(
        formula=[f"{rp_total_var_col_l}{RP_FIRST_WP_ROW}<0"], fill=CF_RED_FILL,
    ))

    r = rp_totals_row
```

**Step 4: Run tests to verify they pass**

Run: `pytest tests/test_integration.py::test_rp_summary_section_exists tests/test_integration.py::test_rp_summary_has_wp_rows -v`
Expected: PASS

**Step 5: Run full suite**

Run: `pytest tests/ -q`
Expected: All pass

**Step 6: Commit**

```bash
git add budget_planner/summary.py tests/test_integration.py
git commit -m "feat(summary): add REPORTING PERIOD SUMMARY section"
```

---

## Task 6: Add Form C preparation rows + CFS threshold

**Files:**
- Modify: `budget_planner/summary.py` (add Form C rows after RP TOTAL row)
- Test: `tests/test_integration.py`

**Step 1: Write the failing test**

Add to `tests/test_integration.py`:

```python
def test_form_c_rows_exist(fresh_wb_path):
    """Summary sheet has Form C prep rows with indirect costs and CFS flag."""
    from openpyxl import load_workbook

    wb = load_workbook(fresh_wb_path)
    ws = wb["Summary"]

    # Find "A. Personnel" label
    personnel_row = None
    for r in range(1, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or "")
        if "Personnel" in val and val.startswith("A."):
            personnel_row = r
            break
    assert personnel_row is not None, "Form C 'A. Personnel' row not found"

    # Check subsequent rows
    assert "Indirect" in str(ws.cell(personnel_row + 1, 1).value or "")
    assert "Total Eligible" in str(ws.cell(personnel_row + 2, 1).value or "")
    assert "Cum" in str(ws.cell(personnel_row + 3, 1).value or "")
    assert "CFS" in str(ws.cell(personnel_row + 4, 1).value or "")

    # CFS cell should have a formula with 430000
    cfs_row = personnel_row + 4
    # Check col C (first RP Act EUR position)
    cfs_val = str(ws.cell(cfs_row, 3 + 2).value or "")  # offset +2 = Act EUR col
    assert "430000" in cfs_val or "CFS" in str(ws.cell(cfs_row, 1).value)

    wb.close()
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_integration.py::test_form_c_rows_exist -v`
Expected: FAIL — "Form C 'A. Personnel' row not found"

**Step 3: Implement Form C rows in summary.py**

Add after the RP CF rules (at the end of the RP section, where `r = rp_totals_row`):

```python
    # ── FORM C PREPARATION ──
    r = rp_totals_row + 2

    # Row labels and formulas use only the Act EUR sub-column (offset +2) per RP
    # A. Personnel EUR — reference TOTAL Act EUR for each RP
    personnel_fc_row = r
    style_cell(summary, r, 1, "A. Personnel EUR", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{rp_totals_row}"
        )
    # Project total
    total_act_col = RP_TOTAL_GROUP_START + 2
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{rp_totals_row}"
    )
    r += 1

    # E. Indirect (25%) — 25% of Personnel EUR
    indirect_row = r
    style_cell(summary, r, 1, "E. Indirect (25%)", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{personnel_fc_row}*0.25"
        )
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{personnel_fc_row}*0.25"
    )
    r += 1

    # Total Eligible = Personnel + Indirect
    eligible_row = r
    style_cell(summary, r, 1, "Total Eligible", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{personnel_fc_row}"
            f"+{col_letter(act_col)}{indirect_row}"
        )
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{personnel_fc_row}"
        f"+{col_letter(total_act_col)}{indirect_row}"
    )
    r += 1

    # Cumulative Eligible — running sum across RPs
    cum_row = r
    style_cell(summary, r, 1, "Cum. Eligible", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        if rp_idx == 0:
            style_cell(summary, r, act_col, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0", border=False).value = (
                f"={col_letter(act_col)}{eligible_row}"
            )
        else:
            prev_act_col = RP_GROUP_START + (rp_idx - 1) * COLS_PER_RP + 2
            style_cell(summary, r, act_col, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0", border=False).value = (
                f"={col_letter(prev_act_col)}{cum_row}"
                f"+{col_letter(act_col)}{eligible_row}"
            )
    # Project total cumulative = same as last RP cumulative (or total eligible)
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{eligible_row}"
    )
    r += 1

    # CFS Required? — flag when cumulative > EUR 430k
    cfs_row = r
    style_cell(summary, r, 1, "CFS Required?", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   border=False).value = (
            f'=IF({col_letter(act_col)}{cum_row}>430000,"YES","No")'
        )
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               border=False).value = (
        f'=IF({col_letter(total_act_col)}{cum_row}>430000,"YES","No")'
    )

    # CF on CFS Required: red fill on "YES"
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        cell_ref = f"{col_letter(act_col)}{cfs_row}"
        summary.conditional_formatting.add(cell_ref, FormulaRule(
            formula=[f'{cell_ref}="YES"'], fill=CF_AMBER_FILL,
        ))
    total_cfs_ref = f"{col_letter(total_act_col)}{cfs_row}"
    summary.conditional_formatting.add(total_cfs_ref, FormulaRule(
        formula=[f'{total_cfs_ref}="YES"'], fill=CF_AMBER_FILL,
    ))
```

Move `summary.freeze_panes` to the very end of the function (it should remain `f"C{SUMMARY_FIRST_WP_ROW}"` — freezing on the calendar-year WP table is correct).

**Step 4: Run tests to verify they pass**

Run: `pytest tests/test_integration.py::test_form_c_rows_exist -v`
Expected: PASS

**Step 5: Run full suite**

Run: `pytest tests/ -q`
Expected: All pass

**Step 6: Commit**

```bash
git add budget_planner/summary.py tests/test_integration.py
git commit -m "feat(summary): add Form C prep rows with CFS threshold"
```

---

## Task 7: Update meta.py

**Files:**
- Modify: `budget_planner/meta.py`

**Step 1: Read current meta.py content**

Read `budget_planner/meta.py` in full to find where sheet documentation blocks are defined.

**Step 2: Add documentation for Config RP section and Summary RP section**

Add to the Config sheet documentation block:

```
REPORTING PERIODS section (below Team Roster):
  Col A: Period ID (e.g. RP1) — blue input
  Col B: Start Month — blue input
  Col C: End Month — blue input
  User-editable: add/remove/modify periods for Form C reporting boundaries.
```

Add to the Summary sheet documentation block:

```
REPORTING PERIOD SUMMARY section (below Budget Overview):
  Same 4-column-per-group layout as the annual summary above,
  but grouped by user-defined reporting periods from Config.
  Columns per RP: Plan FTE, Plan EUR, Act EUR, Var EUR
  PROJECT TOTAL group at the end.

  FORM C PREPARATION rows below RP TOTAL:
    A. Personnel EUR — actual personnel costs per reporting period
    E. Indirect (25%) — 25% flat rate on personnel (minimum; excludes travel/equipment)
    Total Eligible — A + E
    Cum. Eligible — running cumulative across reporting periods
    CFS Required? — "YES" if cumulative eligible > EUR 430,000 (programme audit threshold)

  Note: Indirect costs shown are a lower bound (personnel-only base).
  Actual indirect should include travel, equipment, and other direct costs.
```

**Step 3: Verify meta text matches actual builder output**

Count columns in summary.py's RP section (should be 4 per RP + 4 for project total). Verify against meta text.

**Step 4: Run full suite**

Run: `pytest tests/ -q`
Expected: All pass

**Step 5: Commit**

```bash
git add budget_planner/meta.py
git commit -m "docs(meta): document Config RP section and Summary RP section"
```

---

## Task 8: Rebuild round-trip test

**Files:**
- Test: `tests/test_rebuild.py` or `tests/test_rebuild_e2e.py`

**Step 1: Write the failing test**

Add to `tests/test_rebuild_e2e.py`:

```python
def test_reporting_periods_survive_rebuild(tmp_path):
    """Reporting periods in Config survive a rebuild cycle."""
    from budget_planner.workbook import init_workbook, rebuild_workbook
    from budget_planner.reader import read_reporting_periods
    from budget_planner.config import REPORTING_PERIODS
    from openpyxl import load_workbook

    out = str(tmp_path / "rp_rebuild.xlsx")
    init_workbook(out)

    # Read original RPs
    wb = load_workbook(out)
    original_rps = read_reporting_periods(wb["Config"])
    wb.close()
    assert len(original_rps) == len(REPORTING_PERIODS)

    # Rebuild
    rebuild_workbook(out)

    # Read rebuilt RPs
    wb = load_workbook(out)
    rebuilt_rps = read_reporting_periods(wb["Config"])
    wb.close()

    assert len(rebuilt_rps) == len(original_rps)
    for orig, rebuilt in zip(original_rps, rebuilt_rps):
        assert orig["id"] == rebuilt["id"]
        assert orig["start"] == rebuilt["start"]
        assert orig["end"] == rebuilt["end"]


def test_modified_reporting_periods_survive_rebuild(tmp_path):
    """User-modified reporting periods are preserved on rebuild."""
    from budget_planner.workbook import init_workbook, rebuild_workbook
    from budget_planner.reader import read_reporting_periods
    from openpyxl import load_workbook

    out = str(tmp_path / "rp_mod.xlsx")
    init_workbook(out)

    # Modify RP boundaries in the workbook
    wb = load_workbook(out)
    ws = wb["Config"]

    # Find REPORTING PERIODS section
    rp_section = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIODS" in str(val).upper():
            rp_section = r
            break
    assert rp_section is not None

    # Change RP1 end from 18 to 12
    data_row = rp_section + 2  # skip header + col headers
    ws.cell(data_row, 3).value = 12  # End Month = 12

    wb.save(out)
    wb.close()

    # Rebuild
    rebuild_workbook(out)

    # Verify modification preserved
    wb = load_workbook(out)
    rps = read_reporting_periods(wb["Config"])
    wb.close()

    assert rps[0]["end"] == 12  # Modified value preserved
```

**Step 2: Run tests to verify they pass**

Run: `pytest tests/test_rebuild_e2e.py::test_reporting_periods_survive_rebuild tests/test_rebuild_e2e.py::test_modified_reporting_periods_survive_rebuild -v`
Expected: PASS (if all previous tasks were implemented correctly)

**Step 3: Run full suite**

Run: `pytest tests/ -q`
Expected: All pass

**Step 4: Commit**

```bash
git add tests/test_rebuild_e2e.py
git commit -m "test(rebuild): verify reporting periods survive round-trip"
```

---

## Task 9: Add RP boundary markers to Trends cumulative chart

**Files:**
- Modify: `budget_planner/trends.py:28-29` (add `reporting_periods` parameter)
- Modify: `budget_planner/trends.py:340-454` (add helper row + chart series for RP markers)
- Test: `tests/test_trends.py`

**Step 1: Write the failing test**

Add to `tests/test_trends.py`:

```python
def test_chart_has_rp_boundary_series(fresh_wb_path):
    """Trends cumulative chart has an RP boundary marker series."""
    from openpyxl import load_workbook

    wb = load_workbook(fresh_wb_path)
    ws = wb["Trends"]

    # Find the chart
    assert len(ws._charts) == 1
    chart = ws._charts[0]

    # Should have 4 series: Planned, Actual, Ceiling, RP End
    assert len(chart.series) == 4, (
        f"Expected 4 chart series, got {len(chart.series)}"
    )

    # Last series should be "RP End"
    rp_series = chart.series[3]
    assert rp_series.tx is not None
    # SeriesLabel value should contain "RP"
    label_val = getattr(rp_series.tx, 'v', None)
    assert label_val is not None and "RP" in label_val

    wb.close()
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_trends.py::test_chart_has_rp_boundary_series -v`
Expected: FAIL — "Expected 4 chart series, got 3"

**Step 3: Implement RP boundary markers in trends.py**

Update function signature (line 28) to accept `reporting_periods`:

```python
def build_trends_sheet(ws, wb, layout, actuals_layout, capacity_layout, team_layout, *,
                       project_params=None, work_packages=None, team=None,
                       reporting_periods=None):
```

Add lazy import fallback (after the existing None checks, around line 50):

```python
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS
```

After the existing helper row setup and BEFORE the chart is built (insert after line 418, after `ws.row_dimensions[helper_row_labels].hidden = True`), add:

```python
    # Helper Row 5: RP boundary markers (ceiling value at each RP end month)
    helper_row_rp = helper_row_labels + 1
    for rp in reporting_periods:
        boundary_col = month_col_start + rp["end"] - 1
        if boundary_col <= month_col_end:
            style_cell(ws, helper_row_rp, boundary_col,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                "=partner_ceiling"
            )
    ws.row_dimensions[helper_row_rp].hidden = True
```

After the existing `chart.series.append(s_ceiling)` (line 449), add the RP series:

```python
    # Series 4: RP boundaries (green diamond markers at ceiling level)
    from openpyxl.chart.marker import Marker
    s_rp = Series(idx=3, order=3)
    s_rp.tx = SeriesLabel(v="RP End")
    s_rp.val = NumDataSource(numRef=NumRef(
        f=range_str.format(helper_row_rp, helper_row_rp)))
    s_rp.cat = cat_ref
    s_rp.graphicalProperties.line.solidFill = "00B050"
    s_rp.graphicalProperties.line.dashStyle = "dot"
    s_rp.marker = Marker(symbol="diamond", size=10)
    s_rp.marker.graphicalProperties.solidFill = "00B050"
    chart.series.append(s_rp)
```

Update the return dict to include the new helper row:

```python
    return {
        ...
        "helper_row_labels": helper_row_labels,
        "helper_row_rp": helper_row_rp,
    }
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_trends.py::test_chart_has_rp_boundary_series -v`
Expected: PASS

**Step 5: Run full suite**

Run: `pytest tests/ -q`
Expected: All pass. Check that `test_trends.py` tests referencing helper row counts or chart structure still pass. If any test asserts exactly 3 series, update to 4.

**Step 6: Commit**

```bash
git add budget_planner/trends.py tests/test_trends.py
git commit -m "feat(trends): add RP boundary markers on cumulative chart"
```

---

## Notes

### What this plan deliberately does NOT build (YAGNI)
- **Variance percentage classification** (P2 scope — marked as won't implement)
- **Deviation narrative helper** — generates pre-formatted text for WPs with large deviations. Useful but over-automates Form C preparation.
- **Period-over-period trend** — showing variance direction across RPs. Could be added later.
- **CFS threshold as Config param** — hardcoded at EUR 430k. If programme changes, add a Config row later.
- **Separate Reporting sheet** — the RP section in Summary is sufficient. A dedicated sheet would duplicate data.

### Indirect cost caveat
The 25% indirect calculation in Form C prep uses only personnel costs (category A) as the base. The actual EU indirect cost base is `A + C.1 + C.2 + C.3 + D` (excluding subcontracting). Since this tool only tracks personnel, the indirect shown is a **lower bound**. This is documented in the meta.py text and is appropriate for a planning tool.

### Testing strategy
- Task 2: Config sheet structure test (RP section exists with correct data)
- Task 3: Reader unit test (extraction + fallback)
- Task 5: Integration test (RP summary section exists with correct formula refs)
- Task 6: Integration test (Form C rows + CFS formula)
- Task 8: Rebuild round-trip test (RPs preserved, modifications preserved)
- Task 9: Trends chart series count test (4 series including RP End)

Total new tests: ~7 test functions across 4 test files.
