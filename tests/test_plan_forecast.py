"""Tests for the Plan-Forecast sheet builder (T5).

Plan-Forecast aggregates forecast-EUR leaves from Actuals-Hours into a
WP × month grid that mirrors the Plan sheet shape.
"""
import re

import pytest
from openpyxl import load_workbook
from openpyxl.formatting.rule import Rule

from budget_planner.config import (
    WORK_PACKAGES,
    WP_SCHEDULE,
    PROJECT_MONTHS,
)
from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_PLAN_FORECAST,
    val_columns,
    month_col_end,
    wp_active_months,
)
from budget_planner.styles import CF_GREY_FILL, GREY_FILL, col_letter


# ── Fixtures ──

@pytest.fixture(scope="module")
def assigned_wb(assigned_wb_path):
    return load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def empty_wb(fresh_wb_path):
    return load_workbook(fresh_wb_path, data_only=False)


@pytest.fixture(scope="module")
def pf(assigned_wb):
    return assigned_wb[SHEET_PLAN_FORECAST]


# ── Existence + tab order ──

def test_plan_forecast_sheet_exists(assigned_wb):
    assert SHEET_PLAN_FORECAST in assigned_wb.sheetnames


def test_plan_forecast_tab_order(assigned_wb):
    """Plan-Forecast must sit immediately after Actuals-Hours."""
    names = assigned_wb.sheetnames
    assert names.index(SHEET_PLAN_FORECAST) == names.index("Actuals-Hours") + 1


# ── Shape ──

def test_plan_forecast_wp_count(pf):
    """One WP row per WORK_PACKAGES entry.

    Derives WP-row count by scanning col A for green cross-refs to Config
    (``=Config!A<n>``). The WP block sits below the compact header + section
    header + year/month rows; scanning catches all of them without
    hard-coding row positions.
    """
    wp_id_refs = 0
    for r in range(1, pf.max_row + 1):
        v = pf.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^=Config!A\d+$", v):
            wp_id_refs += 1
    assert wp_id_refs == len(WORK_PACKAGES)


# ── Formula content ──

def _find_wp_row(ws, wp_idx):
    """Return the Plan-Forecast row for WORK_PACKAGES[wp_idx].

    Walks col-A cross-refs in order; the nth one corresponds to the nth WP.
    """
    count = 0
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^=Config!A\d+$", v):
            if count == wp_idx:
                return r
            count += 1
    raise AssertionError(f"No WP row for idx {wp_idx}")


def test_plan_forecast_cell_sums_actuals_hours_leaves(pf):
    """WP3 has Alice + Dave assigned (2 persons).

    For an active month, the cell formula must be:
        =SUM('Actuals-Hours'!<cellA>, 'Actuals-Hours'!<cellB>)

    with correct row positions (one per assignee WP3 row on Actuals-Hours)
    and identical columns (both at forecast_col_start + month_idx).
    """
    wp3_idx = 2
    wp_id = WORK_PACKAGES[wp3_idx]["id"]
    assert wp_id == "WP3"

    # Pick first active month in WP3
    active = sorted(wp_active_months(wp_id, WP_SCHEDULE))
    month_num = active[0]  # M1
    m_idx = month_num - 1
    col = MONTH_COL_START + m_idx

    row = _find_wp_row(pf, wp3_idx)
    formula = pf.cell(row, col).value
    assert isinstance(formula, str), f"Expected formula string, got {formula!r}"
    assert formula.startswith("=SUM("), formula
    assert "Actuals-Hours" in formula, formula

    # Extract the referenced row numbers — they must be distinct (one per
    # assignee) and the column must be the same across refs.
    refs = re.findall(r"Actuals-Hours'?!([A-Z]+)(\d+)", formula)
    # 2 assignees for WP3 in demo: Alice + Dave
    assert len(refs) == 2, f"Expected 2 leaf refs, got {refs} in {formula!r}"
    cols = {r[0] for r in refs}
    rows = {r[1] for r in refs}
    assert len(cols) == 1, f"Leaves must share a column; got {cols}"
    assert len(rows) == 2, f"Leaves must point at distinct rows; got {rows}"


def test_plan_forecast_wp_with_no_assignees_is_zero(empty_wb):
    """With no assignments, every WP row monthly cell (active months) is =0.

    Inactive months remain empty/grey (not =0).
    """
    pf = empty_wb[SHEET_PLAN_FORECAST]
    wp_idx = 0  # WP1
    wp_id = WORK_PACKAGES[wp_idx]["id"]
    active = wp_active_months(wp_id, WP_SCHEDULE)
    row = _find_wp_row(pf, wp_idx)

    # Pick an active month — formula must be =0 (explicit, not SUM of nothing)
    month_num = min(active)
    col = MONTH_COL_START + (month_num - 1)
    assert pf.cell(row, col).value == "=0"


def test_plan_forecast_inactive_months_grey(pf):
    """WP4 starts M2 (demo schedule has WP4 at M2, WP2 is M1-M6).

    Pick a WP whose schedule start > 1 so there is an inactive month at M1.
    WP4 (idx=3) starts at M2 per WP_SCHEDULE.
    """
    wp_idx = 3  # WP4 — schedule starts M2
    wp_id = WORK_PACKAGES[wp_idx]["id"]
    assert wp_id == "WP4"
    active = wp_active_months(wp_id, WP_SCHEDULE)
    assert 1 not in active  # M1 is inactive for WP4

    row = _find_wp_row(pf, wp_idx)
    cell = pf.cell(row, MONTH_COL_START)  # M1 col
    assert cell.value in (None, "")
    fill = cell.fill
    assert fill is not None
    assert str(getattr(fill.fgColor, "rgb", "")).upper().endswith(
        str(GREY_FILL.fgColor.rgb).upper()
    ), f"Inactive month fill should be GREY_FILL, got {fill.fgColor.rgb!r}"


def test_plan_forecast_variance_red_cf(pf):
    """At least one CF rule with red fill applies on the Variance column."""
    vcols = val_columns(PROJECT_MONTHS)
    diff_col = col_letter(vcols["diff"])

    red_rules = 0
    for cf_range in pf.conditional_formatting:
        rng_str = str(cf_range)
        if diff_col not in rng_str:
            continue
        rules = pf.conditional_formatting[cf_range]
        for rule in rules:
            assert isinstance(rule, Rule)
            fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
            if fill is None:
                continue
            bg = getattr(fill.bgColor, "rgb", None)
            if bg and "C7CE" in str(bg).upper():  # CF_RED_FILL signature
                red_rules += 1
    assert red_rules >= 1, (
        f"Expected ≥1 red CF rule on variance col {diff_col}"
    )


def test_plan_forecast_total_row_sums_wp_rows(pf):
    """TOTAL row per-month formula sums the WP rows above it."""
    # Locate WP rows + total row by scanning col A.
    wp_rows = []
    for r in range(1, pf.max_row + 1):
        v = pf.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^=Config!A\d+$", v):
            wp_rows.append(r)
    assert len(wp_rows) == len(WORK_PACKAGES)
    first_wp, last_wp = wp_rows[0], wp_rows[-1]
    total_row = last_wp + 1

    # Month-1 column
    col = MONTH_COL_START
    formula = pf.cell(total_row, col).value
    assert isinstance(formula, str) and formula.startswith("=SUM(")
    # Should reference the WP block: col{first_wp}:col{last_wp}
    cl = col_letter(col)
    assert f"{cl}{first_wp}:{cl}{last_wp}" in formula, formula


# ── Validation columns ──

def test_plan_forecast_ceiling_column_refs_config(pf):
    """Ceiling column on each WP row cross-refs Config!E<wp_row>."""
    vcols = val_columns(PROJECT_MONTHS)
    ceil_col = vcols["budget"]  # 'budget' slot reused for Ceiling

    # Take first WP row
    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, ceil_col).value
    assert isinstance(val, str) and val.startswith("=Config!E"), val


def test_plan_forecast_grey_cf_past_hours_actuals_threshold(pf):
    """A cosmetic grey CF shades forecast-EUR cells whose month index is
    past ``hours_actuals_through_month`` — no hours yet, so the forecast
    aggregates zero and should read as empty/unreliable rather than zero."""
    target_grey = str(CF_GREY_FILL.bgColor.rgb).upper()
    # MONTH_COL_START = col C; the month band spans C..<end>.
    month_start_l = col_letter(MONTH_COL_START)

    found_formula = None
    for cf_range in pf.conditional_formatting:
        rng_str = str(cf_range)
        if month_start_l not in rng_str:
            continue
        for rule in pf.conditional_formatting[cf_range]:
            fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
            if fill is None:
                continue
            bg = str(getattr(fill.bgColor, "rgb", "") or "").upper()
            if not bg.endswith(target_grey):
                continue
            for f in (rule.formula or []):
                if "hours_actuals_through_month" in f:
                    found_formula = f
                    break

    assert found_formula is not None, (
        "Missing grey CF keyed on hours_actuals_through_month on the "
        "Plan-Forecast monthly range"
    )
    assert ">hours_actuals_through_month" in found_formula, found_formula


def test_plan_forecast_sum_column_totals_months(pf):
    """Sum column = SUM(<month range>) for each WP row."""
    vcols = val_columns(PROJECT_MONTHS)
    sum_col = vcols["sum"]
    end_col = month_col_end(PROJECT_MONTHS)

    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, sum_col).value
    assert isinstance(val, str) and val.startswith("=SUM("), val
    assert f"{col_letter(MONTH_COL_START)}{wp_row}:{col_letter(end_col)}{wp_row}" in val
