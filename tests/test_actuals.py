"""Tests for the Actuals sheet builder."""
import pytest
import openpyxl

from budget_planner.config import (
    WORK_PACKAGES, WP_SCHEDULE, NUM_WPS, MONTH_COL_START,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
    VAL_COL_CUM, VAL_COL_PCT_CEIL, VAL_COL_BREACH,
)
from budget_planner.layout import wp_month_bounds
from budget_planner.styles import col_letter


@pytest.fixture(scope="module")
def actuals_workbook(demo_wb_path):
    return openpyxl.load_workbook(demo_wb_path, data_only=False)


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
    # Find first WP row (row after month headers -- look for "WP ID" in col A)
    month_header_row = None
    for r in range(1, total_row):
        if ws.cell(r, 1).value == "WP ID":
            month_header_row = r
            break
    assert month_header_row is not None
    first_wp_row = month_header_row + 1
    wp_count = total_row - first_wp_row
    assert wp_count == NUM_WPS


def test_actuals_demo_values_seeded(actuals_workbook):
    """Fresh demo workbook has non-zero EUR in active months up to DEMO_CURRENT_MONTH."""
    from budget_planner.config import DEMO_CURRENT_MONTH
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    wp = WORK_PACKAGES[0]  # WP1: active from M1, multiplier 1.0
    wp_start, wp_end = wp_month_bounds(wp["id"], WP_SCHEDULE)
    row = first_wp_row
    # Active month within demo range should have a positive value
    c = MONTH_COL_START + wp_start - 1
    assert ws.cell(row, c).value > 0
    # Month after DEMO_CURRENT_MONTH should be empty (not seeded)
    if DEMO_CURRENT_MONTH < wp_end:
        c_after = MONTH_COL_START + DEMO_CURRENT_MONTH
        val = ws.cell(row, c_after).value
        assert val is None or val == 0, f"Expected empty/0 after demo range, got {val}"


def test_actuals_inactive_months_empty(actuals_workbook):
    """Inactive month cells are empty with grey fill."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    # Find any WP with start > 1 -- months before start should be inactive
    late_idx = next(
        i for i, w in enumerate(WORK_PACKAGES)
        if wp_month_bounds(w["id"], WP_SCHEDULE)[0] > 1
    )
    late_start = wp_month_bounds(WORK_PACKAGES[late_idx]["id"], WP_SCHEDULE)[0]
    row = first_wp_row + late_idx
    for m in range(late_start - 1):
        c = MONTH_COL_START + m
        cell = ws.cell(row, c)
        assert cell.value is None or cell.value == "", f"M{m+1} should be inactive, got {cell.value}"
        assert cell.fill.fgColor and cell.fill.fgColor.rgb and "D9D9D9" in str(cell.fill.fgColor.rgb), \
            f"M{m+1} should have grey fill"


def test_actuals_sum_formula(actuals_workbook):
    """Sum validation column has SUM formula."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_SUM)
    assert cell.value and str(cell.value).startswith("=SUM(")


def test_actuals_ceiling_refs_plan(actuals_workbook):
    """Ceiling validation column references Plan config table."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_BUDGET)
    assert cell.value and "Config!E" in str(cell.value)


def test_actuals_variance_formula(actuals_workbook):
    """Variance = Ceiling - Sum."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
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


# ── Cumulative columns ──


def _find_first_wp_row(ws):
    """Find the first WP data row (row after 'WP ID' header)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r + 1
    return None


def _find_total_row(ws):
    """Find the TOTAL row."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            return r
    return None


def _find_year_row(ws):
    """Find the year header row (row before the month header row)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r - 1
    return None


def test_actuals_cum_eur_formula(actuals_workbook):
    """First WP row at VAL_COL_CUM has SUMPRODUCT formula with current_month."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_CUM)
    formula = str(cell.value)
    assert "SUMPRODUCT" in formula
    assert "current_month" in formula


def test_actuals_pct_ceiling_formula(actuals_workbook):
    """First WP row at VAL_COL_PCT_CEIL divides cum by ceiling."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_PCT_CEIL)
    formula = str(cell.value)
    cum_l = col_letter(VAL_COL_CUM)
    ceil_l = col_letter(VAL_COL_BUDGET)
    assert cum_l in formula and ceil_l in formula


def test_actuals_breach_formula(actuals_workbook):
    """First WP row at VAL_COL_BREACH has IF formula with cum-to-date slope.

    Denominator was migrated from ``current_month`` to ``months_elapsed_eur``
    in T7 (Stage 2) so the slope reflects closed finance months rather than
    the raw project-month index. See ``test_actuals_breach_todate.py`` for
    the full invariant.
    """
    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_BREACH)
    formula = str(cell.value)
    assert "IF" in formula
    assert "months_elapsed_eur" in formula


def test_actuals_cum_headers(actuals_workbook):
    """Year row has Cum. EUR, % Ceiling, Breach In headers."""
    ws = actuals_workbook["Actuals"]
    year_row = _find_year_row(ws)
    assert year_row is not None
    assert ws.cell(year_row, VAL_COL_CUM).value == "Cum. EUR"
    assert ws.cell(year_row, VAL_COL_PCT_CEIL).value == "% Ceiling"
    assert ws.cell(year_row, VAL_COL_BREACH).value == "Breach In"


def test_actuals_total_cum_formula(actuals_workbook):
    """TOTAL row at VAL_COL_CUM has SUMPRODUCT formula."""
    ws = actuals_workbook["Actuals"]
    total_row = _find_total_row(ws)
    assert total_row is not None
    cell = ws.cell(total_row, VAL_COL_CUM)
    formula = str(cell.value)
    assert "SUMPRODUCT" in formula
    assert "current_month" in formula


def test_actuals_total_pct_ceiling_uses_partner_ceiling(actuals_workbook):
    """TOTAL row at VAL_COL_PCT_CEIL references partner_ceiling."""
    ws = actuals_workbook["Actuals"]
    total_row = _find_total_row(ws)
    assert total_row is not None
    cell = ws.cell(total_row, VAL_COL_PCT_CEIL)
    formula = str(cell.value)
    assert "partner_ceiling" in formula
