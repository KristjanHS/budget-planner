"""Tests for the Actuals sheet builder."""
import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.config import (
    WORK_PACKAGES, NUM_WPS, MONTH_COL_START,
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
    tc = ws.sheet_properties.tabColor
    assert tc is not None
    assert tc.rgb == "000000FF"


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


def test_actuals_values_are_zero(actuals_workbook):
    """Fresh workbook has 0 in active month cells."""
    ws = actuals_workbook["Actuals"]
    # Find first WP row
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    wp = WORK_PACKAGES[0]
    row = first_wp_row
    c = MONTH_COL_START + wp["start"] - 1
    assert ws.cell(row, c).value == 0


def test_actuals_inactive_months_empty(actuals_workbook):
    """Inactive month cells are empty with grey fill."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    # WP3 starts at month 7 -- months before start should be inactive
    wp3_idx = next(i for i, w in enumerate(WORK_PACKAGES) if w["id"] == "WP3")
    wp3_start = WORK_PACKAGES[wp3_idx]["start"]
    row = first_wp_row + wp3_idx
    for m in range(wp3_start - 1):
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
    assert cell.value and "Plan!G" in str(cell.value)


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
