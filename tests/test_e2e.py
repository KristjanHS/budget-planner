"""End-to-end tests requiring LibreOffice for formula recalculation.

Tests skip gracefully when LibreOffice is not available.

Shared fixtures (`demo_wb_path`, `demo_recalc`) live in `conftest.py` at
session scope so other modules can consume the same recalc'd workbook.
"""

import openpyxl

from budget_planner.layout import MONTH_COL_START, _calendar_year_groups
from budget_planner.config import (
    NUM_WPS,
    PROJECT_MONTHS,
    PROJECT_START,
    WORK_PACKAGES,
    VAL_COL_BUDGET,
    DEFAULT_RATE,
)


# ── Recalc tests ──


def test_demo_mode_recalc_zero_errors(demo_recalc):
    """Generate demo workbook -> recalc -> 0 formula errors across all 5 sheets."""
    assert demo_recalc["status"] == "success"
    assert demo_recalc["total_errors"] == 0


# ── GAP 7: Summary values after recalc ──

# Layout constants matching summary.py
_YEAR_GROUP_START = 3
# Plan FTE, Plan EUR, Act EUR, Var EUR — per-year Forecast EUR retired in
# the posted-cumulative redesign; project-end Forecast EUR sits in its own
# column between the year band and PROJECT TOTAL.
_COLS_PER_YEAR = 4
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
_FORECAST_COL = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR
_TOTAL_GROUP_START = _FORECAST_COL + 1


def _find_summary_first_wp_row(ws):
    """Scan Summary col A for first WP ID (works with data_only=True)."""
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if isinstance(val, str) and val == WORK_PACKAGES[0]["id"]:
            return row
    raise AssertionError("Could not find Summary first WP row")


def test_summary_total_plan_fte_equals_year_sum(demo_wb_path, demo_recalc):
    """Project total Plan FTE should equal sum of yearly Plan FTE values."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    for i in range(NUM_WPS):
        row = first_wp + i
        year_sum = 0.0
        for y in range(_NUM_YEARS):
            val = summary.cell(row, _YEAR_GROUP_START + y * _COLS_PER_YEAR).value
            if isinstance(val, (int, float)):
                year_sum += val
        total_val = summary.cell(row, _TOTAL_GROUP_START).value
        assert isinstance(total_val, (int, float)), f"WP row {row}: total Plan FTE should be numeric, got {total_val}"
        assert abs(total_val - year_sum) < 0.01, f"WP row {row}: total {total_val} != year sum {year_sum}"
    wb.close()


def test_summary_actual_eur_seeded(demo_wb_path, demo_recalc):
    """Fresh demo workbook: Summary Actual EUR reflects seeded actuals data."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    # First calendar year group should have positive actuals for active WPs
    row = first_wp  # WP1 (active from M1)
    col = _YEAR_GROUP_START + 0 * _COLS_PER_YEAR + 2  # First year Actual EUR
    val = summary.cell(row, col).value
    assert isinstance(val, (int, float)) and val > 0, (
        f"WP1 first calendar year Act EUR should be > 0 in demo mode, got {val}"
    )

    # Each calendar-year Act EUR sums exactly that year's WP1 cells on Actuals (unseeded years → 0)
    actuals = wb["Actuals"]
    act_row = next(r for r in range(1, actuals.max_row + 1) if actuals.cell(r, 1).value == WORK_PACKAGES[0]["id"])
    for y, (group_start, span) in enumerate(_CAL_GROUPS):
        cells = [actuals.cell(act_row, MONTH_COL_START + m).value for m in range(group_start, group_start + span)]
        expected = sum(v for v in cells if isinstance(v, (int, float)))
        got = summary.cell(row, _YEAR_GROUP_START + y * _COLS_PER_YEAR + 2).value or 0
        assert got == expected, f"WP1 calendar year {y + 1} Act EUR {got} != Actuals sum {expected}"
    wb.close()


def test_summary_budget_overview_after_recalc(demo_wb_path, demo_recalc):
    """Budget Overview: Personnel Actual > 0 (seeded), Personnel Ceiling > 0."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]

    # Scan col A for Budget Overview labels
    personnel_actual = None
    personnel_ceiling = None
    for row in range(1, summary.max_row + 1):
        label = summary.cell(row, 1).value
        if label == "Personnel Actual":
            personnel_actual = summary.cell(row, 2).value
        elif label == "Personnel Ceiling":
            personnel_ceiling = summary.cell(row, 2).value

    assert personnel_actual is not None, "Personnel Actual label not found"
    assert isinstance(personnel_actual, (int, float)) and personnel_actual > 0, (
        f"Personnel Actual should be > 0 in demo mode, got {personnel_actual}"
    )

    assert personnel_ceiling is not None, "Personnel Ceiling label not found"
    assert isinstance(personnel_ceiling, (int, float)) and personnel_ceiling > 0, (
        f"Personnel Ceiling should be > 0, got {personnel_ceiling}"
    )

    # Cross-check: ceiling = sum(FTE * per-WP rate), all WPs default to DEFAULT_RATE
    expected_ceiling = sum(wp["fte"] for wp in WORK_PACKAGES) * DEFAULT_RATE
    assert abs(personnel_ceiling - expected_ceiling) < 0.01, (
        f"Personnel Ceiling {personnel_ceiling} != expected {expected_ceiling}"
    )
    wb.close()


# ── GAP 8: Actuals sheet post-recalc ──


def _find_actuals_total_row(ws):
    """Scan Actuals col B for the TOTAL row."""
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if isinstance(val, str) and val.upper() == "TOTAL":
            return r
    raise AssertionError("Could not find Actuals TOTAL row")


def test_actuals_total_ceiling_equals_partner_ceiling(demo_wb_path, demo_recalc):
    """Actuals TOTAL Ceiling should equal partner_ceiling after recalc."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    actuals = wb["Actuals"]
    total_row = _find_actuals_total_row(actuals)
    val = actuals.cell(total_row, VAL_COL_BUDGET).value
    expected = sum(wp["fte"] for wp in WORK_PACKAGES) * DEFAULT_RATE
    assert isinstance(val, (int, float)), f"Actuals TOTAL Ceiling should be numeric, got {val}"
    assert abs(val - expected) < 0.01, f"Actuals TOTAL Ceiling {val} != expected partner_ceiling {expected}"
    wb.close()
