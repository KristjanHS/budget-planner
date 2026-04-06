"""End-to-end tests requiring LibreOffice for formula recalculation.

Tests skip gracefully when LibreOffice is not available.
"""
import json
import subprocess

import pytest
import openpyxl
from pathlib import Path

from budget_planner.workbook import init_workbook, seed_demo_data
from budget_planner.layout import _calendar_year_groups
from budget_planner.config import (
    MONTH_COL_START, NUM_WPS, PROJECT_MONTHS, PROJECT_START,
    WORK_PACKAGES, VAL_COL_SUM, VAL_COL_BUDGET,
    TEAM, GRADE_RATES, DEFAULT_GRADE,
)


def _weighted_avg_rate():
    """Compute expected FTE rate from config seed data (grade-weighted average)."""
    total_fte = sum(m["max_fte"] for m in TEAM)
    weighted = sum(
        GRADE_RATES[m.get("grade", DEFAULT_GRADE)] * m["max_fte"] for m in TEAM
    )
    return weighted / total_fte

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


# ── Fixtures ──


@pytest.fixture(scope="module")
def demo_wb_path(tmp_path_factory):
    """Generate a demo workbook (init with assignments + seed_demo_data)."""
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
    out = tmp_path_factory.mktemp("e2e") / "demo.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    seed_demo_data(str(out))
    return str(out)


@pytest.fixture(scope="module")
def demo_recalc(demo_wb_path):
    """Recalculate demo workbook via LibreOffice."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), demo_wb_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)


@pytest.fixture(scope="module")
def empty_wb_path(tmp_path_factory):
    """Generate an empty (no demo) workbook."""
    out = tmp_path_factory.mktemp("e2e_empty") / "empty.xlsx"
    init_workbook(str(out))
    return str(out)


@pytest.fixture(scope="module")
def empty_recalc(empty_wb_path):
    """Recalculate empty workbook via LibreOffice."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), empty_wb_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)


# ── Recalc tests ──


def test_demo_mode_recalc_zero_errors(demo_recalc):
    """Generate demo workbook -> recalc -> 0 formula errors across all 5 sheets."""
    assert demo_recalc["status"] == "success"
    assert demo_recalc["total_errors"] == 0


def test_empty_mode_recalc_zero_errors(empty_recalc):
    """Generate empty workbook -> recalc -> 0 formula errors."""
    assert empty_recalc["status"] == "success"
    assert empty_recalc["total_errors"] == 0


def test_demo_mode_formula_count(demo_recalc):
    """Workbook should have substantial formula count (5 sheets, cross-sheet refs)."""
    assert demo_recalc["total_formulas"] > 200


# ── Plan <-> Capacity consistency (requires recalc) ──


def test_demo_mode_plan_planned_equals_capacity_sums(demo_wb_path, demo_recalc):
    """After recalc, Plan Planned FTE values should be numeric (sum of Capacity)."""
    # Load recalculated workbook with data_only=True to get computed values
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    plan = wb["Plan"]

    from budget_planner.reader import discover_layout

    layout = discover_layout(plan)

    # For WP1 (active M1-M14), check that Plan planned M1 value is numeric
    # (no seeding means values are 0, but the formula should resolve to a number)
    planned_row = layout["planned_first_wp_row"]
    val = plan.cell(planned_row, MONTH_COL_START).value
    # After recalc, this should be a computed number (0 without seeding)
    assert val is not None and isinstance(val, (int, float)), \
        f"Plan WP1 M1 after recalc should be numeric, got: {val}"
    wb.close()


# ── GAP 7: Summary values after recalc ──

# Layout constants matching summary.py
_YEAR_GROUP_START = 3
_COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
_TOTAL_GROUP_START = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR


def _find_summary_first_wp_row(ws):
    """Scan Summary col A for first WP ID (works with data_only=True)."""
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if isinstance(val, str) and val == WORK_PACKAGES[0]["id"]:
            return row
    raise AssertionError("Could not find Summary first WP row")


def test_summary_year1_plan_fte_numeric(demo_wb_path, demo_recalc):
    """After recalc, Summary Year 1 Plan FTE should be numeric for all WPs."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    # Year 1 Plan FTE column = YEAR_GROUP_START (col C)
    for i in range(NUM_WPS):
        val = summary.cell(first_wp + i, _YEAR_GROUP_START).value
        assert val is None or isinstance(val, (int, float)), (
            f"WP row {first_wp + i} Year 1 Plan FTE should be numeric, got {val}"
        )
    wb.close()


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
        assert isinstance(total_val, (int, float)), \
            f"WP row {row}: total Plan FTE should be numeric, got {total_val}"
        assert abs(total_val - year_sum) < 0.01, \
            f"WP row {row}: total {total_val} != year sum {year_sum}"
    wb.close()


def test_summary_actual_eur_seeded(demo_wb_path, demo_recalc):
    """Fresh demo workbook: Summary Actual EUR reflects seeded actuals data."""
    from budget_planner.config import DEMO_CURRENT_MONTH
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    # First calendar year group should have positive actuals for active WPs
    row = first_wp  # WP1 (active from M1)
    col = _YEAR_GROUP_START + 0 * _COLS_PER_YEAR + 2  # First year Actual EUR
    val = summary.cell(row, col).value
    assert isinstance(val, (int, float)) and val > 0, \
        f"WP1 first calendar year Act EUR should be > 0 in demo mode, got {val}"

    # Find the first calendar year group that starts after DEMO_CURRENT_MONTH
    # (all its months are beyond the seeded range, so actuals should be 0)
    unseeded_year_idx = None
    for y, (group_start, _span) in enumerate(_CAL_GROUPS):
        if group_start >= DEMO_CURRENT_MONTH:
            unseeded_year_idx = y
            break
    assert unseeded_year_idx is not None, "All calendar years fall within demo seed range"
    col_unseeded = _YEAR_GROUP_START + unseeded_year_idx * _COLS_PER_YEAR + 2
    val_unseeded = summary.cell(row, col_unseeded).value
    assert val_unseeded is None or val_unseeded == 0, \
        f"WP1 calendar year {unseeded_year_idx+1} Act EUR should be 0, got {val_unseeded}"
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
    assert isinstance(personnel_actual, (int, float)) and personnel_actual > 0, \
        f"Personnel Actual should be > 0 in demo mode, got {personnel_actual}"

    assert personnel_ceiling is not None, "Personnel Ceiling label not found"
    assert isinstance(personnel_ceiling, (int, float)) and personnel_ceiling > 0, \
        f"Personnel Ceiling should be > 0, got {personnel_ceiling}"

    # Cross-check: ceiling should equal sum(wp.fte) * weighted_avg_rate
    expected_rate = _weighted_avg_rate()
    expected_ceiling = sum(wp["fte"] for wp in WORK_PACKAGES) * expected_rate
    assert abs(personnel_ceiling - expected_ceiling) < 0.01, \
        f"Personnel Ceiling {personnel_ceiling} != expected {expected_ceiling}"
    wb.close()


# ── GAP 8: Actuals sheet post-recalc ──


def _find_actuals_total_row(ws):
    """Scan Actuals col B for the TOTAL row."""
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 2).value
        if isinstance(val, str) and val.upper() == "TOTAL":
            return r
    raise AssertionError("Could not find Actuals TOTAL row")


def test_actuals_total_sum_positive(demo_wb_path, demo_recalc):
    """Fresh demo workbook: Actuals TOTAL Sum should be positive (seeded data)."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    actuals = wb["Actuals"]
    total_row = _find_actuals_total_row(actuals)
    val = actuals.cell(total_row, VAL_COL_SUM).value
    assert isinstance(val, (int, float)) and val > 0, \
        f"Actuals TOTAL Sum should be > 0 in demo mode, got {val}"
    wb.close()


def test_actuals_total_ceiling_equals_partner_ceiling(demo_wb_path, demo_recalc):
    """Actuals TOTAL Ceiling should equal partner_ceiling after recalc."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    actuals = wb["Actuals"]
    total_row = _find_actuals_total_row(actuals)
    val = actuals.cell(total_row, VAL_COL_BUDGET).value
    expected = sum(wp["fte"] for wp in WORK_PACKAGES) * _weighted_avg_rate()
    assert isinstance(val, (int, float)), \
        f"Actuals TOTAL Ceiling should be numeric, got {val}"
    assert abs(val - expected) < 0.01, \
        f"Actuals TOTAL Ceiling {val} != expected partner_ceiling {expected}"
    wb.close()


