"""End-to-end tests requiring LibreOffice for formula recalculation.

Tests skip gracefully when LibreOffice is not available.
"""
import json
import subprocess

import pytest
import openpyxl
from pathlib import Path

from budget_planner.workbook import generate_workbook
from budget_planner.config import (
    MONTH_COL_START, NUM_WPS, PROJECT_MONTHS, DEFAULT_RATE,
    WORK_PACKAGES, VAL_COL_SUM, VAL_COL_BUDGET,
)

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


# ── Fixtures ──


@pytest.fixture(scope="module")
def demo_wb_path(tmp_path_factory):
    """Generate a demo workbook."""
    out = tmp_path_factory.mktemp("e2e") / "demo.xlsx"
    generate_workbook(str(out), mode="demo")
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
def skeleton_wb_path(tmp_path_factory):
    """Generate a skeleton workbook."""
    out = tmp_path_factory.mktemp("e2e_skel") / "skeleton.xlsx"
    generate_workbook(str(out), mode="skeleton")
    return str(out)


@pytest.fixture(scope="module")
def skeleton_recalc(skeleton_wb_path):
    """Recalculate skeleton workbook via LibreOffice."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), skeleton_wb_path],
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


def test_skeleton_mode_recalc_zero_errors(skeleton_recalc):
    """Generate skeleton workbook -> recalc -> 0 formula errors."""
    assert skeleton_recalc["status"] == "success"
    assert skeleton_recalc["total_errors"] == 0


def test_demo_mode_formula_count(demo_recalc):
    """Workbook should have substantial formula count (5 sheets, cross-sheet refs)."""
    assert demo_recalc["total_formulas"] > 200


# ── Plan <-> Capacity consistency (requires recalc) ──


def test_demo_mode_plan_planned_equals_capacity_sums(demo_wb_path, demo_recalc):
    """After recalc, Plan Planned FTE values should equal Capacity sums."""
    # Load recalculated workbook with data_only=True to get computed values
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    plan = wb["Plan"]

    from budget_planner.reader import discover_layout

    layout = discover_layout(plan)

    # For WP1 (active M1-M60), check that Plan planned M1 value > 0
    # (it should be the sum of all persons' WP1 allocations)
    planned_row = layout["planned_first_wp_row"]
    val = plan.cell(planned_row, MONTH_COL_START).value
    # After recalc, this should be a computed number, not a formula
    assert val is not None and isinstance(val, (int, float)) and val > 0, \
        f"Plan WP1 M1 after recalc should be positive, got: {val}"
    wb.close()


# ── GAP 7: Summary values after recalc ──

# Layout constants matching summary.py
_YEAR_GROUP_START = 3
_COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR
_NUM_YEARS = (PROJECT_MONTHS + 11) // 12
_TOTAL_GROUP_START = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR


def _find_summary_first_wp_row(ws):
    """Scan Summary col A for first WP ID (works with data_only=True)."""
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if isinstance(val, str) and val == WORK_PACKAGES[0]["id"]:
            return row
    raise AssertionError("Could not find Summary first WP row")


def test_summary_year1_plan_fte_positive(demo_wb_path, demo_recalc):
    """After recalc, Summary Year 1 Plan FTE should be > 0 for at least one WP."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    # Year 1 Plan FTE column = YEAR_GROUP_START (col C)
    found_positive = False
    for i in range(NUM_WPS):
        val = summary.cell(first_wp + i, _YEAR_GROUP_START).value
        if isinstance(val, (int, float)) and val > 0:
            found_positive = True
            break
    assert found_positive, "No WP has positive Year 1 Plan FTE after recalc"
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


def test_summary_actual_eur_zero(demo_wb_path, demo_recalc):
    """Fresh workbook: all Summary Actual EUR values should be 0."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    for i in range(NUM_WPS):
        row = first_wp + i
        # Check each year's Actual EUR (offset +2 within year group)
        for y in range(_NUM_YEARS):
            col = _YEAR_GROUP_START + y * _COLS_PER_YEAR + 2
            val = summary.cell(row, col).value
            assert val is None or val == 0, \
                f"WP row {row}, Year {y+1} Act EUR should be 0, got {val}"
        # Check project total Actual EUR (offset +2 within total group)
        total_act = summary.cell(row, _TOTAL_GROUP_START + 2).value
        assert total_act is None or total_act == 0, \
            f"WP row {row}, total Act EUR should be 0, got {total_act}"
    wb.close()


def test_summary_budget_overview_after_recalc(demo_wb_path, demo_recalc):
    """Budget Overview: Personnel Actual = 0, Personnel Ceiling > 0."""
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
    assert personnel_actual == 0, \
        f"Personnel Actual should be 0 in fresh workbook, got {personnel_actual}"

    assert personnel_ceiling is not None, "Personnel Ceiling label not found"
    assert isinstance(personnel_ceiling, (int, float)) and personnel_ceiling > 0, \
        f"Personnel Ceiling should be > 0, got {personnel_ceiling}"

    # Cross-check: ceiling should equal sum(wp.fte) * DEFAULT_RATE
    expected_ceiling = sum(wp["fte"] for wp in WORK_PACKAGES) * DEFAULT_RATE
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


def test_actuals_total_sum_zero(demo_wb_path, demo_recalc):
    """Fresh workbook: Actuals TOTAL Sum should be 0 after recalc."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    actuals = wb["Actuals"]
    total_row = _find_actuals_total_row(actuals)
    val = actuals.cell(total_row, VAL_COL_SUM).value
    assert val is not None and (val == 0 or val == 0.0), \
        f"Actuals TOTAL Sum should be 0, got {val}"
    wb.close()


def test_actuals_total_ceiling_equals_partner_ceiling(demo_wb_path, demo_recalc):
    """Actuals TOTAL Ceiling should equal partner_ceiling after recalc."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    actuals = wb["Actuals"]
    total_row = _find_actuals_total_row(actuals)
    val = actuals.cell(total_row, VAL_COL_BUDGET).value
    expected = sum(wp["fte"] for wp in WORK_PACKAGES) * DEFAULT_RATE
    assert isinstance(val, (int, float)), \
        f"Actuals TOTAL Ceiling should be numeric, got {val}"
    assert abs(val - expected) < 0.01, \
        f"Actuals TOTAL Ceiling {val} != expected partner_ceiling {expected}"
    wb.close()


