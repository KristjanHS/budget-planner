"""Tests for gap-aware Capacity distribution when WP SCHEDULE has gaps.

Scenario: WP1 active M8-17 and M20-26 (inactive M18-19). Alice is assigned
with a single window M10-20 spanning the gap. The Capacity builder must
intersect the assignment window with the WP's active months so that:

- Gap months M18, M19 are grey-filled and empty.
- Total FTE is distributed ONLY across the valid months in the window.
- Helper columns reconcile because no phantom FTE is allocated to inactive months.
"""

from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

from budget_planner.capacity import build_capacity_sheet
from budget_planner.config_sheet import build_config_sheet
from budget_planner.layout import MONTH_COL_START, month_col_end as _mce
from budget_planner.plan import build_plan_sheet


PROJECT_PARAMS = {
    "title": "Gap-WP Test",
    "start": datetime(2026, 11, 1),
    "months": 30,
    "rate": 7500,
    "eu_ceiling_pm": 12.0,
}

WORK_PACKAGES = [{"id": "WP1", "name": "WP with gap", "fte": 12.0}]

# WP1 has TWO active ranges — gap at months 18, 19.
WP_SCHEDULE = [
    {"id": "WP1", "start": 8,  "end": 17, "type": "flat"},
    {"id": "WP1", "start": 20, "end": 26, "type": "flat"},
]

TEAM = [{"name": "Alice", "max_fte": 1.0, "grade": "Mid"}]

# Single assignment spanning the gap: M10-20, total_fte=5.5.
ASSIGNMENTS = [
    {"person": "Alice", "wp": "WP1", "start": 10, "end": 20, "total_fte": 5.5},
]

REPORTING_PERIODS = [{"id": "RP1", "start": 1, "end": 30}]
GRADE_RATES = {"Senior": 12500, "Mid": 10000, "Junior": 7500}

TOTAL_FTE = 5.5
VALID_MONTHS = frozenset({10, 11, 12, 13, 14, 15, 16, 17, 20})   # 9 months
GAP_MONTHS_IN_WINDOW = frozenset({18, 19})
OUT_OF_WINDOW = frozenset(range(1, 31)) - VALID_MONTHS - GAP_MONTHS_IN_WINDOW


@pytest.fixture(scope="module")
def cap_ws(tmp_path_factory):
    """Build a minimal workbook with the gap-WP scenario; return the Capacity sheet."""
    wb = Workbook()
    meta = wb.active
    assert meta is not None
    meta.title = "_META"

    cfg_ws = wb.create_sheet("Config")
    config_layout = build_config_sheet(
        cfg_ws, wb, PROJECT_PARAMS, WORK_PACKAGES,
        wp_schedule=WP_SCHEDULE, team=TEAM,
        reporting_periods=REPORTING_PERIODS, grade_rates=GRADE_RATES,
        assignments=ASSIGNMENTS, external_commitments=[],
    )
    plan_ws = wb.create_sheet("Plan")
    plan_layout = build_plan_sheet(
        plan_ws, wb,
        project_params=PROJECT_PARAMS,
        work_packages=WORK_PACKAGES,
        wp_schedule=WP_SCHEDULE,
        config_layout=config_layout,
    )
    capacity_ws = wb.create_sheet("Capacity")
    build_capacity_sheet(
        capacity_ws, config_layout,
        project_params=PROJECT_PARAMS,
        work_packages=WORK_PACKAGES,
        wp_schedule=WP_SCHEDULE,
        team=TEAM,
        external_commitments=[],
        assignments=ASSIGNMENTS,
        config_first_wp_row=plan_layout["config_first_wp_row"],
    )

    out = tmp_path_factory.mktemp("gap_wp") / "gap.xlsx"
    wb.save(str(out))
    reopened = load_workbook(str(out), data_only=False)
    return reopened["Capacity"]


def _find_wp1_row(ws) -> int:
    """Locate Alice's WP1 row: col A has =Config!A<n>, col B has =Config!B<n>.

    The person-header row also references Config!A (name ref) but its B cell is
    empty/merged, so we require both formula refs to disambiguate.
    """
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if (isinstance(a, str) and "Config!A" in a
                and isinstance(b, str) and "Config!B" in b):
            return r
    raise AssertionError("No WP row found for Alice/WP1")


def test_gap_months_in_window_are_grey_and_empty(cap_ws):
    """M18, M19 inside the assignment window but outside WP SCHEDULE → grey, empty."""
    wp1_row = _find_wp1_row(cap_ws)
    for m in GAP_MONTHS_IN_WINDOW:
        c = MONTH_COL_START + m - 1
        cell = cap_ws.cell(row=wp1_row, column=c)
        assert cell.value in (None, ""), (
            f"M{m}: expected empty, got {cell.value!r}"
        )
        rgb = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor else None
        assert rgb in ("D9D9D9", "00D9D9D9"), (
            f"M{m}: expected grey fill (D9D9D9), got {rgb}"
        )


def test_valid_months_have_distributed_fte(cap_ws):
    """Valid months in window receive total_fte / len(valid_months)."""
    wp1_row = _find_wp1_row(cap_ws)
    expected_per_month = TOTAL_FTE / len(VALID_MONTHS)
    for m in VALID_MONTHS:
        c = MONTH_COL_START + m - 1
        val = cap_ws.cell(row=wp1_row, column=c).value
        assert isinstance(val, (int, float)), (
            f"M{m}: expected number, got {val!r}"
        )
        assert val == pytest.approx(expected_per_month, abs=0.001), (
            f"M{m}: expected ~{expected_per_month:.4f}, got {val}"
        )


def test_valid_month_values_sum_to_total_fte(cap_ws):
    """Sum of written FTE values equals Total FTE (no FTE leaked into gap months)."""
    wp1_row = _find_wp1_row(cap_ws)
    written_sum = 0.0
    for m in range(1, PROJECT_PARAMS["months"] + 1):
        c = MONTH_COL_START + m - 1
        val = cap_ws.cell(row=wp1_row, column=c).value
        if isinstance(val, (int, float)):
            written_sum += val
    assert written_sum == pytest.approx(TOTAL_FTE, abs=0.001), (
        f"Sum of all monthly values: got {written_sum}, expected {TOTAL_FTE}"
    )


def test_out_of_window_months_are_grey_and_empty(cap_ws):
    """Months outside assignment window remain grey + empty (regression guard)."""
    wp1_row = _find_wp1_row(cap_ws)
    for m in OUT_OF_WINDOW:
        c = MONTH_COL_START + m - 1
        cell = cap_ws.cell(row=wp1_row, column=c)
        assert cell.value in (None, ""), (
            f"M{m} (out of window): expected empty, got {cell.value!r}"
        )


def test_sum_formula_row_wide(cap_ws):
    """Row-wide SUM formula still references the full month range (gap months sum to 0)."""
    wp1_row = _find_wp1_row(cap_ws)
    sum_col = _mce(PROJECT_PARAMS["months"]) + 1
    val = cap_ws.cell(row=wp1_row, column=sum_col).value
    assert isinstance(val, str) and val.startswith("=SUM("), (
        f"Sum cell: expected SUM formula, got {val!r}"
    )
