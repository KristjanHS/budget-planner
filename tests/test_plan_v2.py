"""Plan v2 tests — planned FTE cells are cross-sheet SUM formulas from Capacity."""
import re

import openpyxl
import pytest

from budget_planner.workbook import init_workbook
from budget_planner.reader import discover_layout
from budget_planner.config import (
    WORK_PACKAGES, MONTH_COL_START, PROJECT_MONTHS,
)


@pytest.fixture(scope="module")
def v2_workbook(tmp_path_factory):
    """Generate a workbook with demo assignments for Plan formula testing."""
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
    out = tmp_path_factory.mktemp("plan_v2") / "test.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return openpyxl.load_workbook(str(out), data_only=False)


@pytest.fixture(scope="module")
def v2_layout(v2_workbook):
    return discover_layout(v2_workbook["Plan"])


# ── Planned FTE cells reference Capacity ──

@pytest.mark.parametrize("wp_idx", [0, 2])
def test_planned_fte_formula_refs_match_capacity_layout(
    v2_workbook, v2_layout, wp_idx
):
    """SUM formula row references match the capacity_layout person_groups."""
    plan = v2_workbook["Plan"]
    capacity = v2_workbook["Capacity"]
    row = v2_layout["planned_first_wp_row"] + wp_idx
    wp = WORK_PACKAGES[wp_idx]

    # Find an active month to inspect
    active_month = None
    for m in range(PROJECT_MONTHS):
        month_num = m + 1
        if wp["start"] <= month_num <= wp["end"]:
            active_month = m
            break
    assert active_month is not None

    formula = plan.cell(row, MONTH_COL_START + active_month).value
    assert isinstance(formula, str) and formula.startswith("=SUM(")

    # Extract row numbers from Capacity!XX<row> references
    refs = re.findall(r"Capacity![A-Z]+(\d+)", formula)
    formula_rows = sorted(int(r) for r in refs)
    assert len(formula_rows) > 0, f"No Capacity row refs found in {formula}"

    # Verify each referenced row in Capacity has this WP's ID in column A
    config = v2_workbook["Config"]
    for cap_row in formula_rows:
        cap_val = capacity.cell(cap_row, 1).value
        # Column A is a formula referencing Config (=Config!A<n>),
        # so we follow the reference to get the WP ID
        if isinstance(cap_val, str) and cap_val.startswith("="):
            # Parse =Config!A<row> to get the config row
            match = re.match(r"=Config!A(\d+)", cap_val)
            assert match, f"Unexpected formula in Capacity!A{cap_row}: {cap_val}"
            config_row = int(match.group(1))
            wp_id_from_config = config.cell(config_row, 1).value
            assert wp_id_from_config == wp["id"], (
                f"Capacity row {cap_row} refs config row {config_row} "
                f"with WP ID '{wp_id_from_config}', expected '{wp['id']}'"
            )
        else:
            # Direct value (shouldn't happen for programme WPs, but handle gracefully)
            assert cap_val == wp["id"], (
                f"Capacity!A{cap_row} = '{cap_val}', expected '{wp['id']}'"
            )


@pytest.mark.parametrize("wp_idx", [0, 2])
def test_planned_fte_cells_have_black_font(v2_workbook, v2_layout, wp_idx):
    """Monthly cells in planned block have black font (formula, not input)."""
    plan = v2_workbook["Plan"]
    row = v2_layout["planned_first_wp_row"] + wp_idx
    wp = WORK_PACKAGES[wp_idx]

    for m in range(PROJECT_MONTHS):
        month_num = m + 1
        if wp["start"] <= month_num <= wp["end"]:
            cell = plan.cell(row, MONTH_COL_START + m)
            color = cell.font.color
            # Black font: color value is "000000" or theme 1
            assert color is not None, (
                f"WP {wp['id']} M{month_num}: font color is None"
            )
            if color.rgb and color.rgb != "00000000":
                assert color.rgb in ("FF000000", "00000000", "000000"), (
                    f"WP {wp['id']} M{month_num}: expected black font, "
                    f"got {color.rgb}"
                )


