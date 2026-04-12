"""Plan v2 tests — planned FTE cells are cross-sheet SUM formulas from Capacity."""
import re

import openpyxl
import pytest
from openpyxl import Workbook

from budget_planner.workbook import init_workbook
from budget_planner.reader import discover_layout
from budget_planner import config as _config
from budget_planner.config import (
    WORK_PACKAGES, WP_SCHEDULE, MONTH_COL_START, PROJECT_MONTHS,
)
from budget_planner.config_sheet import build_config_sheet
from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.capacity import build_capacity_sheet
from budget_planner.layout import wp_month_bounds


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
    wp_start, wp_end = wp_month_bounds(wp["id"], WP_SCHEDULE)

    # Find an active month to inspect
    active_month = None
    for m in range(PROJECT_MONTHS):
        month_num = m + 1
        if wp_start <= month_num <= wp_end:
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


def _build_plan_standalone(*, project_params, work_packages, wp_schedule):
    """Build Config + Capacity + Plan in a fresh wb, returning the Plan layout + formulas.

    Used to compare None-defaulted invocations against explicit invocations.
    """
    wb = Workbook()
    cfg_ws = wb.active
    assert cfg_ws is not None
    cfg_ws.title = "Config"
    # build_config_sheet always needs explicit params (they're positional)
    cfg_params = project_params if project_params is not None else _config.default_project_params()
    cfg_wps = work_packages if work_packages is not None else _config.WORK_PACKAGES
    cfg_sched = wp_schedule if wp_schedule is not None else _config.WP_SCHEDULE
    cfg_layout = build_config_sheet(
        cfg_ws, wb, cfg_params, cfg_wps, wp_schedule=cfg_sched,
        team=_config.TEAM, reporting_periods=_config.REPORTING_PERIODS,
        grade_rates=_config.GRADE_RATES,
        assignments=_config.DEMO_CAPACITY_ASSIGNMENTS,
        external_commitments=_config.EXTERNAL_COMMITMENTS,
    )

    plan_ws = wb.create_sheet("Plan")
    layout = build_plan_sheet(
        plan_ws, wb,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        config_layout=cfg_layout,
    )

    cap_ws = wb.create_sheet("Capacity")
    cap_layout = build_capacity_sheet(
        cap_ws, cfg_layout,
        project_params=cfg_params, work_packages=cfg_wps, wp_schedule=cfg_sched,
        team=_config.TEAM, external_commitments=_config.EXTERNAL_COMMITMENTS,
        assignments=_config.DEMO_CAPACITY_ASSIGNMENTS,
        config_first_wp_row=layout["config_first_wp_row"],
    )

    _apply_planned_formulas(
        plan_ws, layout, cap_layout,
        work_packages=work_packages,
        project_params=project_params,
        wp_schedule=wp_schedule,
    )

    # Snapshot the planned FTE block: rows × all month cols
    first_row = layout["planned_first_wp_row"]
    last_row = first_row + len(cfg_wps) - 1
    snapshot = {}
    for row in range(first_row, last_row + 1):
        for col in range(MONTH_COL_START, MONTH_COL_START + PROJECT_MONTHS):
            snapshot[(row, col)] = plan_ws.cell(row, col).value
    return layout, snapshot


@pytest.mark.parametrize("kwargs", [
    {"project_params": None},
    {"work_packages": None},
    {"wp_schedule": None},
    {"project_params": None, "work_packages": None, "wp_schedule": None},
])
def test_build_plan_none_defaults_match_explicit(kwargs):
    """None-defaulted params produce the same Plan sheet as explicit config values.

    Exercises the fallback branches in build_plan_sheet and _apply_planned_formulas
    (lines 14-22 and 86-93 of plan.py). Without this test, mutmut reports those
    branch-guard mutations as surviving.
    """
    explicit = {
        "project_params": _config.default_project_params(),
        "work_packages": _config.WORK_PACKAGES,
        "wp_schedule": _config.WP_SCHEDULE,
    }
    _, baseline = _build_plan_standalone(**explicit)

    mixed = dict(explicit)
    mixed.update(kwargs)
    _, variant = _build_plan_standalone(**mixed)

    assert variant == baseline, (
        f"None-defaulted {list(kwargs)} produced different Plan formulas"
    )


@pytest.mark.parametrize("wp_idx", [0, 2])
def test_planned_fte_cells_have_black_font(v2_workbook, v2_layout, wp_idx):
    """Monthly cells in planned block have black font (formula, not input)."""
    plan = v2_workbook["Plan"]
    row = v2_layout["planned_first_wp_row"] + wp_idx
    wp = WORK_PACKAGES[wp_idx]
    wp_start, wp_end = wp_month_bounds(wp["id"], WP_SCHEDULE)

    for m in range(PROJECT_MONTHS):
        month_num = m + 1
        if wp_start <= month_num <= wp_end:
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


