"""Tests for the Capacity sheet builder."""

import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.config import (
    TEAM, WORK_PACKAGES, EXTERNAL_COMMITMENTS,
    MONTH_COL_START, MONTH_COL_END, PROJECT_MONTHS,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
)
from budget_planner.styles import col_letter
from budget_planner.capacity import build_capacity_sheet


# ── Fixtures ──

@pytest.fixture(scope="module")
def full_wb(tmp_path_factory):
    """Generate a full workbook and load it for inspection."""
    out = tmp_path_factory.mktemp("cap") / "test.xlsx"
    generate_workbook(str(out))
    wb = openpyxl.load_workbook(str(out), data_only=False)
    return wb


@pytest.fixture(scope="module")
def cap_ws(full_wb):
    return full_wb["Capacity"]


@pytest.fixture(scope="module")
def capacity_layout(tmp_path_factory):
    """Build a full workbook and capture the capacity_layout return value."""
    from openpyxl import Workbook as WB
    from budget_planner.plan import build_plan_sheet
    from budget_planner.team import build_team_sheet

    wb = WB()
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = "_META"
    plan_ws = wb.create_sheet("Plan")
    build_plan_sheet(plan_ws, wb)
    team_ws = wb.create_sheet("Team")
    team_layout = build_team_sheet(team_ws, wb)
    cap_ws = wb.create_sheet("Capacity")
    layout = build_capacity_sheet(cap_ws, wb, team_layout)
    return layout


# ── Config block headers ──

def test_capacity_external_config_block_exists(cap_ws):
    """Config block has correct column headers."""
    # Row 1 is section header, row 2 is column headers
    expected = ["Person", "Project", "Total FTE", "Start Month", "End Month"]
    for ci, h in enumerate(expected, 1):
        assert cap_ws.cell(row=2, column=ci).value == h


# ── Person groups ──

def test_capacity_has_person_group_per_team_member(capacity_layout):
    """One section per TEAM member."""
    expected_names = {m["name"] for m in TEAM}
    assert set(capacity_layout["person_groups"].keys()) == expected_names


# ── Test 6: programme WP rows per person ──

def test_capacity_person_has_grant_wp_rows(capacity_layout):
    """Each person has rows for their assigned WPs."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        assert set(group["wp_rows"].keys()) == set(member["wps"].keys()), (
            f"{member['name']}: expected WPs {set(member['wps'].keys())}, "
            f"got {set(group['wp_rows'].keys())}"
        )


# ── Test 7: External project rows ──

def test_capacity_person_has_external_project_rows(capacity_layout):
    """External project rows present for persons with commitments."""
    for entry in EXTERNAL_COMMITMENTS:
        group = capacity_layout["person_groups"][entry["person"]]
        assert entry["project"] in group["external_rows"], (
            f"{entry['person']} missing external row for {entry['project']}"
        )
    # Persons without external commitments should have empty external_rows
    persons_with_ext = {e["person"] for e in EXTERNAL_COMMITMENTS}
    for member in TEAM:
        if member["name"] not in persons_with_ext:
            group = capacity_layout["person_groups"][member["name"]]
            assert len(group["external_rows"]) == 0


# ── Test 8: TOTAL row ──

def test_capacity_person_has_total_row(cap_ws, capacity_layout):
    """TOTAL row exists per person."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        total_row = group["total_row"]
        assert cap_ws.cell(row=total_row, column=2).value == "TOTAL"


# ── Test 9: REMAINING row ──

def test_capacity_person_has_remaining_row(cap_ws, capacity_layout):
    """REMAINING row exists per person."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        remaining_row = group["remaining_row"]
        assert cap_ws.cell(row=remaining_row, column=2).value == "REMAINING"


# ── Test 10: WP rows are blue input ──

def test_capacity_wp_rows_are_blue_input(cap_ws, capacity_layout):
    """Monthly cells in programme WP rows have blue font."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for wp_id, row in group["wp_rows"].items():
            # Check first active month cell
            wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
            c = MONTH_COL_START + wp["start"] - 1  # first active month
            font = cap_ws.cell(row=row, column=c).font
            assert font.color and font.color.rgb in ("0000FF", "000000FF"), (
                f"{member['name']}/{wp_id} M{wp['start']} font: {font.color}"
            )


# ── Test 11: TOTAL row SUM formula ──

def test_capacity_total_row_is_sum_formula(cap_ws, capacity_layout):
    """TOTAL monthly cells have SUM formulas."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        total_row = group["total_row"]
        # Check first month column
        val = cap_ws.cell(row=total_row, column=MONTH_COL_START).value
        assert isinstance(val, str) and val.startswith("=SUM("), (
            f"{member['name']} TOTAL M1: expected SUM formula, got {val}"
        )


# ── Test 12: REMAINING row formula ──

def test_capacity_remaining_row_formula(cap_ws, capacity_layout):
    """REMAINING references Team max_fte minus TOTAL."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        remaining_row = group["remaining_row"]
        val = cap_ws.cell(row=remaining_row, column=MONTH_COL_START).value
        assert isinstance(val, str) and "Team!" in val, (
            f"{member['name']} REMAINING M1: expected Team! ref, got {val}"
        )


# ── Test 13: Inactive months empty ──

def _inactive_params():
    """Generate (person_name, wp_id) for WPs with inactive months.

    Returns 3 representative pairs covering: end < PROJECT_MONTHS,
    start > 1, and a second person for variety.
    """
    params = []
    for member in TEAM:
        for wp_id in member["wps"]:
            wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
            if wp["start"] > 1 or wp["end"] < PROJECT_MONTHS:
                params.append((member["name"], wp_id))
    return params[:3]


@pytest.mark.parametrize("person_name,wp_id", _inactive_params())
def test_capacity_inactive_months_empty(cap_ws, capacity_layout, person_name, wp_id):
    """Inactive months have empty value and grey fill."""
    wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
    group = capacity_layout["person_groups"][person_name]
    row = group["wp_rows"][wp_id]

    for mi in range(PROJECT_MONTHS):
        month_num = mi + 1
        c = MONTH_COL_START + mi
        if month_num < wp["start"] or month_num > wp["end"]:
            val = cap_ws.cell(row=row, column=c).value
            assert val is None or val == "" or val == 0, (
                f"{person_name}/{wp_id} M{month_num}: '{val}' in inactive month"
            )
            fill = cap_ws.cell(row=row, column=c).fill
            assert fill.fgColor and fill.fgColor.rgb in ("D9D9D9", "00D9D9D9"), (
                f"{person_name}/{wp_id} M{month_num}: expected grey fill, got {fill.fgColor}"
            )


# ── Test 14: WP row Sum formula spans MONTH_COL_START:MONTH_COL_END ──

def test_capacity_wp_row_sum_formula(cap_ws, capacity_layout):
    """programme WP rows have Sum = SUM(C{row}:BJ{row})."""
    start_cl = col_letter(MONTH_COL_START)
    end_cl = col_letter(MONTH_COL_END)
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for wp_id, row in group["wp_rows"].items():
            val = cap_ws.cell(row=row, column=VAL_COL_SUM).value
            expected = f"=SUM({start_cl}{row}:{end_cl}{row})"
            assert val == expected, (
                f"{member['name']}/{wp_id} Sum: expected {expected}, got {val}"
            )


# ── Test 15: WP row Budget formula = Team share × Plan FTE ──

def test_capacity_wp_row_budget_formula(cap_ws, capacity_layout):
    """programme WP rows have Budget = Team!{share_col}{share_row} * Plan!C{config_row}."""
    # Rebuild share_wp_columns and share_first_row from a fresh team build
    from budget_planner.team import build_team_sheet as _bts
    from openpyxl import Workbook as _WB

    tmp_wb = _WB()
    tmp_meta = tmp_wb.active
    assert tmp_meta is not None
    tmp_meta.title = "_META"
    tmp_team_ws = tmp_wb.create_sheet("Team")
    team_layout = _bts(tmp_team_ws, tmp_wb)
    share_wp_columns = team_layout["share_wp_columns"]
    share_first_row = team_layout["share_first_row"]

    for person_idx, member in enumerate(TEAM):
        group = capacity_layout["person_groups"][member["name"]]
        share_row = share_first_row + person_idx
        for wp_id, row in group["wp_rows"].items():
            share_cl = col_letter(share_wp_columns[wp_id])
            from budget_planner.capacity import _wp_config_row
            from budget_planner.config import WORK_PACKAGES
            config_row = _wp_config_row(wp_id, WORK_PACKAGES)
            val = cap_ws.cell(row=row, column=VAL_COL_BUDGET).value
            expected = f"=Team!{share_cl}{share_row}*Plan!C{config_row}"
            assert val == expected, (
                f"{member['name']}/{wp_id} Budget: expected {expected}, got {val}"
            )


# ── Test 16: WP row Variance formula = Budget − Sum ──

def test_capacity_wp_row_variance_formula(cap_ws, capacity_layout):
    """programme WP rows have Variance = {budget_col}{row} - {sum_col}{row}."""
    budget_cl = col_letter(VAL_COL_BUDGET)
    sum_cl = col_letter(VAL_COL_SUM)
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for wp_id, row in group["wp_rows"].items():
            val = cap_ws.cell(row=row, column=VAL_COL_DIFF).value
            expected = f"={budget_cl}{row}-{sum_cl}{row}"
            assert val == expected, (
                f"{member['name']}/{wp_id} Variance: expected {expected}, got {val}"
            )


# ── Test 17: External row Sum/Budget/Variance formulas ──

def test_capacity_external_row_validation_formulas(cap_ws, capacity_layout):
    """External rows: Sum same SUM pattern, Budget refs config Total FTE, Variance same."""
    start_cl = col_letter(MONTH_COL_START)
    end_cl = col_letter(MONTH_COL_END)
    budget_cl = col_letter(VAL_COL_BUDGET)
    sum_cl = col_letter(VAL_COL_SUM)

    ext_config_rows = capacity_layout["external_config_rows"]

    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for project, row in group["external_rows"].items():
            # Sum
            sum_val = cap_ws.cell(row=row, column=VAL_COL_SUM).value
            expected_sum = f"=SUM({start_cl}{row}:{end_cl}{row})"
            assert sum_val == expected_sum, (
                f"{member['name']}/{project} Sum: expected {expected_sum}, got {sum_val}"
            )

            # Budget: =C{ext_config_row}
            ext_config_row = ext_config_rows[(member["name"], project)]
            budget_val = cap_ws.cell(row=row, column=VAL_COL_BUDGET).value
            expected_budget = f"=C{ext_config_row}"
            assert budget_val == expected_budget, (
                f"{member['name']}/{project} Budget: expected {expected_budget}, got {budget_val}"
            )

            # Variance: Budget - Sum
            var_val = cap_ws.cell(row=row, column=VAL_COL_DIFF).value
            expected_var = f"={budget_cl}{row}-{sum_cl}{row}"
            assert var_val == expected_var, (
                f"{member['name']}/{project} Variance: expected {expected_var}, got {var_val}"
            )
