"""Tests for the Capacity sheet builder."""

import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.config import (
    TEAM, WORK_PACKAGES, EXTERNAL_COMMITMENTS,
    MONTH_COL_START, PROJECT_MONTHS,
)
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


# ── Test 1: Sheet exists ──

def test_capacity_sheet_exists(full_wb):
    assert "Capacity" in full_wb.sheetnames


# ── Test 2: Tab color ──

def test_capacity_tab_color_purple(cap_ws):
    tab_color = cap_ws.sheet_properties.tabColor
    assert tab_color is not None
    assert tab_color.rgb == "007030A0"


# ── Test 3: Config block headers ──

def test_capacity_external_config_block_exists(cap_ws):
    """Config block has correct column headers."""
    # Row 1 is section header, row 2 is column headers
    expected = ["Person", "Project", "Total FTE", "Start Month", "End Month"]
    for ci, h in enumerate(expected, 1):
        assert cap_ws.cell(row=2, column=ci).value == h


# ── Test 4: Config values ──

def test_capacity_external_config_values_match_seed(cap_ws, capacity_layout):
    """Config block data rows match EXTERNAL_COMMITMENTS entries."""
    for entry in EXTERNAL_COMMITMENTS:
        row = capacity_layout["external_config_rows"][(entry["person"], entry["project"])]
        assert cap_ws.cell(row=row, column=1).value == entry["person"]
        assert cap_ws.cell(row=row, column=2).value == entry["project"]
        assert cap_ws.cell(row=row, column=3).value == entry["fte"]
        assert cap_ws.cell(row=row, column=4).value == entry["start"]
        assert cap_ws.cell(row=row, column=5).value == entry["end"]


# ── Test 5: Person groups ──

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


# ── Test 15: Layout dict ──

def test_capacity_returns_layout_dict(capacity_layout):
    """Layout has all expected keys."""
    assert "config_first_row" in capacity_layout
    assert "config_last_row" in capacity_layout
    assert "year_header_row" in capacity_layout
    assert "month_header_row" in capacity_layout
    assert "person_groups" in capacity_layout
    assert "external_config_rows" in capacity_layout

    # Check person_groups structure
    for member in TEAM:
        name = member["name"]
        assert name in capacity_layout["person_groups"]
        group = capacity_layout["person_groups"][name]
        for key in ("header_row", "wp_rows", "external_rows",
                     "total_row", "remaining_row"):
            assert key in group, f"{name} missing key: {key}"
