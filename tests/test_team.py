"""Tests for the Team sheet builder."""

import pytest
from openpyxl import Workbook

from budget_planner.team import build_team_sheet
from budget_planner.config import TEAM, WORK_PACKAGES


@pytest.fixture(scope="module")
def team_wb():
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Team"
    layout = build_team_sheet(ws, wb)
    return wb, ws, layout


def test_team_sheet_exists(team_wb):
    """build_team_sheet creates a sheet and returns a layout dict."""
    wb, ws, layout = team_wb
    assert ws.title == "Team"
    assert isinstance(layout, dict)
    for key in (
        "roster_header_row", "roster_first_row", "roster_last_row",
        "share_header_row", "share_first_row", "share_last_row",
        "share_validation_row", "share_wp_columns", "person_rows",
    ):
        assert key in layout, f"Missing key: {key}"


def test_team_roster_has_all_persons(team_wb):
    """Column A has all TEAM member names in the roster block."""
    _, ws, layout = team_wb
    names = []
    for r in range(layout["roster_first_row"], layout["roster_last_row"] + 1):
        names.append(ws.cell(row=r, column=1).value)
    expected = [m["name"] for m in TEAM]
    assert names == expected


def test_team_roster_max_fte_values(team_wb):
    """Column B has matching max_fte values for each team member."""
    _, ws, layout = team_wb
    for member in TEAM:
        row = layout["person_rows"][member["name"]]
        cell_val = ws.cell(row=row, column=2).value
        assert cell_val == member["max_fte"], (
            f"{member['name']}: expected {member['max_fte']}, got {cell_val}"
        )


def test_team_roster_cells_are_blue_input(team_wb):
    """Person name and max_fte cells have blue font color."""
    _, ws, layout = team_wb
    for r in range(layout["roster_first_row"], layout["roster_last_row"] + 1):
        name_font = ws.cell(row=r, column=1).font
        fte_font = ws.cell(row=r, column=2).font
        assert name_font.color and name_font.color.rgb in ("0000FF", "000000FF"), (
            f"Row {r} col A font color: {name_font.color}"
        )
        assert fte_font.color and fte_font.color.rgb in ("0000FF", "000000FF"), (
            f"Row {r} col B font color: {fte_font.color}"
        )


def test_team_share_matrix_wp_columns(team_wb):
    """Share matrix has columns for all WP IDs."""
    _, ws, layout = team_wb
    wp_ids = {wp["id"] for wp in WORK_PACKAGES}
    assert set(layout["share_wp_columns"].keys()) == wp_ids
    # Verify headers match
    for wp_id, col in layout["share_wp_columns"].items():
        header_val = ws.cell(row=layout["share_header_row"], column=col).value
        assert header_val == wp_id, f"Expected {wp_id} at col {col}, got {header_val}"


def test_team_share_matrix_values_match_config(team_wb):
    """Cell values match TEAM[*].wps share allocations."""
    _, ws, layout = team_wb
    for i, member in enumerate(TEAM):
        row = layout["share_first_row"] + i
        for wp in WORK_PACKAGES:
            col = layout["share_wp_columns"][wp["id"]]
            cell_val = ws.cell(row=row, column=col).value
            expected = member["wps"].get(wp["id"])
            assert cell_val == expected, (
                f"{member['name']} / {wp['id']}: expected {expected}, got {cell_val}"
            )


def test_team_share_validation_row_has_sum_formulas(team_wb):
    """Validation row has SUM formulas for each WP column."""
    _, ws, layout = team_wb
    val_row = layout["share_validation_row"]
    for wp_id, col in layout["share_wp_columns"].items():
        cell_val = ws.cell(row=val_row, column=col).value
        assert cell_val is not None, f"No formula for {wp_id}"
        assert "SUM(" in str(cell_val).upper(), (
            f"{wp_id} validation cell: expected SUM formula, got {cell_val}"
        )


def test_team_tab_color_orange(team_wb):
    """Tab color is orange (FFA500)."""
    _, ws, _ = team_wb
    tab_color = ws.sheet_properties.tabColor
    assert tab_color is not None
    assert tab_color.rgb == "00FFA500"
