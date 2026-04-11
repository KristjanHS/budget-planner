"""Tests for the team roster section in the Config sheet."""

import pytest
from openpyxl import Workbook
from datetime import datetime

from budget_planner.config_sheet import build_config_sheet
from budget_planner.config import TEAM


def _sample_params():
    return {
        "title": "Test",
        "start": datetime(2026, 7, 1),
        "months": 24,
        "rate": 7500,
    }


def _sample_wps():
    return [
        {"id": "WP1", "name": "Management", "fte": 12.0, "start": 1, "end": 24},
    ]


@pytest.fixture(scope="module")
def config_wb():
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    return wb, ws, layout


def test_config_roster_has_all_persons(config_wb):
    """Column A has all TEAM member names in the roster block."""
    _, ws, layout = config_wb
    names = []
    for r in range(layout["roster_first_row"], layout["roster_last_row"] + 1):
        names.append(ws.cell(row=r, column=1).value)
    expected = [m["name"] for m in TEAM]
    assert names == expected


def test_config_roster_max_fte_values(config_wb):
    """Column B has matching max_fte values for each team member."""
    _, ws, layout = config_wb
    for member in TEAM:
        row = layout["person_rows"][member["name"]]
        cell_val = ws.cell(row=row, column=2).value
        assert cell_val == member["max_fte"], (
            f"{member['name']}: expected {member['max_fte']}, got {cell_val}"
        )


def test_config_roster_cells_are_blue_input(config_wb):
    """Person name, max_fte, and grade cells have blue font color."""
    _, ws, layout = config_wb
    for r in range(layout["roster_first_row"], layout["roster_last_row"] + 1):
        for col in (1, 2, 3):  # Person, Max FTE, Grade
            font = ws.cell(row=r, column=col).font
            assert font.color and font.color.rgb in ("0000FF", "000000FF"), (
                f"Row {r} col {col} font color: {font.color}"
            )


def test_config_roster_below_wp_table(config_wb):
    """Roster section starts after the WP table with a gap row."""
    _, _, layout = config_wb
    # Roster header row should be at least 2 rows after wp_last_row (gap + section header)
    assert layout["roster_header_row"] > layout["wp_last_row"] + 1
