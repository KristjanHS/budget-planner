"""Tests for Config sheet builder."""
from datetime import datetime
from openpyxl import Workbook
from budget_planner.config_sheet import build_config_sheet


def _sample_params():
    return {
        "title": "Test Project",
        "start": datetime(2026, 7, 1),
        "months": 60,
        "rate": 7500,
        "eu_ceiling_pm": 12.0,
    }


def _sample_wps():
    return [
        {"id": "WP1", "name": "Management", "fte": 12.0, "start": 1, "end": 60},
        {"id": "WP2", "name": "Design", "fte": 8.0, "start": 1, "end": 18},
    ]


def test_config_sheet_params():
    """Project parameters are written correctly."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    assert ws.cell(1, 2).value == "Test Project"
    assert ws.cell(2, 2).value == datetime(2026, 7, 1)
    assert ws.cell(3, 2).value == 60
    assert ws.cell(4, 2).value == 7500


def test_config_sheet_wp_table():
    """Work packages are written as a table."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    first = layout["wp_first_row"]
    assert ws.cell(first, 1).value == "WP1"
    assert ws.cell(first + 1, 1).value == "WP2"
    # Table exists
    assert len(ws.tables) == 1
    assert "tbl_WorkPackages" in [t.name for t in ws.tables.values()]


def test_config_sheet_named_ranges():
    """Named ranges for project_start and fte_rate exist."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    names = set(wb.defined_names.keys())
    assert "project_start" in names
    assert "fte_rate" in names


def test_config_sheet_layout_keys():
    """Layout dict has expected keys including roster."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    assert "wp_header_row" in layout
    assert "wp_first_row" in layout
    assert "wp_last_row" in layout
    assert layout["wp_last_row"] == layout["wp_first_row"] + 1  # 2 WPs
    assert "roster_header_row" in layout
    assert "roster_first_row" in layout
    assert "roster_last_row" in layout
    assert "person_rows" in layout


def test_config_sheet_blue_font():
    """WP cells and param values use blue input font."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    # Param values (col B rows 2-4) should be blue (row 1 uses TITLE_FONT)
    assert ws.cell(2, 2).font.color.rgb == "000000FF" or ws.cell(2, 2).font.color.rgb == "0000FF"
    # WP cells should be blue
    first = layout["wp_first_row"]
    font_rgb = ws.cell(first, 1).font.color.rgb
    assert "0000FF" in str(font_rgb)


def test_config_eu_ceiling_param():
    """EU Annual Ceiling (PM) parameter row exists with default value."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    found = False
    for r in range(1, 10):
        if ws.cell(r, 1).value and "EU Annual Ceiling" in str(ws.cell(r, 1).value):
            assert ws.cell(r, 2).value == 12.0
            found = True
            break
    assert found, "EU Annual Ceiling param not found"


def test_config_eu_ceiling_named_range():
    """Named range eu_annual_ceiling_pm exists."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    names = set(wb.defined_names.keys())
    assert "eu_annual_ceiling_pm" in names


def test_config_round_trip_eu_ceiling():
    """EU ceiling param survives write → read round-trip."""
    from budget_planner.reader import read_config_sheet
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {**_sample_params(), "eu_ceiling_pm": 10.0}
    build_config_sheet(ws, wb, params, _sample_wps())
    read_params, read_wps = read_config_sheet(ws)
    assert read_params["eu_ceiling_pm"] == 10.0
    assert len(read_wps) == 2  # WPs still read correctly after row shift
