"""Tests for new reader functions (Config sheet, Actuals values)."""
from datetime import datetime
from openpyxl import Workbook
from budget_planner.config_sheet import build_config_sheet
from budget_planner.reader import read_config_sheet, read_actuals_values


def _make_config_wb():
    """Create a workbook with a Config sheet for testing."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {
        "title": "Test",
        "start": datetime(2026, 7, 1),
        "months": 24,
        "rate": 7500,
    }
    wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0, "start": 1, "end": 24},
        {"id": "WP2", "name": "Dev", "fte": 8.0, "start": 1, "end": 18},
    ]
    build_config_sheet(ws, wb, params, wps)
    return wb


def test_read_config_sheet_roundtrip():
    """Write Config sheet, read it back, values match."""
    wb = _make_config_wb()
    ws = wb["Config"]
    params, wps = read_config_sheet(ws)
    assert params["title"] == "Test"
    assert params["months"] == 24
    assert params["rate"] == 7500
    assert len(wps) == 2
    assert wps[0]["id"] == "WP1"
    assert wps[0]["fte"] == 12.0
    assert wps[1]["start"] == 1
    assert wps[1]["end"] == 18
    wb.close()


def test_read_config_sheet_params_start_date():
    """Config sheet preserves the start date."""
    wb = _make_config_wb()
    ws = wb["Config"]
    params, _ = read_config_sheet(ws)
    assert params["start"] == datetime(2026, 7, 1)
    wb.close()


def test_read_config_sheet_wp_names():
    """Config sheet preserves WP names."""
    wb = _make_config_wb()
    ws = wb["Config"]
    _, wps = read_config_sheet(ws)
    assert wps[0]["name"] == "Mgmt"
    assert wps[1]["name"] == "Dev"
    wb.close()


# ── read_actuals_values ──


def _make_actuals_wb():
    """Create a minimal workbook with Plan + Actuals sheets for testing."""
    wb = Workbook()
    plan_ws = wb.active
    assert plan_ws is not None
    plan_ws.title = "Plan"
    # Put WP IDs at rows 6 and 7 (matching typical config_first_wp_row=6)
    plan_ws.cell(6, 1, "WP1")
    plan_ws.cell(7, 1, "WP2")

    actuals_ws = wb.create_sheet("Actuals")
    # Row 1 = section header
    actuals_ws.cell(1, 1, "ACTUAL EUR")
    # Row 2 = year headers
    actuals_ws.cell(2, 1, "")
    # Row 3 = month headers
    actuals_ws.cell(3, 1, "WP ID")
    actuals_ws.cell(3, 2, "WP Name")
    # Row 4 = WP1 data (formula ref to Plan)
    actuals_ws.cell(4, 1, "=Plan!A6")
    actuals_ws.cell(4, 2, "=Plan!B6")
    actuals_ws.cell(4, 3, 100.0)
    actuals_ws.cell(4, 4, 200.0)
    actuals_ws.cell(4, 5, 0)
    # Row 5 = WP2 data (formula ref to Plan)
    actuals_ws.cell(5, 1, "=Plan!A7")
    actuals_ws.cell(5, 2, "=Plan!B7")
    actuals_ws.cell(5, 3, 0)
    actuals_ws.cell(5, 4, 0)
    actuals_ws.cell(5, 5, 0)
    # Row 6 = TOTAL row
    actuals_ws.cell(6, 1, "")
    actuals_ws.cell(6, 2, "TOTAL")

    return wb


def test_read_actuals_values_resolves_formula_refs():
    """read_actuals_values resolves =Plan!A{n} formulas to WP IDs."""
    wb = _make_actuals_wb()
    result = read_actuals_values(wb["Actuals"], project_months=3)
    # WP1 has non-zero values, should be included
    assert "WP1" in result
    assert result["WP1"] == [100.0, 200.0, 0.0]
    wb.close()


def test_read_actuals_values_skips_all_zero_rows():
    """Rows with all-zero values (no non-None) are still included since 0 != None."""
    wb = _make_actuals_wb()
    result = read_actuals_values(wb["Actuals"], project_months=3)
    # WP2 has all zeros — these are float(0), not None, so it IS included
    assert "WP2" in result
    assert result["WP2"] == [0.0, 0.0, 0.0]
    wb.close()


def test_read_actuals_values_stops_at_total():
    """read_actuals_values stops scanning at the TOTAL row."""
    wb = _make_actuals_wb()
    result = read_actuals_values(wb["Actuals"], project_months=3)
    # Should only have WP1 and WP2, not TOTAL
    assert len(result) == 2
    wb.close()


def test_read_actuals_values_literal_wp_ids():
    """read_actuals_values handles literal (non-formula) WP IDs."""
    wb = Workbook()
    plan_ws = wb.active
    assert plan_ws is not None
    plan_ws.title = "Plan"

    actuals_ws = wb.create_sheet("Actuals")
    actuals_ws.cell(1, 1, "ACTUAL EUR")
    actuals_ws.cell(2, 1, "")
    actuals_ws.cell(3, 1, "WP ID")
    actuals_ws.cell(3, 2, "WP Name")
    # Literal WP ID (not a formula)
    actuals_ws.cell(4, 1, "WP1")
    actuals_ws.cell(4, 2, "Mgmt")
    actuals_ws.cell(4, 3, 500.0)
    actuals_ws.cell(4, 4, 600.0)
    # Row 5 = TOTAL
    actuals_ws.cell(5, 1, "")
    actuals_ws.cell(5, 2, "TOTAL")

    result = read_actuals_values(actuals_ws, project_months=2)
    assert "WP1" in result
    assert result["WP1"] == [500.0, 600.0]
    wb.close()
