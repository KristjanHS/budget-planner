"""Tests for new reader functions (Config sheet, Actuals values)."""
from datetime import datetime

import pytest
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
    assert params["start"] == datetime(2026, 7, 1)
    assert params["months"] == 24
    # FTE Rate is now a formula (grade-based weighted average)
    assert isinstance(params["rate"], str) and params["rate"].startswith("=")
    assert len(wps) == 2
    assert wps[0]["id"] == "WP1"
    assert wps[0]["name"] == "Mgmt"
    assert wps[0]["fte"] == 12.0
    assert wps[1]["id"] == "WP2"
    assert wps[1]["name"] == "Dev"
    assert wps[1]["start"] == 1
    assert wps[1]["end"] == 18
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


# ── Config sheet validation (malformed workbooks) ──


class TestConfigSheetValidation:
    """read_config_sheet rejects malformed Config sheets with clear errors."""

    def _make_empty_config_ws(self):
        """Create a bare Config worksheet with no content."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Config"
        return wb, ws

    def test_missing_title_label(self):
        """Raises ValueError when no row has Project Title."""
        wb, ws = self._make_empty_config_ws()
        # Leave sheet empty
        with pytest.raises(ValueError, match="missing param row matching 'Project Title'"):
            read_config_sheet(ws)
        wb.close()

    def test_wrong_title_label(self):
        """Raises ValueError when no row matches Project Title."""
        wb, ws = self._make_empty_config_ws()
        ws.cell(1, 1, "Name")
        with pytest.raises(ValueError, match="missing param row matching 'Project Title'"):
            read_config_sheet(ws)
        wb.close()

    def test_missing_start_label(self):
        """Raises ValueError when Project Start label is missing."""
        wb, ws = self._make_empty_config_ws()
        ws.cell(1, 1, "Project Title")
        with pytest.raises(ValueError, match="missing param row matching 'Project Start'"):
            read_config_sheet(ws)
        wb.close()

    def test_missing_months_label(self):
        """Raises ValueError when Project Months label is missing."""
        wb, ws = self._make_empty_config_ws()
        ws.cell(1, 1, "Project Title")
        ws.cell(2, 1, "Project Start")
        with pytest.raises(ValueError, match="missing param row matching 'Project Months'"):
            read_config_sheet(ws)
        wb.close()

    def test_missing_rate_label(self):
        """Raises ValueError when FTE Rate label is missing."""
        wb, ws = self._make_empty_config_ws()
        ws.cell(1, 1, "Project Title")
        ws.cell(2, 1, "Project Start")
        ws.cell(3, 1, "Project Months")
        with pytest.raises(ValueError, match="missing param row matching 'FTE Rate'"):
            read_config_sheet(ws)
        wb.close()

    def test_shifted_rows_still_found(self):
        """Label-scan finds params regardless of row position."""
        wb, ws = self._make_empty_config_ws()
        # Params shifted down by 1 — scanner still finds them
        ws.cell(1, 1, "Project ID")  # inserted row (ignored)
        ws.cell(2, 1, "Project Title")
        ws.cell(2, 2, "Test")
        ws.cell(3, 1, "Project Start")
        ws.cell(3, 2, datetime(2026, 1, 1))
        ws.cell(4, 1, "Project Months")
        ws.cell(4, 2, 60)
        ws.cell(5, 1, "FTE Rate (EUR)")
        ws.cell(5, 2, 7500)
        ws.cell(6, 1, "EU Annual Ceiling (PM)")
        ws.cell(6, 2, 12.0)
        # Need a WORK PACKAGES section for full parse
        ws.cell(8, 1, "WORK PACKAGES")
        ws.cell(9, 1, "WP ID")
        params, wps = read_config_sheet(ws)
        assert params["title"] == "Test"
        assert params["rate"] == 7500
        wb.close()

    def test_valid_config_passes_validation(self):
        """A properly built Config sheet passes validation."""
        wb = _make_config_wb()
        ws = wb["Config"]
        params, wps = read_config_sheet(ws)
        assert params["title"] == "Test"
        assert len(wps) == 2
        wb.close()


# ── Actuals reader validation (malformed workbooks) ──


class TestActualsReaderValidation:
    """read_actuals_values rejects malformed Actuals sheets."""

    def test_missing_wp_id_header(self):
        """Raises ValueError when Actuals sheet has no 'WP ID' header."""
        wb = Workbook()
        plan_ws = wb.active
        assert plan_ws is not None
        plan_ws.title = "Plan"

        actuals_ws = wb.create_sheet("Actuals")
        actuals_ws.cell(1, 1, "ACTUAL EUR")
        actuals_ws.cell(2, 1, "")
        # Row 3 is missing "WP ID" — has something else instead
        actuals_ws.cell(3, 1, "Identifier")
        actuals_ws.cell(3, 2, "Name")

        with pytest.raises(ValueError, match="Cannot find 'WP ID' header row"):
            read_actuals_values(actuals_ws, project_months=3)
        wb.close()

    def test_wp_id_header_at_nonstandard_row(self):
        """Actuals reader finds 'WP ID' header even if not at row 3."""
        wb = Workbook()
        plan_ws = wb.active
        assert plan_ws is not None
        plan_ws.title = "Plan"
        plan_ws.cell(6, 1, "WP1")

        actuals_ws = wb.create_sheet("Actuals")
        actuals_ws.cell(1, 1, "ACTUAL EUR")
        actuals_ws.cell(2, 1, "")
        actuals_ws.cell(3, 1, "")  # extra blank row
        actuals_ws.cell(4, 1, "")  # another blank row
        # WP ID header shifted to row 5 (non-standard)
        actuals_ws.cell(5, 1, "WP ID")
        actuals_ws.cell(5, 2, "WP Name")
        # WP data at row 6
        actuals_ws.cell(6, 1, "=Plan!A6")
        actuals_ws.cell(6, 2, "=Plan!B6")
        actuals_ws.cell(6, 3, 42.0)
        # TOTAL row
        actuals_ws.cell(7, 1, "")
        actuals_ws.cell(7, 2, "TOTAL")

        result = read_actuals_values(actuals_ws, project_months=1)
        assert "WP1" in result
        assert result["WP1"] == [42.0]
        wb.close()


def test_read_reporting_periods(fresh_wb_path):
    """Reader extracts reporting periods from Config sheet."""
    from openpyxl import load_workbook
    from budget_planner.reader import read_reporting_periods
    from budget_planner.config import REPORTING_PERIODS

    wb = load_workbook(fresh_wb_path)
    ws = wb["Config"]
    rps = read_reporting_periods(ws)
    wb.close()

    assert len(rps) == len(REPORTING_PERIODS)
    for actual, expected in zip(rps, REPORTING_PERIODS):
        assert actual["id"] == expected["id"]
        assert actual["start"] == expected["start"]
        assert actual["end"] == expected["end"]


def test_read_reporting_periods_fallback():
    """Reader returns default periods when section is missing."""
    from openpyxl import Workbook
    from budget_planner.reader import read_reporting_periods

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    # Empty sheet — no REPORTING PERIODS section
    ws.cell(1, 1, "Project Title")
    ws.cell(1, 2, "Test")

    rps = read_reporting_periods(ws, fallback_months=14)
    wb.close()

    # Should get default 6-month periods for 14-month project
    assert len(rps) == 3
    assert rps[0] == {"id": "RP1", "start": 1, "end": 6}
    assert rps[1] == {"id": "RP2", "start": 7, "end": 12}
    assert rps[2] == {"id": "RP3", "start": 13, "end": 14}
