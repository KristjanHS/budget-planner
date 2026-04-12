"""Tests for new reader functions (Config sheet, Actuals values, grade rates)."""

from datetime import datetime

import pytest
from openpyxl import Workbook

from budget_planner.config_sheet import build_config_sheet
from budget_planner.layout import COL_ROSTER_FTE, SECTION_HEADER_LABELS
from budget_planner.reader import (
    read_config_sheet,
    read_actuals_values,
    read_grade_rates,
    read_team_from_config,
    read_assignments_from_config,
    read_externals_from_config,
    read_capacity_sheet,
    _find_header_row,
    _data_row_after,
)


def _sample_schedule():
    return [
        {"id": "WP1", "start": 1, "end": 24, "type": "flat"},
        {"id": "WP2", "start": 1, "end": 6, "type": "flat"},
    ]


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
        {"id": "WP1", "name": "Mgmt", "fte": 12.0},
        {"id": "WP2", "name": "Dev", "fte": 8.0},
    ]
    build_config_sheet(ws, wb, params, wps, wp_schedule=_sample_schedule())
    return wb


def test_read_config_sheet_roundtrip():
    """Write Config sheet, read it back, values match."""
    wb = _make_config_wb()
    ws = wb["Config"]
    params, wps, wp_sched = read_config_sheet(ws)
    assert params["title"] == "Test"
    assert params["start"] == datetime(2026, 7, 1)
    assert params["months"] == 24
    # FTE Rate is now a CSE array formula (grade-based weighted average)
    from openpyxl.worksheet.formula import ArrayFormula

    assert isinstance(params["rate"], (str, ArrayFormula))
    assert len(wps) == 2
    assert wps[0]["id"] == "WP1"
    assert wps[0]["name"] == "Mgmt"
    assert wps[0]["fte"] == 12.0
    assert wps[1]["id"] == "WP2"
    assert wps[1]["name"] == "Dev"
    # Start/end now come from wp_schedule
    wp2_sched = [s for s in wp_sched if s["id"] == "WP2"]
    assert len(wp2_sched) >= 1
    assert wp2_sched[0]["start"] == 1
    assert wp2_sched[0]["end"] == 18
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
        with pytest.raises(
            ValueError, match="missing param row matching 'Project Title'"
        ):
            read_config_sheet(ws)
        wb.close()

    def test_missing_start_label(self):
        """Raises ValueError when Project Start label is missing."""
        wb, ws = self._make_empty_config_ws()
        ws.cell(1, 1, "Project Title")
        with pytest.raises(
            ValueError, match="missing param row matching 'Project Start'"
        ):
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
        # Need a WP BUDGET section for full parse
        ws.cell(8, 1, "WP BUDGET")
        ws.cell(9, 1, "WP ID")
        params, wps, _ = read_config_sheet(ws)
        assert params["title"] == "Test"
        assert params["rate"] == 7500
        wb.close()

    def test_valid_config_passes_validation(self):
        """A properly built Config sheet passes validation."""
        wb = _make_config_wb()
        ws = wb["Config"]
        params, wps, _ = read_config_sheet(ws)
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


# ── read_grade_rates ──


def test_read_grade_rates_from_generated_wb(fresh_wb_path):
    """read_grade_rates returns correct data from a generated workbook."""
    from openpyxl import load_workbook
    from budget_planner.config import GRADE_RATES

    wb = load_workbook(fresh_wb_path)
    ws = wb["Config"]
    rates = read_grade_rates(ws)
    wb.close()

    assert len(rates) == len(GRADE_RATES)
    for grade, expected_rate in GRADE_RATES.items():
        assert grade in rates
        assert rates[grade] == expected_rate


def test_read_grade_rates_fallback_for_legacy_wb():
    """read_grade_rates returns defaults when RATE BY GRADE section missing."""
    from budget_planner.config import GRADE_RATES

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    # Empty sheet — no RATE BY GRADE section
    ws.cell(1, 1, "Project Title")
    ws.cell(1, 2, "Test")

    rates = read_grade_rates(ws)
    wb.close()

    assert rates == dict(GRADE_RATES)


# ── read_team_from_config with grade ──


def test_read_team_returns_grade(fresh_wb_path):
    """read_team_from_config returns grade field for each person."""
    from openpyxl import load_workbook
    from budget_planner.config import TEAM

    wb = load_workbook(fresh_wb_path)
    ws = wb["Config"]
    roster = read_team_from_config(ws)
    wb.close()

    assert len(roster) == len(TEAM)
    for actual, expected in zip(roster, TEAM):
        assert "grade" in actual
        assert actual["grade"] == expected.get("grade", "Mid")


def test_read_team_defaults_grade_for_legacy_wb():
    """read_team_from_config defaults to 'Mid' when Grade column missing."""
    from budget_planner.config import DEFAULT_GRADE

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"

    # Build a minimal roster section without Grade column
    ws.cell(1, 1, "TEAM ROSTER")
    ws.cell(2, 1, "Person")
    ws.cell(2, 2, "Max FTE")
    # No col C header — legacy format
    ws.cell(3, 1, "Alice")
    ws.cell(3, 2, 1.0)
    ws.cell(4, 1, "Bob")
    ws.cell(4, 2, 0.8)

    roster = read_team_from_config(ws)
    wb.close()

    assert len(roster) == 2
    assert roster[0]["name"] == "Alice"
    assert roster[0]["grade"] == DEFAULT_GRADE
    assert roster[1]["name"] == "Bob"
    assert roster[1]["grade"] == DEFAULT_GRADE


def test_read_team_accepts_european_text_max_fte_and_keeps_following_rows(
    fresh_wb_path,
):
    """Text Max FTE like '0,35' must not truncate the roster (orphan assignments)."""
    from openpyxl import load_workbook

    wb = load_workbook(fresh_wb_path)
    ws = wb["Config"]
    n_before = len(read_team_from_config(ws))
    assert n_before >= 3

    section = _find_header_row(ws, SECTION_HEADER_LABELS["team_roster"], required=True)
    person_header = section + 1
    # Second roster data row: force European-style text (Excel "number stored as text")
    ws.cell(person_header + 2, COL_ROSTER_FTE, value="0,35")

    roster = read_team_from_config(ws)
    wb.close()

    assert len(roster) == n_before
    assert roster[1]["name"] == "Bob"
    assert roster[1]["max_fte"] == pytest.approx(0.35)


# ── read_assignments_from_config / read_externals_from_config ──


def _sample_assignments():
    return [
        {"person": "Alice", "wp": "WP1", "total_fte": 4.5, "start": 1, "end": 14},
        {"person": "Bob", "wp": "WP2", "total_fte": 3.0, "start": 1, "end": 18},
    ]


def _sample_externals():
    return [
        {
            "person": "Alice",
            "project": "ProjectX",
            "fte": 3.0,
            "start": 1,
            "end": 11,
            "is_eu": False,
        },
        {
            "person": "Bob",
            "project": "HorizonEU",
            "fte": 4.0,
            "start": 1,
            "end": 24,
            "is_eu": True,
        },
    ]


def _sample_team():
    return [
        {"name": "Alice", "max_fte": 1.0, "grade": "Senior"},
        {"name": "Bob", "max_fte": 1.0, "grade": "Mid"},
    ]


def _sample_grade_rates():
    return {"Senior": 12500, "Mid": 10000}


def _make_config_wb_with_assignments(assignments=None, externals=None):
    """Create a workbook with a Config sheet that has assignments/externals."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {
        "title": "Test",
        "start": datetime(2026, 7, 1),
        "months": 60,
        "rate": 7500,
        "eu_ceiling_pm": 12.0,
    }
    wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0},
        {"id": "WP2", "name": "Dev", "fte": 8.0},
    ]
    build_config_sheet(
        ws,
        wb,
        params,
        wps,
        team=_sample_team(),
        grade_rates=_sample_grade_rates(),
        assignments=assignments or [],
        external_commitments=externals or [],
    )
    return wb


def test_read_assignments_from_config_roundtrip():
    """Build Config sheet with assignments, read back, values match."""
    assignments = _sample_assignments()
    wb = _make_config_wb_with_assignments(assignments=assignments)
    ws = wb["Config"]
    result = read_assignments_from_config(ws)
    wb.close()

    assert len(result) == len(assignments)
    for actual, expected in zip(result, assignments):
        assert actual["person"] == expected["person"]
        assert actual["wp"] == expected["wp"]
        assert actual["total_fte"] == expected["total_fte"]
        assert actual["start"] == expected["start"]
        assert actual["end"] == expected["end"]


def test_read_assignments_strips_person_wp_whitespace():
    """Person/WP strings are stripped so they match TEAM ROSTER (also stripped)."""
    assignments = _sample_assignments()
    wb = _make_config_wb_with_assignments(assignments=assignments)
    ws = wb["Config"]
    hdr = _find_header_row(ws, SECTION_HEADER_LABELS["assignments"], required=True)
    r0 = _data_row_after(hdr)
    ws.cell(r0, 1, value="Alice  ")
    ws.cell(r0, 2, value="WP1 ")

    result = read_assignments_from_config(ws)
    wb.close()

    assert result[0]["person"] == "Alice"
    assert result[0]["wp"] == "WP1"


def test_read_assignments_from_config_empty():
    """Config sheet without assignments section returns empty list."""
    wb = _make_config_wb_with_assignments()  # no assignments
    ws = wb["Config"]
    result = read_assignments_from_config(ws)
    wb.close()

    assert result == []


def test_read_externals_from_config_roundtrip():
    """Build Config sheet with externals, read back, values match."""
    externals = _sample_externals()
    wb = _make_config_wb_with_assignments(externals=externals)
    ws = wb["Config"]
    result = read_externals_from_config(ws)
    wb.close()

    assert len(result) == len(externals)
    for actual, expected in zip(result, externals):
        assert actual["person"] == expected["person"]
        assert actual["project"] == expected["project"]
        assert actual["fte"] == expected["fte"]
        assert actual["start"] == expected["start"]
        assert actual["end"] == expected["end"]
        assert actual["is_eu"] == expected["is_eu"]


def test_read_externals_from_config_empty():
    """Config sheet without externals section returns empty list."""
    wb = _make_config_wb_with_assignments()  # no externals
    ws = wb["Config"]
    result = read_externals_from_config(ws)
    wb.close()

    assert result == []


def test_read_capacity_prefers_config_assignments():
    """When Config has assignments, read_capacity_sheet uses them over Capacity."""
    assignments = _sample_assignments()
    externals = _sample_externals()

    # Build a workbook with Config sheet that has assignments + externals
    wb = _make_config_wb_with_assignments(assignments=assignments, externals=externals)

    # Build minimal Plan + Capacity sheets so _read_person_grids doesn't error
    plan_ws = wb.create_sheet("Plan")
    plan_ws.cell(1, 1, "Plan stub")

    cap_ws = wb.create_sheet("Capacity")
    # Write a bare "WP ID" header so _read_person_grids can find it
    cap_ws.cell(1, 1, "WP ID")

    config_wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0},
        {"id": "WP2", "name": "Dev", "fte": 8.0},
    ]
    result = read_capacity_sheet(cap_ws, config_wps, project_months=60)
    wb.close()

    # Assignments should come from Config, not Capacity (which has none)
    assert len(result["assignments"]) == len(assignments)
    for actual, expected in zip(result["assignments"], assignments):
        assert actual["person"] == expected["person"]
        assert actual["wp"] == expected["wp"]
        assert actual["total_fte"] == expected["total_fte"]

    # Externals should come from Config too
    assert len(result["external_commitments"]) == len(externals)
    for actual, expected in zip(result["external_commitments"], externals):
        assert actual["person"] == expected["person"]
        assert actual["project"] == expected["project"]
        assert actual["is_eu"] == expected["is_eu"]


def test_iter_named_range_rows_returns_live_row_span_after_middle_insert():
    """Helper must report the post-insertion row span via Excel's defined_names."""
    from openpyxl.workbook.defined_name import DefinedName
    from budget_planner.reader import _iter_named_range_rows

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    for i, val in enumerate(["A", "B", "C"], start=5):
        ws.cell(i, 1, value=val)
    wb.defined_names["roster_ftes"] = DefinedName(
        name="roster_ftes", attr_text="Config!$B$5:$B$7"
    )

    # Excel-style middle insert: shift rows 6+ down by 1.
    ws.insert_rows(6, amount=1)
    ws.cell(6, 1, value="MID")
    # openpyxl does NOT update DefinedName targets on insert — simulate Excel
    # by explicitly rewriting the range. The helper under test treats the
    # defined_names entry as the source of truth.
    wb.defined_names["roster_ftes"].attr_text = "Config!$B$5:$B$8"

    assert _iter_named_range_rows(wb, "roster_ftes") == (5, 8)


def test_iter_named_range_rows_handles_single_cell_ref():
    """Single-cell absolute ref (e.g. $B$5) must return (5, 5)."""
    from openpyxl.workbook.defined_name import DefinedName
    from budget_planner.reader import _iter_named_range_rows

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    wb.defined_names["project_start"] = DefinedName(
        name="project_start", attr_text="Config!$B$5"
    )
    assert _iter_named_range_rows(wb, "project_start") == (5, 5)


def test_iter_named_range_rows_returns_none_for_unknown_name():
    wb = Workbook()
    from budget_planner.reader import _iter_named_range_rows
    assert _iter_named_range_rows(wb, "does_not_exist") is None


def test_read_assignments_tolerates_blank_row_inside_named_range():
    """Blank inside assign_persons range must not drop later assignments."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    ws.cell(20, 1, value="PERSON-WP ASSIGNMENTS")
    ws.cell(21, 1, value="Person")
    ws.cell(21, 2, value="WP")
    ws.cell(21, 3, value="Total FTE")
    ws.cell(21, 4, value="Start")
    ws.cell(21, 5, value="End")
    ws.cell(22, 1, value="Alice")
    ws.cell(22, 2, value="WP1")
    ws.cell(22, 3, value=0.5)
    ws.cell(22, 4, value=1)
    ws.cell(22, 5, value=12)
    # Row 23 blank.
    ws.cell(24, 1, value="Bob")
    ws.cell(24, 2, value="WP2")
    ws.cell(24, 3, value=0.3)
    ws.cell(24, 4, value=6)
    ws.cell(24, 5, value=18)
    wb.defined_names["assign_persons"] = DefinedName(
        name="assign_persons", attr_text="Config!$A$22:$A$24"
    )

    out = read_assignments_from_config(ws)
    persons = [a["person"] for a in out]
    assert persons == ["Alice", "Bob"], f"Got {persons}"


def test_read_team_tolerates_blank_row_inside_named_range():
    """A blank row inside roster_ftes span must not truncate the roster."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    ws.cell(10, 1, value="TEAM ROSTER")
    ws.cell(11, 1, value="Person")
    ws.cell(11, 2, value="Max FTE")
    ws.cell(11, 3, value="Grade")
    ws.cell(12, 1, value="Alice")
    ws.cell(12, 2, value=1.0)
    ws.cell(12, 3, value="Senior")
    # Row 13 intentionally blank.
    ws.cell(14, 1, value="Bob")
    ws.cell(14, 2, value=0.5)
    ws.cell(14, 3, value="Junior")
    wb.defined_names["roster_ftes"] = DefinedName(
        name="roster_ftes", attr_text="Config!$B$12:$B$14"
    )

    roster = read_team_from_config(ws)
    names = [p["name"] for p in roster]
    assert names == ["Alice", "Bob"], f"Got {names}"
