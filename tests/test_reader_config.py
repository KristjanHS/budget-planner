"""Reader tests for Config-sheet parsers (excluding capacity/actuals/layout).

Split from tests/test_reader_capacity.py (Phase 2 mirror-split).

Covers:
* read_plan_config — extracts WP definitions from Plan
* read_team_from_config — roster (name, max_fte, grade, priority)
* read_assignments_from_config — person→WP assignments section
* read_externals_from_config — external commitments section
* TestConfigSheetValidation — read_config_sheet rejects malformed Config
* read_reporting_periods — RP section + fallback
* read_grade_rates — RATE BY GRADE section + fallback

Out-of-scope and lives in sibling files:
* ``tests/test_reader_capacity.py`` — capacity-reader tests.
* ``tests/test_reader_actuals.py`` — Actuals reader.
* ``tests/test_reader_layout_discovery.py`` — discover_layout validation
  + ``_iter_named_range_rows`` named-range helper.
* ``tests/test_reader.py`` — legacy ``GRANT_ACCEPT_CEILING_MIGRATION``
  env-guarded migration paths.
* ``tests/test_reader_team_priority.py`` — TEAM ROSTER Priority field
  round-trip + clamping.
"""

import pytest
from openpyxl import Workbook, load_workbook

from budget_planner.config import (
    DEFAULT_RP_LENGTH,
    DEMO_CAPACITY_ASSIGNMENTS,
    PROJECT_START,
    TEAM,
    WORK_PACKAGES,
)
from budget_planner.config_sheet import build_config_sheet
from budget_planner.layout import (
    COL_ROSTER_FTE,
    SECTION_HEADER_LABELS,
)
from budget_planner.reader import (
    _data_row_after,
    _find_header_row,
    discover_layout,
    read_assignments_from_config,
    read_config_sheet,
    read_externals_from_config,
    read_plan_config,
    read_team_from_config,
)
from tests._reader_helpers import (
    _make_config_wb,
    _make_config_wb_with_assignments,
    _sample_assignments,
    _sample_externals,
)


# ── read_plan_config ──


def test_read_plan_config_extracts_wp_definitions(fresh_wb_path):
    """read_plan_config returns list matching WORK_PACKAGES."""
    import openpyxl

    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    wps = read_plan_config(wb["Plan"], layout)

    assert len(wps) == len(WORK_PACKAGES)
    for got, expected in zip(wps, WORK_PACKAGES):
        assert got["id"] == expected["id"]
        assert got["name"] == expected["name"]
        assert got["fte"] == pytest.approx(expected["fte"])
    wb.close()


# ── read_team_from_config ──


def test_read_team_from_config_names_and_fte(fresh_wb_path):
    """read_team_from_config returns roster with correct name, max_fte, grade."""
    import openpyxl

    wb = openpyxl.load_workbook(fresh_wb_path)
    roster = read_team_from_config(wb["Config"])

    assert len(roster) == len(TEAM)
    for got, expected in zip(roster, TEAM):
        assert got["name"] == expected["name"]
        assert got["max_fte"] == pytest.approx(expected["max_fte"])
        # Grade round-trips through Config — the widened roster with a Rate
        # column doesn't affect the reader's per-person grade extraction.
        assert got["grade"] == expected["grade"]
        # Priority defaults to 2 on a fresh init (neutral).
        assert got["priority"] == 2
    wb.close()


def test_read_team_returns_grade(fresh_wb_path):
    """read_team_from_config returns grade field for each person."""
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
    wb.defined_names["roster_ftes"] = DefinedName(name="roster_ftes", attr_text="Config!$B$12:$B$14")

    roster = read_team_from_config(ws)
    names = [p["name"] for p in roster]
    assert names == ["Alice", "Bob"], f"Got {names}"


# ── read_assignments_from_config / read_externals_from_config ──


def test_read_assignments_from_config_round_trip(assigned_wb_path):
    """Assignments section on Config round-trips correctly."""
    import openpyxl

    wb = openpyxl.load_workbook(assigned_wb_path)
    assignments = read_assignments_from_config(wb["Config"])

    assert len(assignments) == len(DEMO_CAPACITY_ASSIGNMENTS)
    for got, expected in zip(assignments, DEMO_CAPACITY_ASSIGNMENTS):
        assert got["person"] == expected["person"]
        assert got["wp"] == expected["wp"]
        assert got["total_fte"] == pytest.approx(expected["total_fte"])
        assert got["start"] == expected["start"]
        assert got["end"] == expected["end"]
    wb.close()


def test_read_assignments_from_config_empty_for_plain_init(fresh_wb_path):
    """Plain init workbook has no assignments in the Config section."""
    import openpyxl

    wb = openpyxl.load_workbook(fresh_wb_path)
    assignments = read_assignments_from_config(wb["Config"])
    assert assignments == []
    wb.close()


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
    wb.defined_names["assign_persons"] = DefinedName(name="assign_persons", attr_text="Config!$A$22:$A$24")

    out = read_assignments_from_config(ws)
    persons = [a["person"] for a in out]
    assert persons == ["Alice", "Bob"], f"Got {persons}"


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


def test_read_externals_tolerates_blank_row_inside_named_range():
    """Blank inside ext_persons range must not drop later externals."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    ws.cell(30, 1, value="EXTERNAL COMMITMENTS")
    ws.cell(31, 1, value="Person")
    ws.cell(31, 2, value="Project")
    ws.cell(31, 3, value="FTE")
    ws.cell(31, 4, value="Start")
    ws.cell(31, 5, value="End")
    ws.cell(31, 6, value="EU?")
    ws.cell(32, 1, value="Alice")
    ws.cell(32, 2, value="ProjectX")
    ws.cell(32, 3, value=0.1)
    ws.cell(32, 4, value=1)
    ws.cell(32, 5, value=12)
    ws.cell(32, 6, value="No")
    # Row 33 blank.
    ws.cell(34, 1, value="Bob")
    ws.cell(34, 2, value="HorizonY")
    ws.cell(34, 3, value=0.2)
    ws.cell(34, 4, value=6)
    ws.cell(34, 5, value=18)
    ws.cell(34, 6, value="Yes")
    wb.defined_names["ext_persons"] = DefinedName(name="ext_persons", attr_text="Config!$A$32:$A$34")

    out = read_externals_from_config(ws)
    persons = [e["person"] for e in out]
    assert persons == ["Alice", "Bob"], f"Got {persons}"


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

    def test_missing_start_label(self):
        """Raises ValueError when Project Start label is missing."""
        wb, ws = self._make_empty_config_ws()
        ws.cell(1, 1, "Project Title")
        with pytest.raises(ValueError, match="missing param row matching 'Project Start'"):
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
        ws.cell(3, 2, PROJECT_START)
        ws.cell(4, 1, "Project Months")
        ws.cell(4, 2, 60)
        ws.cell(5, 1, "FTE Rate (EUR)")
        ws.cell(5, 2, 11000)  # off the DEFAULT_RATE fallback
        ws.cell(6, 1, "EU Annual Ceiling (PM)")
        ws.cell(6, 2, 12.0)
        # Need a WP BUDGET section for full parse
        ws.cell(8, 1, "WP BUDGET")
        ws.cell(9, 1, "WP ID")
        params, wps = read_config_sheet(ws)
        assert params["title"] == "Test"
        assert params["rate"] == 11000
        wb.close()

    def test_valid_config_passes_validation(self):
        """A properly built Config sheet passes validation."""
        wb = _make_config_wb()
        ws = wb["Config"]
        params, wps = read_config_sheet(ws)
        assert params["title"] == "Test"
        assert len(wps) == 2
        wb.close()

    def test_rebuild_rejects_legacy_wp_schedule_section(self):
        """Legacy two-section Config (WP SCHEDULE separate from WP BUDGET) is rejected.

        Pre-merge workbooks had a standalone WP SCHEDULE section on Config;
        under the single-range schema the reader hard-fails with a migration
        recipe rather than silently collapsing multi-row schedules.
        """
        wb = _make_config_wb()
        ws = wb["Config"]
        # Append a legacy section marker somewhere past the existing content.
        ws.cell(row=ws.max_row + 2, column=1, value="WP SCHEDULE")
        with pytest.raises(ValueError, match="Legacy Config layout detected"):
            read_config_sheet(ws)
        wb.close()

    def test_rebuild_rejects_duplicate_wp_id(self):
        """Duplicate WP id in the merged WP BUDGET section is rejected."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Config"
        params = {
            "title": "Test",
            "start": PROJECT_START,
            "months": 24,
            "rate": 10000,
        }
        wps = [
            {"id": "WP1", "name": "Mgmt", "fte": 12.0, "start": 1, "end": 24, "type": "flat"},
            {"id": "WP1", "name": "Dup", "fte": 8.0, "start": 1, "end": 18, "type": "flat"},
        ]
        build_config_sheet(ws, wb, params, wps)
        with pytest.raises(ValueError, match="Duplicate WP id"):
            read_config_sheet(ws)
        wb.close()


# ── read_reporting_periods ──


def test_read_reporting_periods(fresh_wb_path):
    """Reader extracts reporting periods from Config sheet."""
    from budget_planner.config import REPORTING_PERIODS
    from budget_planner.reader import read_reporting_periods

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
    from budget_planner.reader import read_reporting_periods

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    # Empty sheet — no REPORTING PERIODS section
    ws.cell(1, 1, "Project Title")
    ws.cell(1, 2, "Test")

    length = DEFAULT_RP_LENGTH
    rps = read_reporting_periods(ws, fallback_months=2 * length + 3)
    wb.close()

    # Two full default-length periods plus a 3-month remainder
    assert rps == [
        {"id": "RP1", "start": 1, "end": length},
        {"id": "RP2", "start": length + 1, "end": 2 * length},
        {"id": "RP3", "start": 2 * length + 1, "end": 2 * length + 3},
    ]


# ── read_grade_rates ──


def test_read_grade_rates_from_generated_wb(fresh_wb_path):
    """read_grade_rates returns correct data from a generated workbook."""
    from budget_planner.config import GRADE_RATES
    from budget_planner.reader import read_grade_rates

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
    from budget_planner.reader import read_grade_rates

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


def _grade_rate_sheet(rows):
    """Minimal Config sheet holding just a RATE BY GRADE section of ``rows``."""
    from budget_planner.layout import COL_GRADE_RATE, SECTION_HEADER_LABELS
    from budget_planner.readers.layout_discovery import _data_row_after

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    header = 1
    ws.cell(row=header, column=1, value=SECTION_HEADER_LABELS["rate_by_grade"])
    ws.cell(row=header + 1, column=1, value="Grade")
    ws.cell(row=header + 1, column=COL_GRADE_RATE, value="Rate (EUR/PM) -i")
    r = _data_row_after(header)
    for grade, rate in rows:
        ws.cell(row=r, column=1, value=grade)
        if rate is not None:
            ws.cell(row=r, column=COL_GRADE_RATE, value=rate)
        r += 1
    return wb, ws


def test_read_grade_rates_rejects_formula_rate():
    """A hand-edited formula in a rate cell raises, naming the row and grade.

    Regression: the reader used to ``break`` here, returning a truncated dict
    that surfaced as ``KeyError: '<grade>'`` inside the Timelog solver.
    """
    from budget_planner.reader import read_grade_rates

    wb, ws = _grade_rate_sheet([("Senior", 6550), ("Mid", "=B3*0.85"), ("Junior", 5010)])

    with pytest.raises(ValueError) as exc:
        read_grade_rates(ws)
    wb.close()

    msg = str(exc.value)
    assert "Mid" in msg
    assert "=B3*0.85" in msg
    assert "RATE BY GRADE row 4" in msg


def test_read_grade_rates_blank_name_ends_section():
    """A blank grade name is a section terminator, not a data error."""
    from budget_planner.reader import read_grade_rates

    wb, ws = _grade_rate_sheet([("Senior", 6550), ("Mid", 5567)])
    # Row after the last grade is blank in col A but carries unrelated content.
    ws.cell(row=5, column=1, value=None)
    ws.cell(row=5, column=2, value="TIMELOG SOLVER — STOP REASONS")

    rates = read_grade_rates(ws)
    wb.close()

    assert rates == {"Senior": 6550, "Mid": 5567}
