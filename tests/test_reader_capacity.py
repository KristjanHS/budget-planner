"""Reader tests for Config / Capacity / Plan / Actuals surfaces.

Consolidates the historical ``test_reader_v2.py`` (real-fixture round-trips
via ``fresh_wb_path`` / ``assigned_wb_path``) and ``test_reader_new.py``
(synthetic ``_make_config_wb()`` parse-coverage). Two scopes are kept
distinct on purpose:

* ``_make_config_wb()`` and other synthetic builders exercise fresh-parse
  behaviour against minimal hand-built sheets — fast, isolated, no
  dependence on the demo seeds.
* ``fresh_wb_path`` / ``assigned_wb_path`` exercise the full builder →
  reader integration so the reader sees real cell types, formulas and
  named ranges as the prod builder emits them.

Out-of-scope and lives in sibling files:
* ``tests/test_reader.py`` — legacy ``GRANT_ACCEPT_CEILING_MIGRATION``
  env-guarded migration paths.
* ``tests/test_reader_team_priority.py`` — TEAM ROSTER Priority field
  round-trip + clamping.
"""
import pathlib

import pytest
from openpyxl import Workbook, load_workbook

from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS,
    EXTERNAL_COMMITMENTS,
    PROJECT_MONTHS,
    PROJECT_START,
    TEAM,
    WORK_PACKAGES,
)
from budget_planner.config_sheet import build_config_sheet
from budget_planner.layout import (
    COL_ROSTER_FTE,
    COL_WP_RATE,
    SECTION_HEADER_LABELS,
)
from budget_planner.reader import (
    _data_row_after,
    _find_header_row,
    discover_layout,
    read_assignments_from_config,
    read_capacity_sheet,
    read_config_sheet,
    read_externals_from_config,
    read_plan_config,
    read_team_from_config,
)
from budget_planner.workbook import init_workbook


# ── Synthetic Config builder (fresh-parse coverage) ──


def _make_config_wb():
    """Create a workbook with a Config sheet for testing."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {
        "title": "Test",
        "start": PROJECT_START,
        "months": 24,
        "rate": 7500,
    }
    wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0,
         "start": 1, "end": 24, "type": "flat"},
        {"id": "WP2", "name": "Dev",  "fte":  8.0,
         "start": 1, "end": 18, "type": "flat"},
    ]
    build_config_sheet(ws, wb, params, wps)
    return wb


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
        "start": PROJECT_START,
        "months": 60,
        "rate": 7500,
        "eu_ceiling_pm": 12.0,
    }
    wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0,
         "start": 1, "end": 60, "type": "flat"},
        {"id": "WP2", "name": "Dev", "fte": 8.0,
         "start": 1, "end": 60, "type": "flat"},
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

    section = _find_header_row(
        ws, SECTION_HEADER_LABELS["team_roster"], required=True
    )
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
    wb.defined_names["roster_ftes"] = DefinedName(
        name="roster_ftes", attr_text="Config!$B$12:$B$14"
    )

    roster = read_team_from_config(ws)
    names = [p["name"] for p in roster]
    assert names == ["Alice", "Bob"], f"Got {names}"


# ── read_capacity_sheet (with assignments) ──


def test_read_capacity_discovers_person_groups(assigned_wb_path):
    """read_capacity_sheet finds all persons."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    expected_names = {m["name"] for m in TEAM}
    assert set(cap["person_groups"].keys()) == expected_names
    wb.close()


def test_read_capacity_classifies_grant_vs_external_rows(assigned_wb_path):
    """First person's programme WP rows match their assignments."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    member = TEAM[0]
    group = cap["person_groups"][member["name"]]
    # Person has only their assigned WPs
    expected_wps = {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS
                    if a["person"] == member["name"]}
    assert set(group["wp_values"].keys()) == expected_wps
    expected_ext = {
        e["project"] for e in EXTERNAL_COMMITMENTS
        if e["person"] == member["name"]
    }
    assert set(group["external_values"].keys()) == expected_ext
    wb.close()


def test_read_capacity_assigned_values_are_distributed(assigned_wb_path):
    """Assigned WP monthly values have auto-distributed non-None values."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    # Pick first person's first assignment
    asgn = DEMO_CAPACITY_ASSIGNMENTS[0]
    values = cap["person_groups"][asgn["person"]]["wp_values"][asgn["wp"]]
    assert len(values) == PROJECT_MONTHS
    # Active months should have non-None values
    active_values = [values[m - 1] for m in range(asgn["start"], asgn["end"] + 1)]
    assert all(v is not None for v in active_values), (
        "Expected distributed values, got None at some active months"
    )
    # Total should match total_fte
    assert sum(v for v in active_values if v is not None) == pytest.approx(
        asgn["total_fte"], rel=0.01
    )
    wb.close()


def test_read_capacity_external_commitments_round_trip(assigned_wb_path):
    """read_capacity_sheet extracts external_commitments matching config."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    ext = cap["external_commitments"]
    assert len(ext) == len(EXTERNAL_COMMITMENTS)
    for got, expected in zip(ext, EXTERNAL_COMMITMENTS):
        assert got["person"] == expected["person"]
        assert got["project"] == expected["project"]
        assert got["fte"] == pytest.approx(expected["fte"])
        assert got["start"] == expected["start"]
        assert got["end"] == expected["end"]
    wb.close()


def test_read_capacity_prefers_config_assignments():
    """When Config has assignments, read_capacity_sheet uses them over Capacity."""
    assignments = _sample_assignments()
    externals = _sample_externals()

    # Build a workbook with Config sheet that has assignments + externals
    wb = _make_config_wb_with_assignments(
        assignments=assignments, externals=externals
    )

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
    wb.defined_names["assign_persons"] = DefinedName(
        name="assign_persons", attr_text="Config!$A$22:$A$24"
    )

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
    wb.defined_names["ext_persons"] = DefinedName(
        name="ext_persons", attr_text="Config!$A$32:$A$34"
    )

    out = read_externals_from_config(ws)
    persons = [e["person"] for e in out]
    assert persons == ["Alice", "Bob"], f"Got {persons}"


# ── Init workbook (no assignments) ──


def test_init_capacity_no_wp_rows(fresh_wb_path):
    """Init workbook without assignments has no WP rows per person."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    assert len(cap["person_groups"]) > 0, "Should discover at least one person"

    for name, group in cap["person_groups"].items():
        # No WP rows in plain init (no assignments)
        assert len(group["wp_values"]) == 0, (
            f"{name}: expected no WP rows in plain init, "
            f"got {set(group['wp_values'].keys())}")
    wb.close()


# ── discover_layout validation (malformed Plan sheets) ──


class TestDiscoverLayoutValidation:
    """discover_layout rejects malformed Plan sheets with clear errors."""

    def test_missing_planned_header(self):
        """Raises ValueError when PLANNED block header is missing."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Plan"
        ws.cell(1, 1, "Title:")
        ws.cell(1, 2, "Test Project")
        with pytest.raises(ValueError, match="PLANNED FTE"):
            discover_layout(ws)
        wb.close()

    def test_missing_wp_data_rows_in_planned_block(self):
        """Raises ValueError when PLANNED block has no WP data rows."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Plan"
        # Compact header
        ws.cell(1, 1, "Title:")
        # PLANNED header present but no WP data follows
        ws.cell(3, 1, "PLANNED FTE (PERSON-MONTHS)")
        ws.cell(4, 1, "")  # year row (empty)
        ws.cell(5, 1, "WP ID")  # month header row
        # No WP formula rows follow
        with pytest.raises(ValueError, match="Cannot find first WP data row"):
            discover_layout(ws)
        wb.close()

    def test_no_wp_budget_section_raises(self):
        """Plan has literal WP IDs (no =Config!A{n} formula) AND Config is
        missing the WP BUDGET section → discover_layout must raise instead of
        silently falling back to row 8, which would write WP data to the wrong
        Config rows on rebuild (T3)."""
        wb = Workbook()
        plan_ws = wb.active
        assert plan_ws is not None
        plan_ws.title = "Plan"
        plan_ws.cell(1, 1, "Title:")
        plan_ws.cell(3, 1, "PLANNED FTE (PERSON-MONTHS)")
        plan_ws.cell(4, 1, "")
        plan_ws.cell(5, 1, "WP ID")
        # Literal WP ID (not a =Config!A{n} formula) — exercises the fallback
        plan_ws.cell(6, 1, "WP1")
        plan_ws.cell(6, 2, "Package 1")
        # Config sheet exists but has no WP BUDGET section header
        wb.create_sheet("Config")
        with pytest.raises(ValueError, match="WP Budget section"):
            discover_layout(plan_ws)
        wb.close()

    def test_discover_layout_finds_wp_formula_rows(self, fresh_wb_path):
        """discover_layout correctly finds WP data rows via formula scan."""
        import openpyxl
        wb = openpyxl.load_workbook(fresh_wb_path)
        layout = discover_layout(wb["Plan"])
        num_wps = len(WORK_PACKAGES)
        # Verify the layout is internally consistent
        assert layout["planned_last_wp_row"] == layout["planned_first_wp_row"] + num_wps - 1
        assert layout["planned_total_row"] == layout["planned_last_wp_row"] + 1
        # Verify first WP row has a formula in col A
        first_val = wb["Plan"].cell(layout["planned_first_wp_row"], 1).value
        assert str(first_val).startswith("="), (
            f"Expected formula at planned first WP row, got {first_val!r}"
        )
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
        ws.cell(3, 2, PROJECT_START)
        ws.cell(4, 1, "Project Months")
        ws.cell(4, 2, 60)
        ws.cell(5, 1, "FTE Rate (EUR)")
        ws.cell(5, 2, 7500)
        ws.cell(6, 1, "EU Annual Ceiling (PM)")
        ws.cell(6, 2, 12.0)
        # Need a WP BUDGET section for full parse
        ws.cell(8, 1, "WP BUDGET")
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
        with pytest.raises(
            ValueError, match="Legacy Config layout detected"
        ):
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
            "rate": 7500,
        }
        wps = [
            {"id": "WP1", "name": "Mgmt", "fte": 12.0,
             "start": 1, "end": 24, "type": "flat"},
            {"id": "WP1", "name": "Dup",  "fte":  8.0,
             "start": 1, "end": 18, "type": "flat"},
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


# ── _iter_named_range_rows helper ──


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
    from budget_planner.reader import _iter_named_range_rows
    wb = Workbook()
    assert _iter_named_range_rows(wb, "does_not_exist") is None


# ── Actuals reader validation (malformed workbooks) ──


class TestActualsReaderValidation:
    """read_actuals_values rejects malformed Actuals sheets."""

    def test_missing_wp_id_header(self):
        """Raises ValueError when Actuals sheet has no 'WP ID' header."""
        from budget_planner.reader import read_actuals_values

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
        from budget_planner.reader import read_actuals_values

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


# ── T1: Formula rate round-trip ──


def _find_row_by_label(ws, label_prefix):
    """Return the 1-based row index of the first col-A cell whose value
    starts with ``label_prefix`` (case-insensitive, substring)."""
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        if label_prefix.lower() in label.lower():
            return r
    raise AssertionError(f"Label '{label_prefix}' not found in column A")


def _set_cell(ws, row: int, col: int, value) -> None:
    """Set a cell value with pyright-safe type narrowing (skips MergedCell)."""
    from openpyxl.cell import Cell
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell), (
        f"Expected regular Cell at ({row}, {col}); got {type(cell).__name__}"
    )
    cell.value = value


class TestFormulaRateRoundTrip:
    """reader.py preserves formula strings for WP rates.

    FTE Rate (B4) is a dashboard readout — the builder always regenerates
    it from the live grade/roster layout, so custom B4 formulas are
    intentionally not preserved. See config_sheet.build_config_sheet.
    """

    def test_formula_rate_roundtrip_wp_rate(self, tmp_path):
        """Formula in per-WP Rate cell survives read AND is rewritten on save."""
        path = str(tmp_path / "wp_rate_formula.xlsx")
        init_workbook(output_path=path)

        # Overwrite the first WP's rate cell with a formula
        wb = load_workbook(path)
        cfg = wb["Config"]
        _, original_wps = read_config_sheet(cfg)
        # Find the WP Budget first data row by locating "WP BUDGET" header
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2  # header + column-headers row
        formula = "=Config!$B$4*1.1"
        _set_cell(cfg, first_wp_row, COL_WP_RATE, formula)
        wb.save(path)
        wb.close()

        # Re-read: rate_raw should preserve formula, rate should fall back to default
        wb = load_workbook(path)
        _, wps_after = read_config_sheet(wb["Config"])
        wb.close()

        assert len(wps_after) == len(original_wps)
        assert wps_after[0]["rate_raw"] == formula, (
            f"Expected rate_raw to preserve formula; got {wps_after[0]['rate_raw']!r}"
        )
        assert isinstance(wps_after[0]["rate"], (int, float)), (
            "numeric `rate` should remain numeric for downstream consumers"
        )

        # Now rewrite via the config-sheet builder and confirm the formula
        # propagates back to the cell.
        out_wb = Workbook()
        out_ws = out_wb.active
        assert out_ws is not None
        out_ws.title = "Config"
        params_sample = {
            "title": "Rebuild Test",
            "start": PROJECT_START,
            "months": 60,
            "rate": 7500,
            "eu_ceiling_pm": 12.0,
        }
        build_config_sheet(out_ws, out_wb, params_sample, wps_after)
        # Find the rate cell on the rebuilt sheet
        rebuilt_wp_header = _find_row_by_label(out_ws, "WP BUDGET")
        rebuilt_first_wp = rebuilt_wp_header + 2
        rebuilt_rate_val = out_ws.cell(rebuilt_first_wp, COL_WP_RATE).value
        assert rebuilt_rate_val == formula, (
            f"Builder lost formula on rewrite; got {rebuilt_rate_val!r}"
        )

    def test_formula_wp_rate_roundtrip(self, tmp_path):
        """End-to-end: per-WP Rate formula survives init → edit → rebuild."""
        path = str(tmp_path / "wp_rate_roundtrip.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2
        formula = "=B2*1.1"
        _set_cell(cfg, first_wp_row, COL_WP_RATE, formula)
        wb.save(path)
        wb.close()

        # Full rebuild must preserve the per-WP formula in its Config cell
        from budget_planner.workbook import rebuild_workbook
        rebuild_workbook(input_path=path)

        wb = load_workbook(path)
        rebuilt_val = wb["Config"].cell(first_wp_row, COL_WP_RATE).value
        wb.close()
        assert rebuilt_val == formula, (
            f"Rebuild clobbered user WP Rate formula; got {rebuilt_val!r}"
        )

    def test_numeric_rate_still_works(self, tmp_path):
        """Baseline: a numeric per-WP rate round-trips as a float."""
        path = str(tmp_path / "wp_rate_numeric.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2
        _set_cell(cfg, first_wp_row, COL_WP_RATE, 7500)
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        _, wps_after = read_config_sheet(wb["Config"])
        wb.close()

        assert wps_after[0]["rate"] == pytest.approx(7500.0)
        # rate_raw captures the literal cell value (here: numeric)
        assert wps_after[0]["rate_raw"] == 7500


# ── T2: Anchored-label section header contract ──


class TestExactSectionHeaderMatching:
    """reader._find_header_row is exact-match, not substring — and
    layout drift (renamed sections, inserted blank rows) is detected loudly."""

    def test_exact_header_match_required(self, tmp_path):
        """Renaming the 'WP BUDGET' cell to a substring-containing label
        no longer matches; read_config_sheet raises ValueError with the
        canonical label in the message."""
        path = str(tmp_path / "exact_match.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_budget_row = _find_row_by_label(cfg, "WP BUDGET")
        # Rename to something that still contains "WP BUDGET" as a substring
        # (under the old substring matcher this would have matched).
        _set_cell(cfg, wp_budget_row, 1, "Legacy WP Budget (archived)")
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        with pytest.raises(ValueError, match="WP BUDGET"):
            read_config_sheet(wb["Config"])
        wb.close()

    def test_blank_row_between_header_and_data_detected(self, tmp_path):
        """Inserting a blank row between the 'WP BUDGET' title row and the
        data rows must NOT silently succeed.

        Under the old substring-match + hardcoded +2 offset, the reader would
        start scanning at title_row+2 — which after a blank insert lands on
        the original column-header row. It would then treat 'WP ID' as a WP
        identifier and call ``float('FTE Budget')`` — a surprising ValueError
        whose message points at parsing, not at layout drift. This test
        documents that a ValueError IS raised (so drift is not silently
        swallowed), which is the current, cleanly-observable outcome."""
        path = str(tmp_path / "blank_row_drift.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_budget_row = _find_row_by_label(cfg, "WP BUDGET")
        # Insert a blank row right after the section title row. This pushes
        # the column-header row (WP ID, WP Name, …) down by one, so the
        # reader's data-scan now starts on the column-header row.
        cfg.insert_rows(wp_budget_row + 1)
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        # Reader should raise — either because float('FTE Budget') fails, or
        # (better, once the reader is layout-aware) with an explicit drift
        # message. Either way: NOT a silent success.
        with pytest.raises((ValueError, TypeError)):
            read_config_sheet(wb["Config"])
        wb.close()

    def test_section_header_labels_covers_all_sections(self):
        """Every label in SECTION_HEADER_LABELS must appear as a literal
        string in some builder module, preventing drift between the
        canonical label dict and what the builders actually emit."""
        builder_dir = pathlib.Path(__file__).resolve().parent.parent / "budget_planner"
        builder_sources = {
            p.name: p.read_text()
            for p in builder_dir.glob("*.py")
            if p.name not in {"reader.py", "__init__.py", "__main__.py"}
        }
        for key, label in SECTION_HEADER_LABELS.items():
            # The literal label string must appear in at least one non-reader
            # module (builders, layout helper, or workbook orchestrator).
            quoted_variants = (f'"{label}"', f"'{label}'")
            found_in = [
                name for name, src in builder_sources.items()
                if any(q in src for q in quoted_variants)
                or f'SECTION_HEADER_LABELS["{key}"]' in src
            ]
            assert found_in, (
                f"Label {key!r}={label!r} is declared in SECTION_HEADER_LABELS "
                f"but no builder module references it (searched "
                f"{sorted(builder_sources)})"
            )
