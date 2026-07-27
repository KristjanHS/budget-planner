"""Reader tests for the Capacity sheet (read_capacity_sheet).

After the Phase 2 mirror-split, this file's scope is narrowed to capacity
reading only. The two scopes from the original consolidated file are
preserved here:

* Synthetic ``_make_config_wb_with_assignments()`` exercises fresh-parse
  behaviour against minimal hand-built sheets — fast, isolated, no
  dependence on the demo seeds.
* ``fresh_wb_path`` / ``assigned_wb_path`` exercise the full builder →
  reader integration so the reader sees real cell types, formulas and
  named ranges as the prod builder emits them.

Out-of-scope and lives in sibling files:
* ``tests/test_reader_config.py`` — Config-section readers (team,
  assignments, externals, reporting periods, grade rates, named-range
  helpers, TestConfigSheetValidation).
* ``tests/test_reader_actuals.py`` — Actuals reader, formula-rate
  round-trip, exact section-header matching.
* ``tests/test_reader_layout_discovery.py`` — discover_layout validation.
* ``tests/test_reader.py`` — legacy ``GRANT_ACCEPT_CEILING_MIGRATION``
  env-guarded migration paths.
* ``tests/test_reader_team_priority.py`` — TEAM ROSTER Priority field
  round-trip + clamping.
"""

import pytest

from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS,
    EXTERNAL_COMMITMENTS,
    PROJECT_MONTHS,
    TEAM,
)
from budget_planner.reader import (
    discover_layout,
    read_capacity_sheet,
    read_plan_config,
)
from tests._reader_helpers import (
    _make_config_wb_with_assignments,
    _sample_assignments,
    _sample_externals,
)


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
    expected_wps = {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == member["name"]}
    assert set(group["wp_values"].keys()) == expected_wps
    expected_ext = {e["project"] for e in EXTERNAL_COMMITMENTS if e["person"] == member["name"]}
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
    assert all(v is not None for v in active_values), "Expected distributed values, got None at some active months"
    # Total should match total_fte
    assert sum(v for v in active_values if v is not None) == pytest.approx(asgn["total_fte"], rel=0.01)
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
            f"{name}: expected no WP rows in plain init, got {set(group['wp_values'].keys())}"
        )
    wb.close()
