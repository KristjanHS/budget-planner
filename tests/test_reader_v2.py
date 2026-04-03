"""Tests for reader extensions: read_plan_config,
read_team_from_config, read_capacity_sheet."""
import pytest
from openpyxl import Workbook

from budget_planner.reader import (
    discover_layout,
    read_plan_config,
    read_team_from_config,
    read_capacity_sheet,
)
from budget_planner.config import (
    WORK_PACKAGES, TEAM, PROJECT_MONTHS, EXTERNAL_COMMITMENTS,
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
        assert got["start"] == expected["start"]
        assert got["end"] == expected["end"]
    wb.close()


# ── read_team_from_config ──


def test_read_team_from_config_names_and_fte(fresh_wb_path):
    """read_team_from_config returns roster with correct names and max_fte."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    roster = read_team_from_config(wb["Config"])

    assert len(roster) == len(TEAM)
    for got, expected in zip(roster, TEAM):
        assert got["name"] == expected["name"]
        assert got["max_fte"] == pytest.approx(expected["max_fte"])
    wb.close()


# ── read_capacity_sheet ──


def test_read_capacity_discovers_person_groups(fresh_wb_path):
    """read_capacity_sheet finds all 4 persons."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    expected_names = {m["name"] for m in TEAM}
    assert set(cap["person_groups"].keys()) == expected_names
    wb.close()


def test_read_capacity_classifies_grant_vs_external_rows(fresh_wb_path):
    """First person's programme WP rows and external rows match config."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    member = TEAM[0]
    group = cap["person_groups"][member["name"]]
    # Every person now has a row for every WP
    all_wp_ids = {wp["id"] for wp in WORK_PACKAGES}
    assert set(group["wp_values"].keys()) == all_wp_ids
    expected_ext = {
        e["project"] for e in EXTERNAL_COMMITMENTS
        if e["person"] == member["name"]
    }
    assert set(group["external_values"].keys()) == expected_ext
    wb.close()


def test_read_capacity_monthly_values_are_empty(fresh_wb_path):
    """Monthly values are all None — no share-based seeding."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    # Pick first person's first WP
    member = TEAM[0]
    first_wp = WORK_PACKAGES[0]
    values = cap["person_groups"][member["name"]]["wp_values"][first_wp["id"]]
    assert len(values) == PROJECT_MONTHS
    # All values should be None (no seeding)
    assert all(v is None for v in values), (
        f"Expected all None, got non-None at indices "
        f"{[i for i, v in enumerate(values) if v is not None]}"
    )
    wb.close()


def test_read_capacity_handles_variable_geometry(fresh_wb_path):
    """Each person has rows for ALL WPs and matching external set."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    all_wp_ids = {wp["id"] for wp in WORK_PACKAGES}
    for member in TEAM:
        group = cap["person_groups"][member["name"]]
        assert set(group["wp_values"].keys()) == all_wp_ids, (
            f"{member['name']}: WP set mismatch"
        )
        expected_ext = {
            e["project"] for e in EXTERNAL_COMMITMENTS
            if e["person"] == member["name"]
        }
        assert set(group["external_values"].keys()) == expected_ext, (
            f"{member['name']}: external set mismatch"
        )
    wb.close()


def test_read_capacity_external_commitments_round_trip(fresh_wb_path):
    """read_capacity_sheet extracts external_commitments matching config."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
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


# ── Skeleton mode round-trip ──


def test_skeleton_capacity_all_monthly_values_none(skeleton_wb_path):
    """Skeleton workbook has all-None monthly values for every person/WP."""
    import openpyxl
    wb = openpyxl.load_workbook(skeleton_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    assert len(cap["person_groups"]) > 0, "Should discover at least one person"

    for name, group in cap["person_groups"].items():
        for wp_id, values in group["wp_values"].items():
            assert len(values) == PROJECT_MONTHS, (
                f"{name} {wp_id}: expected {PROJECT_MONTHS} values, got {len(values)}")
            assert all(v is None for v in values), (
                f"{name} {wp_id}: expected all None in skeleton, "
                f"got non-None at indices "
                f"{[i for i, v in enumerate(values) if v is not None]}")
        for ext_name, values in group["external_values"].items():
            assert len(values) == PROJECT_MONTHS, (
                f"{name} {ext_name}: expected {PROJECT_MONTHS} values, "
                f"got {len(values)}")
            assert all(v is None for v in values), (
                f"{name} {ext_name}: expected all None in skeleton, "
                f"got non-None at indices "
                f"{[i for i, v in enumerate(values) if v is not None]}")
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
        with pytest.raises(ValueError, match="Cannot find 'PLANNED'"):
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
