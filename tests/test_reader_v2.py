"""Tests for reader extensions: read_plan_config,
read_team_sheet, read_capacity_sheet."""
import pytest

from budget_planner.reader import (
    discover_layout,
    read_plan_config,
    read_team_sheet,
    read_capacity_sheet,
)
from budget_planner.config import (
    WORK_PACKAGES, TEAM, PROJECT_MONTHS,
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


# ── read_team_sheet ──


def test_read_team_roster_names_and_fte(fresh_wb_path):
    """read_team_sheet returns roster with correct names and max_fte."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    team_data = read_team_sheet(wb["Team"])

    roster = team_data["roster"]
    assert len(roster) == len(TEAM)
    for got, expected in zip(roster, TEAM):
        assert got["name"] == expected["name"]
        assert got["max_fte"] == pytest.approx(expected["max_fte"])
        assert "row" in got
    wb.close()


def test_read_team_share_matrix(fresh_wb_path):
    """read_team_sheet returns shares matching config."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    team_data = read_team_sheet(wb["Team"])

    shares = team_data["shares"]
    for member in TEAM:
        name = member["name"]
        assert name in shares, f"Missing shares for {name}"
        for wp_id, expected_share in member["wps"].items():
            assert wp_id in shares[name], (
                f"{name} missing share for {wp_id}")
            assert shares[name][wp_id] == pytest.approx(expected_share), (
                f"{name} {wp_id} share mismatch")
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
    """Alice has WP1/WP2/WP3 as programme rows and ProjectX as external."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    alice = cap["person_groups"]["Alice"]
    assert set(alice["wp_values"].keys()) == {"WP1", "WP2", "WP3"}
    assert set(alice["external_values"].keys()) == {"ProjectX"}
    wb.close()


def test_read_capacity_extracts_monthly_values(fresh_wb_path):
    """Monthly values match seed values (share x fte / active_months)."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    # Alice WP1: share=0.25, fte=12.0, start=1, end=60 -> active=60
    # seed = 0.25 * 12.0 / 60 = 0.05
    alice_wp1 = cap["person_groups"]["Alice"]["wp_values"]["WP1"]
    assert len(alice_wp1) == PROJECT_MONTHS
    # First month should have the seed value
    assert alice_wp1[0] == pytest.approx(0.05, abs=0.001)

    # Alice WP3: share=0.25, fte=24.0, start=7, end=42 -> active=36
    # seed = 0.25 * 24.0 / 36 = 0.1667
    # Months 1-6 are inactive (None), month 7+ has seed
    alice_wp3 = cap["person_groups"]["Alice"]["wp_values"]["WP3"]
    for m in range(6):  # months 1-6 (indices 0-5) are inactive
        assert alice_wp3[m] is None, f"Month {m+1} should be inactive"
    assert alice_wp3[6] == pytest.approx(0.1667, abs=0.001)
    wb.close()


def test_read_capacity_handles_variable_geometry(fresh_wb_path):
    """Bob has 3 WPs + 1 external, Carol has 3 WPs + 0 external."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    bob = cap["person_groups"]["Bob"]
    assert set(bob["wp_values"].keys()) == {"WP1", "WP3", "WP7"}
    assert set(bob["external_values"].keys()) == {"ProjectX"}

    carol = cap["person_groups"]["Carol"]
    assert set(carol["wp_values"].keys()) == {"WP4", "WP5", "WP6"}
    assert len(carol["external_values"]) == 0
    wb.close()
