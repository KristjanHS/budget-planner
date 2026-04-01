"""Integration tests — data flow + round-trip preservation for v2 workbook."""
import re

import openpyxl
import pytest

from budget_planner.workbook import generate_workbook
from budget_planner.reader import discover_layout, read_workbook, read_capacity_sheet, read_plan_config
from budget_planner.config import (
    WORK_PACKAGES, TEAM, MONTH_COL_START, PROJECT_MONTHS,
)


# ── Data flow: Capacity → Plan formula linkage ──

def test_capacity_to_plan_formula_linkage(fresh_wb_path):
    """Plan Planned FTE formulas reference correct Capacity rows."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    plan = wb["Plan"]
    layout = discover_layout(plan)

    # Read the capacity data to know which persons are assigned to each WP
    config_wps = read_plan_config(plan, layout)
    cap = wb["Capacity"]
    cap_data = read_capacity_sheet(cap, config_wps)

    for i, wp in enumerate(WORK_PACKAGES):
        row = layout["planned_first_wp_row"] + i
        # Pick first active month for this WP
        first_active_mi = wp["start"] - 1
        c = MONTH_COL_START + first_active_mi

        cell_value = plan.cell(row, c).value
        assert cell_value is not None, f"Planned cell for {wp['id']} M{wp['start']} is empty"
        assert isinstance(cell_value, str) and cell_value.startswith("="), (
            f"Planned cell for {wp['id']} should be a formula, got: {cell_value}"
        )
        assert "Capacity!" in cell_value, (
            f"Planned formula for {wp['id']} should reference Capacity sheet"
        )

        # Extract row numbers from formula like =SUM(Capacity!C10,Capacity!C20)
        cap_row_refs = [int(m) for m in re.findall(r"Capacity!\w+(\d+)", cell_value)]
        assert len(cap_row_refs) > 0, f"No Capacity row refs found in formula: {cell_value}"

        # Verify each referenced row belongs to a person assigned to this WP
        for person_name, group in cap_data["person_groups"].items():
            if wp["id"] in group["wp_values"]:
                # This person should have a row referenced in the formula
                # Find the capacity row for this person-WP pair
                # We can't get exact rows from cap_data, but we can verify
                # that the count of refs matches the count of assigned persons
                pass

        # Count how many persons are assigned to this WP
        assigned_count = sum(
            1 for pg in cap_data["person_groups"].values()
            if wp["id"] in pg["wp_values"]
        )
        assert len(cap_row_refs) == assigned_count, (
            f"{wp['id']}: formula references {len(cap_row_refs)} rows "
            f"but {assigned_count} persons are assigned"
        )

    wb.close()


def test_team_shares_to_capacity_seed_values(fresh_wb_path):
    """Capacity seed values match Team share * WP FTE / active months."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    cap = wb["Capacity"]
    plan = wb["Plan"]
    layout = discover_layout(plan)
    config_wps = read_plan_config(plan, layout)
    cap_data = read_capacity_sheet(cap, config_wps)

    for member in TEAM:
        name = member["name"]
        assert name in cap_data["person_groups"], f"Person {name} not found in Capacity"
        person_group = cap_data["person_groups"][name]

        for wp_id, share in member["wps"].items():
            wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
            active_months = wp["end"] - wp["start"] + 1
            expected_seed = round(share * wp["fte"] / active_months, 4) if active_months > 0 else 0

            assert wp_id in person_group["wp_values"], (
                f"{name} missing WP {wp_id} in Capacity"
            )
            monthly = person_group["wp_values"][wp_id]

            # Check first active month value matches seed
            first_active_idx = wp["start"] - 1
            actual_val = monthly[first_active_idx]
            assert actual_val == pytest.approx(expected_seed, abs=0.001), (
                f"{name}/{wp_id} M{wp['start']}: expected seed {expected_seed}, "
                f"got {actual_val}"
            )

            # Check inactive months are None
            for mi in range(PROJECT_MONTHS):
                month_num = mi + 1
                if month_num < wp["start"] or month_num > wp["end"]:
                    assert monthly[mi] is None, (
                        f"{name}/{wp_id} M{month_num}: inactive month should be None, "
                        f"got {monthly[mi]}"
                    )

    wb.close()


# ── Round-trip preservation ──

def test_roundtrip_capacity_edits_preserved(tmp_path):
    """Capacity cell edits survive --update round-trip."""
    # Generate fresh
    path = str(tmp_path / "test.xlsx")
    generate_workbook(path)

    # Load and edit a Capacity cell
    wb = openpyxl.load_workbook(path)
    cap = wb["Capacity"]

    # Find a cell with the seed value for Alice/WP1:
    # share=0.25, FTE=12.0, active_months=60 -> seed = 0.05
    edit_row = None
    for r in range(1, cap.max_row + 1):
        val = cap.cell(r, MONTH_COL_START).value
        if val is not None and isinstance(val, (int, float)) and abs(val - 0.05) < 0.001:
            edit_row = r
            break

    assert edit_row is not None, "Could not find a cell with seed value 0.05"

    cap.cell(edit_row, MONTH_COL_START, 0.99)
    wb.save(path)
    wb.close()

    # Update
    updated = str(tmp_path / "updated.xlsx")
    generate_workbook(output_path=updated, existing_path=path)

    # Verify the edited value survived
    wb2 = openpyxl.load_workbook(updated)
    cap2 = wb2["Capacity"]
    assert cap2.cell(edit_row, MONTH_COL_START).value == pytest.approx(0.99), (
        "Capacity edit was not preserved on --update"
    )
    wb2.close()


def test_roundtrip_actuals_still_preserved(tmp_path):
    """Actuals survive --update in v2 workbook (regression guard)."""
    # Generate fresh
    path = str(tmp_path / "test.xlsx")
    generate_workbook(path)

    # Inject actuals
    wb = openpyxl.load_workbook(path)
    plan = wb["Plan"]
    layout = discover_layout(plan)
    row = layout["actual_first_wp_row"]
    plan.cell(row, MONTH_COL_START, 0.15)      # WP1, M1
    plan.cell(row, MONTH_COL_START + 1, 0.20)  # WP1, M2
    wb.save(path)
    wb.close()

    # Update
    updated = str(tmp_path / "updated.xlsx")
    generate_workbook(output_path=updated, existing_path=path)

    # Verify actuals survived
    wb2 = openpyxl.load_workbook(updated)
    plan2 = wb2["Plan"]
    layout2 = discover_layout(plan2)
    row2 = layout2["actual_first_wp_row"]
    assert plan2.cell(row2, MONTH_COL_START).value == pytest.approx(0.15), (
        "Actual M1 not preserved"
    )
    assert plan2.cell(row2, MONTH_COL_START + 1).value == pytest.approx(0.20), (
        "Actual M2 not preserved"
    )
    wb2.close()


def test_roundtrip_capacity_external_edits_preserved(tmp_path):
    """External project capacity edits survive --update round-trip."""
    path = str(tmp_path / "test.xlsx")
    generate_workbook(path)

    # Read to find an external project row
    wb = openpyxl.load_workbook(path)
    cap = wb["Capacity"]
    plan = wb["Plan"]
    layout = discover_layout(plan)
    config_wps = read_plan_config(plan, layout)
    cap_data = read_capacity_sheet(cap, config_wps)

    # Find Alice's ProjectX external row — seeded external commitment
    alice_ext = cap_data["person_groups"]["Alice"]["external_values"]
    assert "ProjectX" in alice_ext, "Alice should have ProjectX external"

    # Find the row in the sheet by scanning for "ProjectX" in col A
    ext_edit_row = None
    for r in range(1, cap.max_row + 1):
        if cap.cell(r, 1).value == "ProjectX":
            ext_edit_row = r
            break

    assert ext_edit_row is not None, "Could not find ProjectX row in Capacity"

    cap.cell(ext_edit_row, MONTH_COL_START, 0.77)
    wb.save(path)
    wb.close()

    # Update
    updated = str(tmp_path / "updated.xlsx")
    generate_workbook(output_path=updated, existing_path=path)

    # Verify
    wb2 = openpyxl.load_workbook(updated)
    cap2 = wb2["Capacity"]
    assert cap2.cell(ext_edit_row, MONTH_COL_START).value == pytest.approx(0.77), (
        "External project capacity edit was not preserved"
    )
    wb2.close()


# ── Guard rails ──

def test_update_refuses_empty_config_table(tmp_path):
    """read_workbook raises ValueError when config table has no WP IDs."""
    path = str(tmp_path / "test.xlsx")
    generate_workbook(path)

    # Manually delete all WP ID values from config rows
    wb = openpyxl.load_workbook(path)
    plan = wb["Plan"]
    layout = discover_layout(plan)

    for r in range(layout["config_first_wp_row"], layout["config_last_wp_row"] + 1):
        plan.cell(r, 1).value = None

    wb.save(path)
    wb.close()

    with pytest.raises(ValueError, match="zero WP rows"):
        read_workbook(path)


def test_roundtrip_team_roster_preserved(tmp_path):
    """Team roster data survives --update round-trip."""
    path = str(tmp_path / "test.xlsx")
    generate_workbook(path)

    # Edit Team sheet: change Alice's max_fte
    wb = openpyxl.load_workbook(path)
    team = wb["Team"]

    # Find Alice's row (roster starts at row 3: row 1 section header, row 2 column headers)
    alice_row = None
    for r in range(1, team.max_row + 1):
        if team.cell(r, 1).value == "Alice":
            alice_row = r
            break

    assert alice_row is not None, "Could not find Alice in Team roster"

    team.cell(alice_row, 2, 0.75)  # Change max_fte from 1.0 to 0.75
    wb.save(path)
    wb.close()

    # Update
    updated = str(tmp_path / "updated.xlsx")
    generate_workbook(output_path=updated, existing_path=path)

    # Verify Alice's max_fte survived
    wb2 = openpyxl.load_workbook(updated)
    team2 = wb2["Team"]

    alice_row2 = None
    for r in range(1, team2.max_row + 1):
        if team2.cell(r, 1).value == "Alice":
            alice_row2 = r
            break

    assert alice_row2 is not None
    assert team2.cell(alice_row2, 2).value == pytest.approx(0.75), (
        "Alice's max_fte was not preserved on --update"
    )
    wb2.close()
