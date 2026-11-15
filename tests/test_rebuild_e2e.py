"""End-to-end rebuild tests: modify workbook data, rebuild, verify results.

These tests prove rebuild works correctly after changing FTE budgets,
WP start/end months, team names, and allocations — without hardcoding
any specific values from config.py.
"""

import json
import re
import subprocess
from pathlib import Path

import pytest
import openpyxl
from openpyxl.cell import Cell

from budget_planner.workbook import init_workbook, rebuild_workbook
from budget_planner.reader import (
    discover_layout,
    read_config_sheet,
    read_team_from_config,
    read_capacity_sheet,
    read_plan_config,
    read_assignments_from_config,
)
from budget_planner.config import MONTH_COL_START

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


def _get_config_first_wp_row():
    """Derive Config WP data start row from a generated workbook."""
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        init_workbook(path)
        wb = openpyxl.load_workbook(path)
        try:
            from budget_planner.reader import _find_header_row

            return _find_header_row(wb["Config"], "WP BUDGET") + 2
        finally:
            wb.close()
    finally:
        import os

        os.unlink(path)


CONFIG_FIRST_WP_ROW = _get_config_first_wp_row()


def _find_sched_rows(config_ws):
    """Find (row, id, start, end) entries from the merged WP BUDGET section.

    Under the single-range schema, each WP row carries Start (col H) and End
    (col I) directly; there is no separate WP SCHEDULE section.
    """
    from budget_planner.reader import _find_header_row
    from budget_planner.layout import COL_WP_START, COL_WP_END

    wp_header = _find_header_row(config_ws, "WP BUDGET")
    first_row = wp_header + 2  # skip section header + column headers
    entries = []
    r = first_row
    while r <= config_ws.max_row:
        wp_id = config_ws.cell(r, 1).value
        col_b = config_ws.cell(r, 2).value
        if wp_id is None or str(wp_id).strip() == "":
            break
        if isinstance(col_b, str) and col_b.strip().upper() == "TOTAL":
            break
        entries.append(
            {
                "row": r,
                "id": str(wp_id),
                "start": config_ws.cell(r, COL_WP_START).value,
                "end": config_ws.cell(r, COL_WP_END).value,
            }
        )
        r += 1
    return entries


# ── Helpers ──


def _set(ws, row, col, value):
    """Set a cell value."""
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _find_first_wp_row_actuals(ws):
    """Find the first WP data row in Actuals."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r + 1
    return None


def _find_total_row(ws, col=2):
    """Find TOTAL row in a sheet."""
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, col).value
        if isinstance(val, str) and val.upper() == "TOTAL":
            return r
    return None


# ── Fixtures ──


@pytest.fixture
def init_path(tmp_path):
    """Create a fresh workbook for rebuild testing."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    return path


@pytest.fixture
def assigned_init_path(tmp_path):
    """Create a fresh workbook with demo assignments for rebuild testing."""
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS

    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path, assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return path


# ── Test: Change WP FTE budgets ──


class TestRebuildWithChangedFTE:
    """Modify WP FTE values in Config sheet, rebuild, verify."""

    def test_changed_fte_round_trips(self, init_path):
        """Modified WP FTE values survive rebuild and appear in Plan config."""
        wb = openpyxl.load_workbook(init_path)
        _, original_wps = read_config_sheet(wb["Config"])
        # Double the FTE of the first WP
        new_fte = original_wps[0]["fte"] * 2
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, new_fte)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        _, rebuilt_wps = read_config_sheet(wb["Config"])
        assert rebuilt_wps[0]["fte"] == pytest.approx(new_fte)
        wb.close()

    def test_changed_fte_updates_plan_ceiling(self, init_path):
        """Plan config table reflects new FTE after rebuild."""
        wb = openpyxl.load_workbook(init_path)
        _, original_wps = read_config_sheet(wb["Config"])
        new_fte = original_wps[0]["fte"] * 3
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, new_fte)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        plan = wb["Plan"]
        layout = discover_layout(plan)
        plan_wps = read_plan_config(plan, layout)
        assert plan_wps[0]["fte"] == pytest.approx(new_fte)
        wb.close()


# ── Test: Change WP start/end months ──


class TestRebuildWithChangedPeriods:
    """Modify WP active period, rebuild, verify inactive months."""

    def test_changed_start_month_greys_out_new_inactive(self, init_path):
        """Changing a WP start to a later month makes earlier months inactive."""
        wb = openpyxl.load_workbook(init_path)
        _, wps = read_config_sheet(wb["Config"])
        sched_entries = _find_sched_rows(wb["Config"])

        # Find a schedule entry that starts at M1 and push its start to M7
        target_entry = next((e for e in sched_entries if e["start"] == 1), None)
        if target_entry is None:
            pytest.skip("No schedule entry starts at M1 to modify")

        # Map schedule WP ID to budget WP index
        target_idx = next(i for i, wp in enumerate(wps) if wp["id"] == target_entry["id"])

        from budget_planner.layout import COL_WP_START

        new_start = 7
        _set(wb["Config"], target_entry["row"], COL_WP_START, new_start)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        # Verify Plan inactive months are empty
        wb = openpyxl.load_workbook(init_path)
        plan = wb["Plan"]
        layout = discover_layout(plan)
        row = layout["planned_first_wp_row"] + target_idx
        for m in range(new_start - 1):  # months 1 to new_start-1
            val = plan.cell(row, MONTH_COL_START + m).value
            assert val is None or val == "" or val == 0, f"M{m + 1} should be inactive after start change, got {val}"
        wb.close()

    def test_changed_end_month_greys_out_new_inactive(self, init_path):
        """Shortening a WP end month makes later months inactive."""
        wb = openpyxl.load_workbook(init_path)
        params, wps = read_config_sheet(wb["Config"])
        sched_entries = _find_sched_rows(wb["Config"])

        # Find a schedule entry that ends at the project end and shorten it
        target_entry = next(
            (e for e in sched_entries if e["end"] == params["months"]),
            None,
        )
        if target_entry is None:
            pytest.skip("No schedule entry ends at project end to modify")

        # Map schedule WP ID to budget WP index
        target_idx = next(i for i, wp in enumerate(wps) if wp["id"] == target_entry["id"])

        from budget_planner.layout import COL_WP_END

        new_end = params["months"] * 3 // 4
        _set(wb["Config"], target_entry["row"], COL_WP_END, new_end)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        plan = wb["Plan"]
        layout = discover_layout(plan)
        row = layout["planned_first_wp_row"] + target_idx
        # Month after new_end should be inactive
        for m in range(new_end, params["months"]):
            val = plan.cell(row, MONTH_COL_START + m).value
            assert val is None or val == "" or val == 0, f"M{m + 1} should be inactive after end change, got {val}"
        wb.close()


# ── Test: Change team member names ──


class TestRebuildWithChangedTeam:
    """Modify team member names in Config roster, rebuild, verify."""

    @staticmethod
    def _find_person_without_externals(roster, config_ws):
        """Find a team member who has no external commitments.

        Args:
            roster: list of dicts with 'name', 'max_fte', 'row' keys.
            config_ws: Config worksheet (externals live in the EXTERNAL
                COMMITMENTS section on Config).
        """
        ext_persons = set()
        for r in range(1, config_ws.max_row + 1):
            val = config_ws.cell(r, 1).value
            if val and "EXTERNAL COMMITMENTS" in str(val).upper():
                # Skip section header + column header rows
                dr = r + 2
                while dr <= config_ws.max_row:
                    person = config_ws.cell(dr, 1).value
                    if person is None or str(person).strip() == "":
                        break
                    ext_persons.add(str(person))
                    dr += 1
                break
        for p in roster:
            if p["name"] not in ext_persons:
                return p
        return None

    @staticmethod
    def _read_roster_with_rows(config_ws):
        """Read team roster from Config, including row numbers."""
        roster = read_team_from_config(config_ws)
        # Find the roster rows by scanning for TEAM ROSTER header
        from budget_planner.reader import _find_header_row

        section_row = _find_header_row(config_ws, "TEAM ROSTER")
        # Column header is always the next row after section header
        person_header = section_row + 1
        first_row = person_header + 1
        for i, member in enumerate(roster):
            member["row"] = first_row + i
        return roster

    def test_renamed_person_appears_in_capacity(self, init_path):
        """Renaming a person in Config roster creates correct Capacity group."""
        wb = openpyxl.load_workbook(init_path)
        roster = self._read_roster_with_rows(wb["Config"])
        person = self._find_person_without_externals(roster, wb["Config"])
        if person is None:
            wb.close()
            pytest.skip("All team members have external commitments")
        old_name = person["name"]
        new_name = "NewPerson"
        _set(wb["Config"], person["row"], 1, new_name)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        plan = wb["Plan"]
        layout = discover_layout(plan)
        config_wps = read_plan_config(plan, layout)
        cap_data = read_capacity_sheet(wb["Capacity"], config_wps)
        assert new_name in cap_data["person_groups"], f"Renamed person '{new_name}' not found in Capacity"
        assert old_name not in cap_data["person_groups"], f"Old name '{old_name}' still in Capacity after rename"
        wb.close()

    def test_renamed_person_capacity_has_no_wps(self, init_path):
        """Renamed person has no WP rows (old-name assignments are orphaned)."""
        wb = openpyxl.load_workbook(init_path)
        roster = self._read_roster_with_rows(wb["Config"])
        person = self._find_person_without_externals(roster, wb["Config"])
        if person is None:
            wb.close()
            pytest.skip("All team members have external commitments")

        new_name = "RenamedPerson"
        _set(wb["Config"], person["row"], 1, new_name)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        plan = wb["Plan"]
        layout = discover_layout(plan)
        config_wps = read_plan_config(plan, layout)
        cap_data = read_capacity_sheet(wb["Capacity"], config_wps)
        group = cap_data["person_groups"][new_name]
        # Assignments reference old name → orphaned → no WP rows for renamed person
        assert set(group["wp_values"].keys()) == set()
        wb.close()


# ── Test: Change team allocations/shares ──


class TestRebuildWithChangedAllocations:
    """Modify Capacity FTE values, rebuild, verify preservation."""

    def test_modified_assignment_total_fte_preserved_after_rebuild(self, assigned_init_path):
        """Modified total_fte in Config assignment section survives rebuild."""
        wb = openpyxl.load_workbook(assigned_init_path)
        config_ws = wb["Config"]
        original = read_assignments_from_config(config_ws)
        assert len(original) > 0, "Need at least one assignment"

        # Find the assignment section row on Config and modify total_fte
        for r in range(1, config_ws.max_row + 1):
            val = config_ws.cell(r, 1).value
            if isinstance(val, str) and val == original[0]["person"]:
                wp_val = config_ws.cell(r, 2).value
                if isinstance(wp_val, str) and wp_val == original[0]["wp"]:
                    new_total_fte = 99.99
                    _set(config_ws, r, 3, new_total_fte)
                    break
        else:
            pytest.fail("Could not find assignment section row on Config")

        wb.save(assigned_init_path)
        wb.close()

        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        rebuilt = read_assignments_from_config(wb["Config"])
        first = next(a for a in rebuilt if a["person"] == original[0]["person"] and a["wp"] == original[0]["wp"])
        assert first["total_fte"] == pytest.approx(new_total_fte)
        wb.close()


# ── Test: Insert new rows in config blocks, rebuild ──


class TestRebuildWithInsertedRows:
    """User adds new rows in Excel config blocks, then rebuilds."""

    def test_added_assignment_row_creates_wp_row_after_rebuild(self, assigned_init_path):
        """New assignment row added in Config section -> person gets new WP row."""
        wb = openpyxl.load_workbook(assigned_init_path)
        config_ws = wb["Config"]
        original = read_assignments_from_config(config_ws)

        # Find the last assignment row on Config and add a new one below it
        last_asgn_row = None
        for r in range(1, config_ws.max_row + 1):
            person = config_ws.cell(r, 1).value
            wp = config_ws.cell(r, 2).value
            if person and wp and any(a["person"] == str(person) and a["wp"] == str(wp) for a in original):
                last_asgn_row = r

        assert last_asgn_row is not None

        # Pick a person and WP that aren't already assigned together — read
        # from the workbook, not config.py, so tests survive seed data changes.
        target_person = original[0]["person"]
        _, all_wps = read_config_sheet(config_ws)
        assigned_wps = {a["wp"] for a in original if a["person"] == target_person}
        new_wp = next((wp["id"] for wp in all_wps if wp["id"] not in assigned_wps), None)
        if new_wp is None:
            pytest.skip("All WPs already assigned to first person")

        # Insert a new row (like a user would in Excel)
        new_row = last_asgn_row + 1
        config_ws.insert_rows(new_row)
        _set(config_ws, new_row, 1, target_person)
        _set(config_ws, new_row, 2, new_wp)
        _set(config_ws, new_row, 3, 2.0)  # total FTE
        _set(config_ws, new_row, 4, 1)  # start month
        _set(config_ws, new_row, 5, 12)  # end month
        wb.save(assigned_init_path)
        wb.close()

        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        _, wps = read_config_sheet(wb["Config"])
        cap_data = read_capacity_sheet(wb["Capacity"], wps)
        person_wps = cap_data["person_groups"][target_person]["wp_values"]
        assert new_wp in person_wps, f"{target_person} should have WP row for {new_wp} after rebuild"
        wb.close()

    def test_added_external_commitment_row_appears_after_rebuild(self, assigned_init_path):
        """New external commitment added in Config section -> person gets external row."""
        wb = openpyxl.load_workbook(assigned_init_path)
        config_ws = wb["Config"]

        # Find EXTERNAL COMMITMENTS section on Config
        ext_header = None
        for r in range(1, config_ws.max_row + 1):
            val = config_ws.cell(r, 1).value
            if val and "EXTERNAL COMMITMENTS" in str(val).upper():
                ext_header = r
                break
        assert ext_header is not None

        # Find last data row in the section
        r = ext_header + 2  # skip section header + column headers
        while r <= config_ws.max_row and config_ws.cell(r, 1).value:
            r += 1
        insert_row = r  # first empty row after existing data

        # Read from workbook, not config.py, so tests survive seed data changes.
        all_people = read_team_from_config(config_ws)
        new_person = all_people[-1]["name"]
        new_project = "NewExtProject"
        # Insert a new row to preserve the blank separator between sections
        config_ws.insert_rows(insert_row)
        _set(config_ws, insert_row, 1, new_person)
        _set(config_ws, insert_row, 2, new_project)
        _set(config_ws, insert_row, 3, 2.0)
        _set(config_ws, insert_row, 4, 1)
        _set(config_ws, insert_row, 5, 12)
        wb.save(assigned_init_path)
        wb.close()

        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        _, wps = read_config_sheet(wb["Config"])
        cap_data = read_capacity_sheet(wb["Capacity"], wps)
        person_ext = cap_data["person_groups"][new_person]["external_values"]
        assert new_project in person_ext, f"{new_person} should have external row for {new_project} after rebuild"
        wb.close()

    def test_added_team_member_gets_capacity_group_after_rebuild(self, init_path):
        """New person added to Config TEAM ROSTER -> gets Capacity person group."""
        wb = openpyxl.load_workbook(init_path)
        config_ws = wb["Config"]

        # Find last roster data row
        roster_section = None
        for r in range(1, config_ws.max_row + 1):
            if config_ws.cell(r, 1).value and "TEAM ROSTER" in str(config_ws.cell(r, 1).value).upper():
                roster_section = r
                break
        assert roster_section is not None

        # Scan to end of roster data
        r = roster_section + 2  # skip section header + column headers
        while r <= config_ws.max_row and config_ws.cell(r, 1).value:
            r += 1
        insert_row = r

        _set(config_ws, insert_row, 1, "Eve")
        _set(config_ws, insert_row, 2, 0.5)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        _, wps = read_config_sheet(wb["Config"])
        cap_data = read_capacity_sheet(wb["Capacity"], wps)
        assert "Eve" in cap_data["person_groups"], "'Eve' should have a Capacity person group after rebuild"
        # Verify Eve appears in rebuilt roster
        roster = read_team_from_config(wb["Config"])
        names = {p["name"] for p in roster}
        assert "Eve" in names
        wb.close()


# ── Test: Cross-sheet formula integrity after rebuild ──


class TestRebuildFormulaIntegrity:
    """After rebuild, Plan formulas reference correct Capacity rows."""

    def test_plan_formulas_reference_capacity_after_rebuild(self, assigned_init_path):
        """Plan Planned FTE formulas reference Capacity sheet after rebuild."""
        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        plan = wb["Plan"]
        layout = discover_layout(plan)
        config_wps = read_plan_config(plan, layout)
        cap_data = read_capacity_sheet(wb["Capacity"], config_wps)
        _, cfg_wps_ranges = read_config_sheet(wb["Config"])
        from budget_planner.layout import wp_month_bounds

        for i, wp in enumerate(config_wps):
            row = layout["planned_first_wp_row"] + i
            wp_start, _ = wp_month_bounds(wp["id"], cfg_wps_ranges)
            first_active_col = MONTH_COL_START + wp_start - 1
            cell_val = plan.cell(row, first_active_col).value

            # WPs without any person assigned have no Capacity refs
            assigned_count = sum(1 for pg in cap_data["person_groups"].values() if wp["id"] in pg["wp_values"])
            if assigned_count == 0:
                continue

            assert cell_val is not None, f"Planned cell for {wp['id']} is empty"
            assert isinstance(cell_val, str) and "Capacity!" in cell_val, (
                f"Planned formula for {wp['id']} should reference Capacity"
            )

            # Count person refs in formula should match assigned persons
            cap_refs = re.findall(r"Capacity!\w+(\d+)", cell_val)
            assert len(cap_refs) == assigned_count, (
                f"{wp['id']}: formula refs {len(cap_refs)} rows, but {assigned_count} persons assigned"
            )
        wb.close()

    def test_actuals_has_row_per_wp_after_rebuild(self, init_path):
        """Actuals has one row per WP after rebuild."""
        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        _, wps = read_config_sheet(wb["Config"])
        first_wp = _find_first_wp_row_actuals(wb["Actuals"])
        total_row = _find_total_row(wb["Actuals"])
        assert first_wp is not None and total_row is not None
        wp_count = total_row - first_wp
        assert wp_count == len(wps)
        wb.close()

    def test_summary_has_row_per_wp_after_rebuild(self, init_path):
        """Summary has one formula row per WP after rebuild."""
        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        _, wps = read_config_sheet(wb["Config"])
        summary = wb["Summary"]
        wp_formula_rows = 0
        for r in range(1, summary.max_row + 1):
            val = summary.cell(r, 1).value
            if isinstance(val, str) and val.startswith("=Config!A"):
                wp_formula_rows += 1
        # Budget Overview + RP Summary each have one formula row per WP
        assert wp_formula_rows == len(wps) * 2
        wb.close()


# ── Test: Full scenario — multiple changes at once ──


class TestRebuildFullScenario:
    """Combine FTE, period, name, and allocation changes in one rebuild."""

    def test_multiple_changes_rebuild_produces_valid_workbook(self, init_path):
        """Apply multiple data changes, rebuild, verify structural integrity."""
        wb = openpyxl.load_workbook(init_path)
        params, wps = read_config_sheet(wb["Config"])
        sched_entries = _find_sched_rows(wb["Config"])

        # 1. Change first WP FTE
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, 99.0)

        # 2. Shorten second WP end month via WP BUDGET End col (I)
        from budget_planner.layout import COL_WP_END

        wp2_sched = next(e for e in sched_entries if e["id"] == wps[1]["id"])
        original_end = wp2_sched["end"]
        shortened_end = max(wp2_sched["start"], original_end - 6)
        _set(wb["Config"], wp2_sched["row"], COL_WP_END, shortened_end)

        # 3. Rename a team member without external commitments
        roster = TestRebuildWithChangedTeam._read_roster_with_rows(wb["Config"])
        person = TestRebuildWithChangedTeam._find_person_without_externals(roster, wb["Config"])
        assert person is not None, "Need at least one person without externals"
        _set(wb["Config"], person["row"], 1, "Zara")

        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        # Verify everything rebuilt correctly
        wb = openpyxl.load_workbook(init_path)

        # Config round-trip
        rebuilt_params, rebuilt_wps = read_config_sheet(wb["Config"])
        assert rebuilt_wps[0]["fte"] == pytest.approx(99.0)
        wp2_rebuilt = next(w for w in rebuilt_wps if w["id"] == wps[1]["id"])
        assert wp2_rebuilt["end"] == shortened_end
        # Untouched project params survive the rebuild (regression gate for
        # builders that might clobber title/months on a mixed-edit rebuild).
        assert rebuilt_params["title"] == params["title"]
        assert rebuilt_params["months"] == params["months"]

        # Team name round-trip
        rebuilt_roster = read_team_from_config(wb["Config"])
        rebuilt_names = {p["name"] for p in rebuilt_roster}
        assert "Zara" in rebuilt_names, f"'Zara' not in rebuilt roster: {rebuilt_names}"

        # Capacity has renamed person
        plan = wb["Plan"]
        layout = discover_layout(plan)
        config_wps = read_plan_config(plan, layout)
        cap_data = read_capacity_sheet(wb["Capacity"], config_wps)
        assert "Zara" in cap_data["person_groups"]

        # All sheets present
        for name in ["Config", "Plan", "Capacity", "Actuals", "Summary", "Trends"]:
            assert name in wb.sheetnames

        # WP count consistent across sheets
        num_wps = len(rebuilt_wps)
        first_wp = _find_first_wp_row_actuals(wb["Actuals"])
        total_row = _find_total_row(wb["Actuals"])
        assert first_wp is not None, "Could not find first WP row in Actuals"
        assert total_row is not None, "Could not find TOTAL row in Actuals"
        assert total_row - first_wp == num_wps

        wb.close()


# ── Test: Recalc after rebuild with changed data ──


class TestRebuildRecalc:
    """After data changes + rebuild, LibreOffice recalc produces zero errors."""

    @pytest.fixture(scope="class")
    def recalc_state(self, tmp_path_factory):
        """Build workbook with changed FTE, rebuild, recalc once for the class."""
        path = str(tmp_path_factory.mktemp("rebuild_recalc") / "test.xlsx")
        init_workbook(output_path=path)

        wb = openpyxl.load_workbook(path)
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, 50.0)
        wb.save(path)
        wb.close()

        rebuild_workbook(input_path=path)

        result = subprocess.run(
            ["python3", str(SCRIPTS_DIR / "recalc.py"), path],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            pytest.skip("LibreOffice not available")
        return {"path": path, "recalc": json.loads(result.stdout)}

    def test_recalc_zero_errors_after_fte_change(self, recalc_state):
        """Changed FTE + rebuild + recalc = no formula errors."""
        data = recalc_state["recalc"]
        assert data["status"] == "success"
        assert data["total_errors"] == 0

    def test_recalc_partner_ceiling_matches_config_budget_minus_i(self, recalc_state):
        """partner_ceiling (Plan!G1) equals SUM of the Config Budget -i column
        after recalc. Under the indirect-view restructure (D4), Budget +i is
        authoritative blue-input and Budget -i is the derived ``/ 1.25`` cell;
        plan.py still aggregates col E for backwards compat with deferred
        consumers, so partner_ceiling tracks Config Budget -i, not live FTE."""
        wb = openpyxl.load_workbook(recalc_state["path"], data_only=True)
        config = wb["Config"]
        _, wps = read_config_sheet(config)

        from budget_planner.reader import _find_header_row

        wp_section = _find_header_row(config, "WP BUDGET")
        first_wp = wp_section + 2
        expected_ceiling = 0.0
        for i in range(len(wps)):
            r = first_wp + i
            minus_i = config.cell(r, 5).value  # Budget -i column
            assert isinstance(minus_i, (int, float)), f"WP row {r}: Budget -i={minus_i} — expected numeric after recalc"
            expected_ceiling += minus_i

        plan = wb["Plan"]
        ceiling_val = plan.cell(1, 7).value
        assert isinstance(ceiling_val, (int, float)), (
            f"partner_ceiling should be numeric after recalc, got {ceiling_val}"
        )
        assert abs(ceiling_val - expected_ceiling) < 0.01, (
            f"partner_ceiling {ceiling_val} != SUM(Budget -i) {expected_ceiling}"
        )
        wb.close()


# ── Test: Reporting period round-trip through rebuild ──


def test_reporting_periods_survive_rebuild(tmp_path):
    """Reporting periods in Config survive a rebuild cycle."""
    from budget_planner.workbook import init_workbook, rebuild_workbook
    from budget_planner.reader import read_reporting_periods
    from budget_planner.config import REPORTING_PERIODS
    from openpyxl import load_workbook

    out = str(tmp_path / "rp_rebuild.xlsx")
    init_workbook(out)

    # Read original RPs
    wb = load_workbook(out)
    original_rps = read_reporting_periods(wb["Config"])
    wb.close()
    assert len(original_rps) == len(REPORTING_PERIODS)

    # Rebuild
    rebuild_workbook(out)

    # Read rebuilt RPs
    wb = load_workbook(out)
    rebuilt_rps = read_reporting_periods(wb["Config"])
    wb.close()

    assert len(rebuilt_rps) == len(original_rps)
    for orig, rebuilt in zip(original_rps, rebuilt_rps):
        assert orig["id"] == rebuilt["id"]
        assert orig["start"] == rebuilt["start"]
        assert orig["end"] == rebuilt["end"]


def test_modified_reporting_periods_survive_rebuild(tmp_path):
    """User-modified reporting periods are preserved on rebuild."""
    from budget_planner.workbook import init_workbook, rebuild_workbook
    from budget_planner.reader import read_reporting_periods
    from openpyxl import load_workbook

    out = str(tmp_path / "rp_mod.xlsx")
    init_workbook(out)

    # Modify RP boundaries in the workbook
    wb = load_workbook(out)
    ws = wb["Config"]

    # Find REPORTING PERIODS section
    rp_section = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIODS" in str(val).upper():
            rp_section = r
            break
    assert rp_section is not None

    # Change RP1 end from 18 to 12
    data_row = rp_section + 2  # skip header + col headers
    _set(ws, data_row, 3, 12)  # End Month = 12

    wb.save(out)
    wb.close()

    # Rebuild
    rebuild_workbook(out)

    # Verify modification preserved
    wb = load_workbook(out)
    rps = read_reporting_periods(wb["Config"])
    wb.close()

    assert rps[0]["end"] == 12  # Modified value preserved
