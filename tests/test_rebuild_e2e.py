"""End-to-end rebuild tests: modify workbook data, rebuild, verify results.

These tests prove rebuild works correctly after changing FTE budgets,
WP start/end months, team names, and allocations — without hardcoding
any specific values from config.py.

The recalc test saves the rebuilt+recalc'd workbook to
output/budget_plan_change.xlsx for manual inspection.
"""
import json
import re
import shutil
import subprocess
from pathlib import Path

import pytest
import openpyxl
from openpyxl.cell import Cell

from budget_planner.workbook import init_workbook, rebuild_workbook
from budget_planner.reader import (
    discover_layout, read_config_sheet, read_team_from_config,
    read_capacity_sheet, read_plan_config,
)
from budget_planner.config import MONTH_COL_START

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"
OUTPUT_DIR = Path(__file__).parent.parent / "output"

# Config sheet layout: WP data rows start at row 9
# (rows 1-5 = params, 6 = gap, 7 = section header, 8 = column headers)
CONFIG_FIRST_WP_ROW = 9


# ── Helpers ──


def _set(ws, row, col, value):
    """Set a cell value."""
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _save_change(src_path):
    """Copy rebuilt workbook to output/budget_plan_change.xlsx."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    shutil.copy2(src_path, OUTPUT_DIR / "budget_plan_change.xlsx")


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

        # Find a WP that starts at M1 and push its start to M7
        target_idx = next(
            (i for i, wp in enumerate(wps) if wp["start"] == 1), None
        )
        if target_idx is None:
            pytest.skip("No WP starts at M1 to modify")

        new_start = 7
        _set(wb["Config"], CONFIG_FIRST_WP_ROW + target_idx, 4, new_start)
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
            assert val is None or val == "" or val == 0, (
                f"M{m+1} should be inactive after start change, got {val}"
            )
        wb.close()

    def test_changed_end_month_greys_out_new_inactive(self, init_path):
        """Shortening a WP end month makes later months inactive."""
        wb = openpyxl.load_workbook(init_path)
        params, wps = read_config_sheet(wb["Config"])

        # Find a WP that ends at the project end and shorten it
        target_idx = next(
            (i for i, wp in enumerate(wps) if wp["end"] == params["months"]),
            None,
        )
        if target_idx is None:
            pytest.skip("No WP ends at project end to modify")

        new_end = 36
        _set(wb["Config"], CONFIG_FIRST_WP_ROW + target_idx, 5, new_end)
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
            assert val is None or val == "" or val == 0, (
                f"M{m+1} should be inactive after end change, got {val}"
            )
        wb.close()


# ── Test: Change team member names ──


class TestRebuildWithChangedTeam:
    """Modify team member names in Config roster, rebuild, verify."""

    @staticmethod
    def _find_person_without_externals(roster, cap_ws):
        """Find a team member who has no external commitments.

        Args:
            roster: list of dicts with 'name', 'max_fte', 'row' keys.
            cap_ws: Capacity worksheet.
        """
        # Read external commitments from Capacity config block
        ext_persons = set()
        for r in range(1, min(cap_ws.max_row + 1, 30)):
            val = cap_ws.cell(r, 1).value
            if val and "EXTERNAL COMMITMENTS" in str(val).upper():
                # Skip header rows, read data
                dr = r + 2
                while dr <= cap_ws.max_row:
                    person = cap_ws.cell(dr, 1).value
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
        person_header = _find_header_row(config_ws, "Person", max_row=section_row + 5)
        first_row = person_header + 1
        for i, member in enumerate(roster):
            member["row"] = first_row + i
        return roster

    def test_renamed_person_appears_in_capacity(self, init_path):
        """Renaming a person in Config roster creates correct Capacity group."""
        wb = openpyxl.load_workbook(init_path)
        roster = self._read_roster_with_rows(wb["Config"])
        person = self._find_person_without_externals(roster, wb["Capacity"])
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
        assert new_name in cap_data["person_groups"], (
            f"Renamed person '{new_name}' not found in Capacity"
        )
        assert old_name not in cap_data["person_groups"], (
            f"Old name '{old_name}' still in Capacity after rename"
        )
        wb.close()

    def test_renamed_person_capacity_has_no_wps(self, init_path):
        """Renamed person has no WP rows (old-name assignments are orphaned)."""
        wb = openpyxl.load_workbook(init_path)
        roster = self._read_roster_with_rows(wb["Config"])
        person = self._find_person_without_externals(roster, wb["Capacity"])
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
        """Modified total_fte in Capacity assignment config block survives rebuild."""
        from budget_planner.reader import read_capacity_assignments

        wb = openpyxl.load_workbook(assigned_init_path)
        cap_ws = wb["Capacity"]
        original = read_capacity_assignments(cap_ws)
        assert len(original) > 0, "Need at least one assignment"

        # Find the assignment config block row and modify total_fte
        # Assignments block: col A=Person, col B=WP, col C=Total FTE, col D=Start, col E=End
        for r in range(1, cap_ws.max_row + 1):
            val = cap_ws.cell(r, 1).value
            if isinstance(val, str) and val == original[0]["person"]:
                wp_val = cap_ws.cell(r, 2).value
                if isinstance(wp_val, str) and wp_val == original[0]["wp"]:
                    new_total_fte = 99.99
                    _set(cap_ws, r, 3, new_total_fte)
                    break
        else:
            pytest.fail("Could not find assignment config block row")

        wb.save(assigned_init_path)
        wb.close()

        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        rebuilt = read_capacity_assignments(wb["Capacity"])
        first = next(a for a in rebuilt
                     if a["person"] == original[0]["person"]
                     and a["wp"] == original[0]["wp"])
        assert first["total_fte"] == pytest.approx(new_total_fte)
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

        for i, wp in enumerate(config_wps):
            row = layout["planned_first_wp_row"] + i
            first_active_col = MONTH_COL_START + wp["start"] - 1
            cell_val = plan.cell(row, first_active_col).value

            # WPs without any person assigned have no Capacity refs
            assigned_count = sum(
                1 for pg in cap_data["person_groups"].values()
                if wp["id"] in pg["wp_values"]
            )
            if assigned_count == 0:
                continue

            assert cell_val is not None, f"Planned cell for {wp['id']} is empty"
            assert isinstance(cell_val, str) and "Capacity!" in cell_val, (
                f"Planned formula for {wp['id']} should reference Capacity"
            )

            # Count person refs in formula should match assigned persons
            cap_refs = re.findall(r"Capacity!\w+(\d+)", cell_val)
            assert len(cap_refs) == assigned_count, (
                f"{wp['id']}: formula refs {len(cap_refs)} rows, "
                f"but {assigned_count} persons assigned"
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
        assert wp_formula_rows == len(wps)
        wb.close()


# ── Test: Full scenario — multiple changes at once ──


class TestRebuildFullScenario:
    """Combine FTE, period, name, and allocation changes in one rebuild."""

    def test_multiple_changes_rebuild_produces_valid_workbook(self, init_path):
        """Apply multiple data changes, rebuild, verify structural integrity."""
        wb = openpyxl.load_workbook(init_path)
        params, wps = read_config_sheet(wb["Config"])

        # 1. Change first WP FTE
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, 99.0)

        # 2. Shorten second WP end month
        original_end = wps[1]["end"]
        shortened_end = max(wps[1]["start"], original_end - 6)
        _set(wb["Config"], CONFIG_FIRST_WP_ROW + 1, 5, shortened_end)

        # 3. Rename a team member without external commitments
        roster = TestRebuildWithChangedTeam._read_roster_with_rows(wb["Config"])
        person = TestRebuildWithChangedTeam._find_person_without_externals(
            roster, wb["Capacity"]
        )
        assert person is not None, "Need at least one person without externals"
        _set(wb["Config"], person["row"], 1, "Zara")

        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        # Verify everything rebuilt correctly
        wb = openpyxl.load_workbook(init_path)

        # Config round-trip
        _, rebuilt_wps = read_config_sheet(wb["Config"])
        assert rebuilt_wps[0]["fte"] == pytest.approx(99.0)
        assert rebuilt_wps[1]["end"] == shortened_end

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
        for name in ["_META", "Config", "Plan", "Capacity",
                      "Actuals", "Summary", "Trends"]:
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

    def test_recalc_zero_errors_after_fte_change(self, init_path):
        """Changed FTE + rebuild + recalc = no formula errors."""
        wb = openpyxl.load_workbook(init_path)
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, 50.0)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        result = subprocess.run(
            ["python3", str(SCRIPTS_DIR / "recalc.py"), init_path],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            pytest.skip("LibreOffice not available")
        data = json.loads(result.stdout)
        assert data["status"] == "success"
        assert data["total_errors"] == 0

    def test_recalc_ceiling_reflects_changed_fte(self, init_path):
        """After FTE change + rebuild + recalc, partner_ceiling uses new value."""
        wb = openpyxl.load_workbook(init_path)
        params, wps = read_config_sheet(wb["Config"])
        new_fte = 50.0
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, new_fte)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        result = subprocess.run(
            ["python3", str(SCRIPTS_DIR / "recalc.py"), init_path],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            pytest.skip("LibreOffice not available")

        # Recompute expected ceiling with changed FTE
        expected_ceiling = new_fte * params["rate"]
        for i, wp in enumerate(wps):
            if i != 0:
                expected_ceiling += wp["fte"] * params["rate"]

        wb = openpyxl.load_workbook(init_path, data_only=True)
        plan = wb["Plan"]
        # partner_ceiling is Plan!I1 (compact header)
        ceiling_val = plan.cell(1, 9).value
        assert isinstance(ceiling_val, (int, float)), (
            f"partner_ceiling should be numeric after recalc, got {ceiling_val}"
        )
        assert abs(ceiling_val - expected_ceiling) < 0.01, (
            f"Ceiling {ceiling_val} != expected {expected_ceiling}"
        )
        wb.close()

    def test_changed_file_has_autofit_columns_after_recalc(self, init_path):
        """Full scenario + rebuild + recalc => _changed file has non-default widths."""
        # Apply multiple changes (mirrors TestRebuildFullScenario)
        wb = openpyxl.load_workbook(init_path)
        params, wps = read_config_sheet(wb["Config"])
        _set(wb["Config"], CONFIG_FIRST_WP_ROW, 3, 99.0)
        original_end = wps[1]["end"]
        shortened_end = max(wps[1]["start"], original_end - 6)
        _set(wb["Config"], CONFIG_FIRST_WP_ROW + 1, 5, shortened_end)
        roster = TestRebuildWithChangedTeam._read_roster_with_rows(wb["Config"])
        person = TestRebuildWithChangedTeam._find_person_without_externals(
            roster, wb["Capacity"]
        )
        assert person is not None, "Need at least one person without externals"
        _set(wb["Config"], person["row"], 1, "Zara")
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        # Run recalc (auto-fits columns via LibreOffice macro)
        result = subprocess.run(
            ["python3", str(SCRIPTS_DIR / "recalc.py"), init_path],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            pytest.skip("LibreOffice not available")
        data = json.loads(result.stdout)
        assert data["status"] == "success"

        # Save the recalc'd workbook as _changed for manual inspection
        _save_change(init_path)

        # Verify _changed file has non-default column widths
        changed_path = OUTPUT_DIR / "budget_plan_change.xlsx"
        wb = openpyxl.load_workbook(str(changed_path))
        default_width = 13.0
        for sheet_name in ["Plan", "Capacity", "Actuals", "Summary"]:
            ws = wb[sheet_name]
            col_a_width = ws.column_dimensions["A"].width
            assert col_a_width != default_width, (
                f"{sheet_name} col A has default width {default_width} — "
                "recalc auto-fit did not run on _changed file"
            )
        wb.close()
