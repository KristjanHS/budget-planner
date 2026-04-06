"""Tests for rebuild round-trip, preservation, orphaned data, and validation."""
import os
from datetime import datetime

import pytest
from openpyxl import load_workbook
from openpyxl.cell import Cell

from budget_planner.workbook import init_workbook, rebuild_workbook, _validate_config


def _set_cell(ws: object, row: int, col: int, value: float | int) -> None:
    """Set a cell value with pyright-safe type narrowing."""
    cell = ws.cell(row, col)  # type: ignore[union-attr]
    assert isinstance(cell, Cell)
    cell.value = value


@pytest.fixture
def init_wb(tmp_path):
    """Create a fresh workbook for rebuild testing."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    return path


class TestRebuildRoundTrip:
    """Rebuild preserves structure and data."""

    def test_rebuild_produces_all_sheets(self, init_wb):
        """Rebuild creates the same sheets as init."""
        rebuild_workbook(input_path=init_wb)
        wb = load_workbook(init_wb)
        for name in ["_META", "Config", "Plan", "Capacity",
                      "Actuals", "Summary", "Trends"]:
            assert name in wb.sheetnames, f"Missing sheet: {name}"
        wb.close()

    def test_rebuild_preserves_config_params(self, init_wb):
        """Project params survive a rebuild."""
        from budget_planner.reader import read_config_sheet
        wb = load_workbook(init_wb)
        original_params, _ = read_config_sheet(wb["Config"])
        wb.close()

        rebuild_workbook(input_path=init_wb)

        wb = load_workbook(init_wb)
        rebuilt_params, _ = read_config_sheet(wb["Config"])
        assert rebuilt_params["title"] == original_params["title"]
        assert rebuilt_params["months"] == original_params["months"]
        wb.close()

    def test_rebuild_preserves_wp_count(self, init_wb):
        """Same number of WPs after rebuild."""
        from budget_planner.reader import read_config_sheet
        wb = load_workbook(init_wb)
        _, original_wps = read_config_sheet(wb["Config"])
        wp_count = len(original_wps)
        wb.close()

        rebuild_workbook(input_path=init_wb)

        wb = load_workbook(init_wb)
        _, rebuilt_wps = read_config_sheet(wb["Config"])
        assert len(rebuilt_wps) == wp_count
        wb.close()


class TestCapacityPreservation:
    """Capacity assignment config block values are preserved on rebuild."""

    @pytest.fixture
    def assigned_wb(self, tmp_path):
        """Create a workbook with demo assignments for rebuild testing."""
        from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
        path = str(tmp_path / "assigned.xlsx")
        init_workbook(output_path=path, assignments=DEMO_CAPACITY_ASSIGNMENTS)
        return path

    def test_assignment_config_block_preserved(self, assigned_wb):
        """Person-WP assignment config block values survive rebuild."""
        from budget_planner.reader import read_capacity_assignments
        from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS

        wb = load_workbook(assigned_wb)
        original = read_capacity_assignments(wb["Capacity"])
        assert len(original) == len(DEMO_CAPACITY_ASSIGNMENTS)
        wb.close()

        rebuild_workbook(input_path=assigned_wb)

        wb = load_workbook(assigned_wb)
        rebuilt = read_capacity_assignments(wb["Capacity"])
        assert len(rebuilt) == len(original)
        for got, expected in zip(rebuilt, original):
            assert got["person"] == expected["person"]
            assert got["wp"] == expected["wp"]
            assert got["total_fte"] == pytest.approx(expected["total_fte"])
            assert got["start"] == expected["start"]
            assert got["end"] == expected["end"]
        wb.close()

    def test_rebuild_keeps_structure(self, init_wb):
        """Rebuild produces valid workbook structure."""
        rebuild_workbook(input_path=init_wb)

        wb = load_workbook(init_wb)
        assert "Config" in wb.sheetnames
        assert "Capacity" in wb.sheetnames
        assert "Actuals" in wb.sheetnames
        wb.close()


class TestActualsPreservation:
    """Actuals EUR values are always preserved on rebuild."""

    def test_actuals_values_preserved(self, init_wb):
        """Modified Actuals EUR values survive rebuild."""
        # Find the first WP data row in Actuals (row 4 per actuals.py layout)
        wb = load_workbook(init_wb)
        act = wb["Actuals"]

        # Actuals WP rows start at row 4, first month col is C (3)
        # Find the first active month cell (non-grey, non-empty formula)
        test_row = 4
        test_col = 3
        _set_cell(act, test_row, test_col, 5000)
        wb.save(init_wb)
        wb.close()

        rebuild_workbook(input_path=init_wb)

        wb = load_workbook(init_wb)
        act = wb["Actuals"]
        # Find the cell with value 5000 (row may shift)
        found = False
        for r in range(1, act.max_row + 1):
            val = act.cell(r, test_col).value
            if val == 5000:
                found = True
                break
        assert found, "Actuals value 5000 not found after rebuild"
        wb.close()


class TestValidation:
    """_validate_config catches bad data."""

    def _params(self, **overrides):
        p = {"title": "Test", "start": datetime(2026, 7, 1),
             "months": 60, "rate": 7500}
        p.update(overrides)
        return p

    def _wps(self):
        return [{"id": "WP1", "name": "Mgmt", "fte": 12.0,
                 "start": 1, "end": 60}]

    def _team(self):
        return [{"name": "Alice", "max_fte": 1.0}]

    def test_valid_config_no_errors(self):
        errors, warnings = _validate_config(
            self._params(), self._wps(), self._team(), [])
        assert len(errors) == 0

    def test_months_out_of_range(self):
        errors, _ = _validate_config(
            self._params(months=0), self._wps(), self._team(), [])
        assert any("PROJECT_MONTHS" in e for e in errors)

    def test_months_too_high(self):
        errors, _ = _validate_config(
            self._params(months=200), self._wps(), self._team(), [])
        assert any("PROJECT_MONTHS" in e for e in errors)

    def test_rate_zero(self):
        errors, _ = _validate_config(
            self._params(rate=0), self._wps(), self._team(), [])
        assert any("FTE rate" in e for e in errors)

    def test_empty_wps(self):
        errors, _ = _validate_config(
            self._params(), [], self._team(), [])
        assert any("empty" in e.lower() for e in errors)

    def test_empty_team(self):
        errors, _ = _validate_config(
            self._params(), self._wps(), [], [])
        assert any("empty" in e.lower() for e in errors)

    def test_duplicate_wp_id(self):
        wps = [
            {"id": "WP1", "name": "A", "fte": 1, "start": 1, "end": 12},
            {"id": "WP1", "name": "B", "fte": 1, "start": 1, "end": 12},
        ]
        errors, _ = _validate_config(
            self._params(), wps, self._team(), [])
        assert any("Duplicate" in e for e in errors)

    def test_external_person_not_in_team(self):
        ext = [{"person": "Unknown", "project": "X", "fte": 1,
                "start": 1, "end": 12}]
        errors, _ = _validate_config(
            self._params(), self._wps(), self._team(), ext)
        assert any("Unknown" in e for e in errors)

    def test_wp_start_gt_end_warning(self):
        wps = [{"id": "WP1", "name": "Bad", "fte": 1,
                "start": 10, "end": 5}]
        _, warnings = _validate_config(
            self._params(), wps, self._team(), [])
        assert any("start" in w.lower() for w in warnings)



class TestRebuildErrors:
    """Rebuild handles error cases gracefully."""

    def test_rebuild_missing_file(self):
        with pytest.raises(FileNotFoundError):
            rebuild_workbook(input_path="/tmp/nonexistent.xlsx")

    def test_rebuild_no_config_sheet(self, tmp_path):
        """Rebuild fails if workbook has no Config sheet."""
        from openpyxl import Workbook
        path = str(tmp_path / "no_config.xlsx")
        wb = Workbook()
        wb.save(path)
        wb.close()
        with pytest.raises(ValueError, match="Config"):
            rebuild_workbook(input_path=path)


class TestBackup:
    """Rebuild creates timestamped backups."""

    def test_rebuild_creates_backup(self, init_wb, tmp_path):
        rebuild_workbook(input_path=init_wb)
        # Check that a backup file was created in the same directory
        backup_files = [f for f in os.listdir(tmp_path)
                        if f != "test.xlsx" and f.endswith(".xlsx")]
        assert len(backup_files) >= 1, "No backup file created"


class TestGradeRoundTrip:
    """Grade rates and roster grades survive rebuild."""

    def test_grade_rates_survive_rebuild(self, tmp_path):
        """Grade rates from Config survive a rebuild round-trip."""
        from budget_planner.reader import read_grade_rates
        from budget_planner.config import GRADE_RATES

        path = str(tmp_path / "grade_test.xlsx")
        init_workbook(output_path=path)

        # Read grade rates from initial workbook
        wb = load_workbook(path)
        original_rates = read_grade_rates(wb["Config"])
        wb.close()
        assert original_rates == dict(GRADE_RATES)

        # Rebuild
        rebuild_workbook(input_path=path)

        # Read grade rates from rebuilt workbook
        wb = load_workbook(path)
        rebuilt_rates = read_grade_rates(wb["Config"])
        wb.close()

        assert rebuilt_rates == original_rates

    def test_roster_grades_survive_rebuild(self, tmp_path):
        """Roster grade assignments survive a rebuild round-trip."""
        from budget_planner.reader import read_team_from_config
        from budget_planner.config import TEAM

        path = str(tmp_path / "roster_grade_test.xlsx")
        init_workbook(output_path=path)

        # Read roster from initial workbook
        wb = load_workbook(path)
        original_roster = read_team_from_config(wb["Config"])
        wb.close()
        assert len(original_roster) == len(TEAM)

        # Rebuild
        rebuild_workbook(input_path=path)

        # Read roster from rebuilt workbook
        wb = load_workbook(path)
        rebuilt_roster = read_team_from_config(wb["Config"])
        wb.close()

        assert len(rebuilt_roster) == len(original_roster)
        for got, expected in zip(rebuilt_roster, original_roster):
            assert got["name"] == expected["name"]
            assert got["max_fte"] == pytest.approx(expected["max_fte"])
            assert got["grade"] == expected["grade"]
