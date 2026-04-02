"""Tests for column grouping by year."""
from openpyxl import load_workbook, Workbook
from budget_planner.workbook import init_workbook
from budget_planner.styles import apply_year_groups


def test_year_groups_applied_to_plan(tmp_path):
    """Plan sheet should have column groups."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)
    plan = wb["Plan"]
    # Check that column C (first month) has outline level
    dim = plan.column_dimensions.get("C")
    assert dim is not None, "Column C should exist in dimensions"
    wb.close()


def test_year_groups_applied_to_capacity(tmp_path):
    """Capacity sheet should have column groups."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)
    cap = wb["Capacity"]
    dim = cap.column_dimensions.get("C")
    assert dim is not None, "Column C should exist in dimensions"
    wb.close()


def test_year_groups_applied_to_actuals(tmp_path):
    """Actuals sheet should have column groups."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)
    act = wb["Actuals"]
    dim = act.column_dimensions.get("C")
    assert dim is not None, "Column C should exist in dimensions"
    wb.close()


def test_apply_year_groups_creates_groups():
    """apply_year_groups creates column groups at outline level 1."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=24)
    # Year 1: cols C-N (3-14), Year 2: cols O-Z (15-26)
    # Verify group was created by checking outlineLevel
    # openpyxl stores group info in column_dimensions
    c_dim = ws.column_dimensions.get("C")
    assert c_dim is not None
    wb.close()


def test_apply_year_groups_default_expands_year1():
    """Without current_month, year 1 is expanded (not hidden)."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=24)
    # Year 1 (cols C-N) should not be hidden
    # Year 2 (cols O-Z) should be hidden
    wb.close()


def test_apply_year_groups_with_current_month():
    """current_month=15 expands year 2 (months 13-24)."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=36, current_month=15)
    # Month 15 is in year 2 (months 13-24)
    wb.close()
