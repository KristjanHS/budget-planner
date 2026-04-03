"""Tests for column grouping by year."""
import pytest
from openpyxl import load_workbook, Workbook
from budget_planner.workbook import init_workbook
from budget_planner.styles import apply_year_groups, col_letter


@pytest.fixture(scope="module")
def grouping_wb(tmp_path_factory):
    """Generate one workbook for all grouping tests (module-scoped)."""
    path = str(tmp_path_factory.mktemp("grouping") / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)
    yield wb
    wb.close()


@pytest.mark.parametrize("sheet_name", ["Plan", "Capacity", "Actuals"])
def test_year_groups_applied(grouping_wb, sheet_name):
    """Each main sheet should have column groups with outlineLevel >= 1."""
    ws = grouping_wb[sheet_name]
    dim = ws.column_dimensions.get("C")
    assert dim is not None, f"Column C should exist in {sheet_name} dimensions"
    assert dim.outlineLevel is not None and dim.outlineLevel >= 1, (
        f"{sheet_name} column C should have outlineLevel >= 1"
    )


def test_apply_year_groups_default_expands_year1():
    """Without current_month, year 1 is expanded (not hidden), year 2 is hidden."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=24)
    # openpyxl stores group attributes on the first column of each group range
    # Year 1 start col C (col 3): expanded -> hidden=False
    y1_dim = ws.column_dimensions[col_letter(3)]
    assert y1_dim.outlineLevel == 1, "Year 1 should have outlineLevel 1"
    assert not y1_dim.hidden, "Year 1 should not be hidden (expanded)"
    # Year 2 start col O (col 15): collapsed -> hidden=True
    y2_dim = ws.column_dimensions[col_letter(15)]
    assert y2_dim.outlineLevel == 1, "Year 2 should have outlineLevel 1"
    assert y2_dim.hidden, "Year 2 should be hidden (collapsed)"
    wb.close()


def test_apply_year_groups_with_current_month():
    """current_month=15 expands year 2 (months 13-24), collapses years 1 and 3."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=36, current_month=15)
    # Year 1 start col C (col 3): hidden (collapsed)
    y1_dim = ws.column_dimensions[col_letter(3)]
    assert y1_dim.hidden, "Year 1 should be hidden"
    # Year 2 start col O (col 15): expanded (month 15 is in year 2)
    y2_dim = ws.column_dimensions[col_letter(15)]
    assert not y2_dim.hidden, "Year 2 should be expanded"
    # Year 3 start col AA (col 27): hidden (collapsed)
    y3_dim = ws.column_dimensions[col_letter(27)]
    assert y3_dim.hidden, "Year 3 should be hidden"
    wb.close()
