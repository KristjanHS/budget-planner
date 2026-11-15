"""Tests for column grouping by year."""

import pytest
from openpyxl import load_workbook, Workbook
from budget_planner.styles import apply_year_groups, col_letter


@pytest.fixture(scope="module")
def grouping_wb(fresh_wb_path):
    wb = load_workbook(fresh_wb_path)
    yield wb
    wb.close()


@pytest.mark.parametrize("sheet_name", ["Plan", "Capacity", "Actuals"])
def test_year_groups_applied(grouping_wb, sheet_name):
    """Each main sheet should have column groups with outlineLevel >= 1.

    Column C (first month) is intentionally ungrouped so that adjacent
    year-groups don't merge in Excel.  Column D (second month) carries
    the group.
    """
    ws = grouping_wb[sheet_name]
    dim = ws.column_dimensions.get("D")
    assert dim is not None, f"Column D should exist in {sheet_name} dimensions"
    assert dim.outlineLevel is not None and dim.outlineLevel >= 1, (
        f"{sheet_name} column D should have outlineLevel >= 1"
    )


def test_apply_year_groups_default_expands_year1():
    """Without current_month, year 1 is expanded (not hidden), year 2 is hidden.

    First month of each year (cols 3 and 15) is ungrouped to prevent
    adjacent groups from merging in Excel.
    """
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=24)
    # First month of year 1 (col 3) should NOT be grouped
    y1_first = ws.column_dimensions[col_letter(3)]
    assert (y1_first.outlineLevel or 0) == 0, "First month of year 1 should be ungrouped"
    # Second month of year 1 (col 4): grouped, expanded -> hidden=False
    y1_dim = ws.column_dimensions[col_letter(4)]
    assert y1_dim.outlineLevel == 1, "Year 1 should have outlineLevel 1"
    assert not y1_dim.hidden, "Year 1 should not be hidden (expanded)"
    # First month of year 2 (col 15) should NOT be grouped
    y2_first = ws.column_dimensions[col_letter(15)]
    assert (y2_first.outlineLevel or 0) == 0, "First month of year 2 should be ungrouped"
    # Second month of year 2 (col 16): grouped, collapsed -> hidden=True
    y2_dim = ws.column_dimensions[col_letter(16)]
    assert y2_dim.outlineLevel == 1, "Year 2 should have outlineLevel 1"
    assert y2_dim.hidden, "Year 2 should be hidden (collapsed)"
    wb.close()


def test_apply_year_groups_with_current_month():
    """current_month=15 expands year 2 (months 13-24), collapses years 1 and 3."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=36, current_month=15)
    # Year 1: first month (col 3) ungrouped, second month (col 4) hidden (collapsed)
    y1_dim = ws.column_dimensions[col_letter(4)]
    assert y1_dim.hidden, "Year 1 should be hidden"
    # Year 2: first month (col 15) ungrouped, second month (col 16) expanded
    y2_dim = ws.column_dimensions[col_letter(16)]
    assert not y2_dim.hidden, "Year 2 should be expanded"
    # Year 3: first month (col 27) ungrouped, second month (col 28) hidden
    y3_dim = ws.column_dimensions[col_letter(28)]
    assert y3_dim.hidden, "Year 3 should be hidden"
    wb.close()


def test_apply_year_groups_expands_first_two_calendar_years_of_project():
    """With project_start, the project's own first two calendar years expand; later years collapse."""
    from datetime import datetime

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=36, project_start=datetime(2030, 1, 1))
    hidden = [bool(ws.column_dimensions[col_letter(c)].hidden) for c in (4, 16, 28)]  # 2nd month of each year
    assert hidden == [False, False, True]
    wb.close()
