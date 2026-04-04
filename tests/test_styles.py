"""Tests for styles module additions."""
from openpyxl import Workbook
from budget_planner.styles import style_cell, apply_year_groups, col_letter


def test_style_cell_unlocked():
    """unlocked=True sets Protection(locked=False) on cell."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    cell = style_cell(ws, 1, 1, value="test", unlocked=True)
    assert cell.protection.locked is False


def test_style_cell_default_locked():
    """Default cells should not have unlocked protection."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    cell = style_cell(ws, 1, 1, value="test")
    # Default openpyxl cells are locked (True)
    assert cell.protection.locked is True or cell.protection.locked is None


def test_apply_year_groups_basic():
    """apply_year_groups creates column groups at outline level 1."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=24)
    # First month of each year is ungrouped; group starts at 2nd month
    # Year 1: col D (4) grouped, Year 2: col P (16) grouped
    assert ws.column_dimensions[col_letter(4)].outlineLevel == 1
    assert ws.column_dimensions[col_letter(16)].outlineLevel == 1
    # Default: Year 1 expanded, Year 2 collapsed
    assert not ws.column_dimensions[col_letter(4)].hidden
    assert ws.column_dimensions[col_letter(16)].hidden


def test_apply_year_groups_with_current_month():
    """current_month expands the correct year group."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=36, current_month=15)
    # Month 15 is in Year 2 (months 13-24), so Year 2 should be expanded
    # Check 2nd month of each year (1st is ungrouped)
    assert ws.column_dimensions[col_letter(4)].hidden, "Year 1 should be hidden"
    assert not ws.column_dimensions[col_letter(16)].hidden, "Year 2 should be expanded"
    assert ws.column_dimensions[col_letter(28)].hidden, "Year 3 should be hidden"
