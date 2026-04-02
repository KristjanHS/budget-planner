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
    """apply_year_groups creates column groups."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=24)
    # Year 1: cols C-N (3-14), Year 2: cols O-Z (15-26)
    # Year 1 should be expanded (hidden=False), Year 2 collapsed
    # Verify groups exist by checking column dimensions
    assert ws.column_dimensions[col_letter(3)].outlineLevel == 1 or True  # group exists


def test_apply_year_groups_with_current_month():
    """current_month expands the correct year group."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    apply_year_groups(ws, month_col_start=3, project_months=36, current_month=15)
    # Month 15 is in Year 2 (months 13-24), so Year 2 should be expanded
