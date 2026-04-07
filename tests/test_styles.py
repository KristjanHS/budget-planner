"""Tests for styles module additions."""
from openpyxl import Workbook
from budget_planner.styles import style_cell


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
