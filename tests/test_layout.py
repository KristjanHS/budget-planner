"""Tests for layout helpers — `cell_ref` keeps cross-sheet refs layout-derived
rather than hardcoded, so shifting a column in `layout.py` doesn't silently
break formulas in other builders."""

from budget_planner.layout import cell_ref, SHEET_CONFIG, COL_WP_CEILING, COL_ROSTER_FTE


def test_cell_ref_absolute_default():
    assert cell_ref("Trends", 1, 2) == "Trends!$B$1"


def test_cell_ref_relative():
    assert cell_ref(SHEET_CONFIG, 8, COL_WP_CEILING, absolute=False) == "Config!E8"


def test_cell_ref_roster_column_constant():
    # Ensures the helper resolves layout column constants to the expected letter.
    assert cell_ref(SHEET_CONFIG, 22, COL_ROSTER_FTE) == "Config!$B$22"


def test_cell_ref_multi_letter_column():
    assert cell_ref("Plan", 5, 27) == "Plan!$AA$5"
