"""Tab order: Config, Timelog Plan, Actuals-Hours must lead the sheet list.

Phase 4 of the Timelog Plan v2 design: the two frequently-edited tabs
(Timelog Plan, Actuals-Hours) sit at positions 2 and 3 (1-indexed) so the
user can find them without scrolling. Position 1 remains Config.
"""

from __future__ import annotations

from openpyxl import load_workbook

from budget_planner.layout import (
    SHEET_ACTUALS_HOURS,
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
)


def test_config_timelog_actuals_hours_first(assigned_wb_path):
    wb = load_workbook(assigned_wb_path, data_only=False)
    assert wb.sheetnames[0:3] == [
        SHEET_CONFIG,
        SHEET_TIMELOG_PLAN,
        SHEET_ACTUALS_HOURS,
    ], wb.sheetnames
    wb.close()
