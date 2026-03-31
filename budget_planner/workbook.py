"""Workbook orchestrator — create and update modes."""
import os

from openpyxl import Workbook

from budget_planner.plan import build_plan_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.meta import build_meta_sheet
from budget_planner.reader import read_workbook


def generate_workbook(output_path: str = "output/budget_plan.xlsx",
                      existing_path: str | None = None) -> str:
    """Generate workbook. If existing_path given, preserve user data."""
    preserved = read_workbook(existing_path) if existing_path else {}

    wb = Workbook()

    # _META is the first sheet (default active sheet)
    meta_ws = wb.active
    meta_ws.title = "_META"

    # Plan sheet
    plan_ws = wb.create_sheet("Plan")
    layout = build_plan_sheet(plan_ws, wb, preserved_data=preserved)

    # Summary sheet
    build_summary_sheet(wb, layout)

    # _META sheet (needs layout for row references)
    build_meta_sheet(meta_ws, layout)

    # Reorder: _META, Plan, Summary
    wb.move_sheet("_META", offset=0)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    wb.save(output_path)
    print(f"Saved to {output_path}")
    return output_path
