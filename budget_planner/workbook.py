"""Workbook orchestrator — create and update modes."""
import os
import shutil
from datetime import datetime

from openpyxl import Workbook

from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.team import build_team_sheet
from budget_planner.capacity import build_capacity_sheet
from budget_planner.actuals import build_actuals_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.meta import build_meta_sheet


def snapshot_workbook(filepath: str) -> str:
    """Copy workbook to <stem>.snapshot.xlsx."""
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")
    stem, ext = os.path.splitext(filepath)
    snapshot_path = f"{stem}.snapshot{ext}"
    shutil.copy2(filepath, snapshot_path)
    print(f"Snapshot saved: {snapshot_path}")
    return snapshot_path


def restore_workbook(filepath: str) -> str:
    """Restore workbook from <stem>.snapshot.xlsx, backing up current first."""
    stem, ext = os.path.splitext(filepath)
    snapshot_path = f"{stem}.snapshot{ext}"
    if not os.path.exists(snapshot_path):
        raise FileNotFoundError(
            f"No snapshot found at {snapshot_path}. "
            f"Create one first with: python -m budget_planner --snapshot {filepath}"
        )
    # Back up current file with timestamp
    if os.path.exists(filepath):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{stem}.{timestamp}{ext}"
        shutil.copy2(filepath, backup_path)
        print(f"Backup of current: {backup_path}")
    shutil.copy2(snapshot_path, filepath)
    print(f"Restored from: {snapshot_path}")
    return filepath


def generate_workbook(output_path: str = "output/budget_plan.xlsx",
                      mode: str = "demo") -> str:
    """Generate workbook."""
    wb = Workbook()

    # _META is the first sheet (default active sheet)
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = "_META"

    # Plan sheet
    plan_ws = wb.create_sheet("Plan")
    layout = build_plan_sheet(plan_ws, wb)

    # Team sheet
    team_ws = wb.create_sheet("Team")
    team_layout = build_team_sheet(team_ws, wb)

    # Capacity sheet
    capacity_ws = wb.create_sheet("Capacity")
    capacity_layout = build_capacity_sheet(
        capacity_ws, wb, team_layout, mode=mode,
    )

    # Overwrite planned FTE with cross-sheet formulas from Capacity
    _apply_planned_formulas(plan_ws, layout, capacity_layout)

    # Actuals sheet
    actuals_layout = build_actuals_sheet(wb, layout)

    # Summary sheet
    build_summary_sheet(wb, layout, actuals_layout)

    # _META sheet (needs layout for row references)
    build_meta_sheet(meta_ws, layout)

    # Sheets created in order: _META, Plan, Team, Capacity, Actuals, Summary
    wb.move_sheet("_META", offset=0)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # Back up existing file with timestamp before overwriting
    if os.path.exists(output_path):
        stem, ext = os.path.splitext(output_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{stem}.{timestamp}{ext}"
        shutil.copy2(output_path, backup_path)
        print(f"Backup: {backup_path}")

    wb.save(output_path)
    print(f"Saved to {output_path}")
    return output_path
