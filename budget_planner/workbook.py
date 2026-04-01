"""Workbook orchestrator — create and update modes."""
import os
import shutil
from datetime import datetime

from openpyxl import Workbook

from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.team import build_team_sheet
from budget_planner.capacity import build_capacity_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.meta import build_meta_sheet
from budget_planner.reader import read_workbook


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
                      existing_path: str | None = None,
                      mode: str = "demo") -> str:
    """Generate workbook. If existing_path given, preserve user data."""
    preserved = read_workbook(existing_path) if existing_path else {}

    wb = Workbook()

    # _META is the first sheet (default active sheet)
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = "_META"

    # Plan sheet
    plan_ws = wb.create_sheet("Plan")
    layout = build_plan_sheet(plan_ws, wb, preserved_data=preserved)

    # Team sheet — use preserved data from --update if available
    team_data = None
    if preserved.get("team"):
        reader_team = preserved["team"]
        team_data = []
        for person in reader_team["roster"]:
            shares = reader_team.get("shares", {}).get(person["name"], {})
            team_data.append({
                "name": person["name"],
                "max_fte": person["max_fte"],
                "wps": shares,
            })

    team_ws = wb.create_sheet("Team")
    team_layout = build_team_sheet(team_ws, wb, team_data=team_data)

    # Capacity sheet — use preserved data from --update if available
    capacity_ws = wb.create_sheet("Capacity")
    capacity_layout = build_capacity_sheet(
        capacity_ws, wb, team_layout,
        preserved_data=preserved.get("capacity"), mode=mode,
    )

    # Overwrite planned FTE with cross-sheet formulas from Capacity
    _apply_planned_formulas(plan_ws, layout, capacity_layout)

    # Summary sheet
    build_summary_sheet(wb, layout)

    # _META sheet (needs layout for row references)
    build_meta_sheet(meta_ws, layout)

    # Sheets are already in correct order: _META, Plan, Team, Capacity, Summary
    # (created in that order above)
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
