"""Workbook orchestrator — init, rebuild, snapshot/restore modes."""
import os
import shutil
from datetime import datetime

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.config_sheet import build_config_sheet
from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.team import build_team_sheet
from budget_planner.capacity import build_capacity_sheet
from budget_planner.actuals import build_actuals_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.trends import build_trends_sheet
from budget_planner.meta import build_meta_sheet
from budget_planner.styles import BLUE_INPUT, HEADER_FONT, style_cell, apply_year_groups


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


def _validate_config(project_params, work_packages, team, external_commitments):
    """Validate extracted config data.

    Returns:
        (errors, warnings) — lists of strings. Errors abort rebuild.
    """
    errors: list[str] = []
    warnings: list[str] = []

    # Hard errors
    if project_params["months"] <= 0 or project_params["months"] > 120:
        errors.append(f"PROJECT_MONTHS must be 1-120, got {project_params['months']}")
    if project_params["rate"] <= 0:
        errors.append(f"FTE rate must be > 0, got {project_params['rate']}")
    if not work_packages:
        errors.append("Work packages table is empty")
    if not team:
        errors.append("Team roster is empty")

    # Duplicate WP IDs
    seen_ids: set[str] = set()
    for wp in work_packages:
        if wp["id"] in seen_ids:
            errors.append(f"Duplicate WP ID: {wp['id']}")
        seen_ids.add(wp["id"])

    # External commitment person must exist in team
    team_names = {m["name"] for m in team}
    for ext in external_commitments:
        if ext["person"] not in team_names:
            errors.append(
                f"External commitment person '{ext['person']}' not in team roster"
            )

    # Warnings
    for wp in work_packages:
        if wp["start"] > wp["end"]:
            warnings.append(f"WP {wp['id']}: start ({wp['start']}) > end ({wp['end']})")
        if wp["fte"] <= 0:
            warnings.append(f"WP {wp['id']}: FTE budget is {wp['fte']} (expected > 0)")

    for member in team:
        total_share = sum(member["wps"].values()) if member["wps"] else 0
        if member["wps"] and abs(total_share - 1.0) > 0.05:
            warnings.append(
                f"'{member['name']}': share seeds sum to {total_share:.2f} (expected ~1.0)"
            )

    return errors, warnings


def _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      *, mode="demo", preserved_capacity=None, preserved_actuals=None,
                      rebuild_warnings=None):
    """Build all sheets into the workbook.

    Expects wb to already have a '_META' sheet as the active/first sheet.
    """
    meta_ws = wb["_META"]

    # Config sheet (new — holds project params + WP table)
    config_ws = wb.create_sheet("Config")
    build_config_sheet(config_ws, wb, project_params, work_packages)

    # Plan sheet
    plan_ws = wb.create_sheet("Plan")
    layout = build_plan_sheet(plan_ws, wb,
                              project_params=project_params,
                              work_packages=work_packages)

    # Team sheet
    team_ws = wb.create_sheet("Team")
    team_layout = build_team_sheet(team_ws, wb,
                                   team=team,
                                   work_packages=work_packages)

    # Capacity sheet
    capacity_ws = wb.create_sheet("Capacity")
    capacity_layout = build_capacity_sheet(
        capacity_ws, wb, team_layout, mode=mode,
        project_params=project_params, work_packages=work_packages,
        team=team, external_commitments=external_commitments,
        preserved=preserved_capacity,
    )

    # Overwrite planned FTE with cross-sheet formulas from Capacity
    _apply_planned_formulas(plan_ws, layout, capacity_layout,
                            work_packages=work_packages,
                            project_params=project_params)

    # Trends sheet skeleton + current_month named range
    trends_ws = wb.create_sheet("Trends")
    trends_ws.sheet_properties.tabColor = "FFD700"
    style_cell(trends_ws, 1, 1, "Current Month:", font=HEADER_FONT, border=False)
    style_cell(trends_ws, 1, 2, 1, font=BLUE_INPUT, border=False)
    trends_ws["B1"].comment = Comment(
        "Set to the last month with completed actuals (e.g., 12 for end of Year 1)",
        "Programme Budget Tool",
    )
    wb.defined_names.add(DefinedName("current_month", attr_text="Trends!$B$1"))

    # Actuals sheet
    actuals_layout = build_actuals_sheet(
        wb, layout,
        project_params=project_params,
        work_packages=work_packages,
        preserved=preserved_actuals,
    )

    # Summary sheet
    build_summary_sheet(wb, layout, actuals_layout,
                        project_params=project_params,
                        work_packages=work_packages)

    # Trends sheet (populate after Actuals and Summary so layouts are available)
    build_trends_sheet(trends_ws, wb, layout, actuals_layout,
                       capacity_layout, team_layout,
                       project_params=project_params,
                       work_packages=work_packages,
                       team=team)

    # _META sheet (needs layout for row references)
    build_meta_sheet(meta_ws, layout,
                     project_params=project_params,
                     work_packages=work_packages,
                     team=team,
                     external_commitments=external_commitments,
                     rebuild_warnings=rebuild_warnings)

    # Sheets: _META, Config, Plan, Team, Capacity, Trends, Actuals, Summary
    wb.move_sheet("_META", offset=0)


def _apply_protection(wb):
    """Protect all sheets with password 'grant'.

    Blue input cells are already unlocked via style_cell auto-detection.
    """
    for ws in wb.worksheets:
        ws.protection.set_password("grant")


def _apply_all_year_groups(wb, project_params, current_month=None):
    """Apply column grouping by year to sheets with monthly columns."""
    for sheet_name in ["Plan", "Capacity", "Actuals"]:
        if sheet_name in wb.sheetnames:
            apply_year_groups(wb[sheet_name], 3, project_params["months"],
                              current_month)


def _backup_existing(filepath):
    """Back up existing file with timestamp. Returns backup path or None."""
    if os.path.exists(filepath):
        stem, ext = os.path.splitext(filepath)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{stem}.{timestamp}{ext}"
        shutil.copy2(filepath, backup_path)
        print(f"Backup: {backup_path}")
        return backup_path
    return None


def init_workbook(output_path: str = "output/budget_plan.xlsx",
                  mode: str = "demo") -> str:
    """Create a fresh workbook from config.py seed data.

    Args:
        output_path: Where to save the workbook.
        mode: 'demo' pre-fills Capacity; 'skeleton' leaves it empty.

    Returns:
        The output path.
    """
    from budget_planner import config

    project_params = {
        "title": config.PROJECT_TITLE,
        "start": config.PROJECT_START,
        "months": config.PROJECT_MONTHS,
        "rate": config.DEFAULT_RATE,
    }
    work_packages = config.WORK_PACKAGES
    team = config.TEAM
    external_commitments = config.EXTERNAL_COMMITMENTS

    wb = Workbook()
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = "_META"

    _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      mode=mode)

    _apply_protection(wb)
    _apply_all_year_groups(wb, project_params)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    _backup_existing(output_path)
    wb.save(output_path)
    print(f"Saved to {output_path}")
    return output_path


def rebuild_workbook(input_path: str = "output/budget_plan.xlsx",
                     mode: str = "demo") -> str:
    """Rebuild workbook from existing file, preserving Capacity and Actuals data.

    Reads Config sheet, Team sheet, and Capacity sheet from the existing workbook.
    Validates extracted data, regenerates all sheets, and re-injects preserved values.

    Args:
        input_path: Path to the existing workbook.
        mode: 'demo' preserves Capacity FTE values; 'skeleton' clears them.

    Returns:
        The output path (same as input_path).
    """
    from openpyxl import load_workbook
    from budget_planner.reader import (
        read_config_sheet, read_team_sheet, read_capacity_sheet,
        read_actuals_values,
    )

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File not found: {input_path}")

    # Print source file timestamp
    mod_time = datetime.fromtimestamp(os.path.getmtime(input_path))
    print(f"Reading from: {os.path.basename(input_path)} "
          f"(modified {mod_time.strftime('%Y-%m-%d %H:%M')})")

    existing_wb = load_workbook(input_path, data_only=False)

    # ── Extract data ──

    # Config sheet → project params + WPs
    if "Config" not in existing_wb.sheetnames:
        existing_wb.close()
        raise ValueError(
            "Workbook has no Config sheet. "
            "Run 'python -m budget_planner init' to create a fresh workbook first."
        )
    project_params, work_packages = read_config_sheet(existing_wb["Config"])

    # Team sheet → roster + shares
    if "Team" not in existing_wb.sheetnames:
        existing_wb.close()
        raise ValueError("Workbook has no Team sheet.")
    team_data = read_team_sheet(existing_wb["Team"])
    team = [
        {
            "name": p["name"],
            "max_fte": p["max_fte"],
            "wps": team_data["shares"].get(p["name"], {}),
        }
        for p in team_data["roster"]
    ]

    # Capacity sheet → external commitments + FTE values
    if "Capacity" not in existing_wb.sheetnames:
        existing_wb.close()
        raise ValueError("Workbook has no Capacity sheet.")
    capacity_data = read_capacity_sheet(
        existing_wb["Capacity"], work_packages,
        project_months=project_params["months"],
    )
    external_commitments = capacity_data["external_commitments"]

    # Preserve Capacity FTE values (unless skeleton mode)
    preserved_capacity = None
    if mode != "skeleton":
        preserved_capacity = {}
        for person_name, groups in capacity_data["person_groups"].items():
            for wp_id, monthly in groups["wp_values"].items():
                preserved_capacity[(person_name, wp_id)] = monthly

    # Preserve Actuals EUR values (always)
    preserved_actuals = None
    if "Actuals" in existing_wb.sheetnames:
        preserved_actuals = read_actuals_values(
            existing_wb["Actuals"], project_params["months"],
        )

    # Read current_month from Trends for column grouping
    current_month = 1
    if "Trends" in existing_wb.sheetnames:
        cm = existing_wb["Trends"].cell(1, 2).value
        if isinstance(cm, (int, float)):
            current_month = int(cm)

    existing_wb.close()

    # ── Validate ──
    errors, warnings = _validate_config(
        project_params, work_packages, team, external_commitments
    )

    # Check for orphaned preserved data
    wp_ids = {wp["id"] for wp in work_packages}
    team_names = {m["name"] for m in team}

    if preserved_capacity:
        for key in list(preserved_capacity.keys()):
            person_name, wp_id = key
            if person_name not in team_names:
                warnings.append(
                    f"Orphaned Capacity: person '{person_name}' removed from team"
                )
                del preserved_capacity[key]
            elif wp_id not in wp_ids:
                warnings.append(
                    f"Orphaned Capacity: WP '{wp_id}' removed for {person_name}"
                )
                del preserved_capacity[key]

    if preserved_actuals:
        for wp_id in list(preserved_actuals.keys()):
            if wp_id not in wp_ids:
                warnings.append(f"Orphaned Actuals: WP '{wp_id}' removed")
                del preserved_actuals[wp_id]

    if errors:
        raise ValueError(
            "Rebuild validation failed:\n"
            + "\n".join(f"  ERROR: {e}" for e in errors)
        )

    for w in warnings:
        print(f"  WARNING: {w}")

    # ── Generate new workbook ──
    wb = Workbook()
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = "_META"

    _build_all_sheets(
        wb, project_params, work_packages, team, external_commitments,
        mode=mode,
        preserved_capacity=preserved_capacity,
        preserved_actuals=preserved_actuals,
        rebuild_warnings=warnings if warnings else None,
    )

    _apply_protection(wb)
    _apply_all_year_groups(wb, project_params, current_month=current_month)

    # Back up then overwrite
    _backup_existing(input_path)
    wb.save(input_path)
    print(f"Rebuilt: {input_path}")
    return input_path


def generate_workbook(output_path: str = "output/budget_plan.xlsx",
                      mode: str = "demo") -> str:
    """Generate workbook (backward-compatible wrapper for init_workbook)."""
    return init_workbook(output_path=output_path, mode=mode)
