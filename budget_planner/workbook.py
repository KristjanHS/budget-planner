"""Workbook orchestrator — init, rebuild, snapshot/restore modes."""
import os
import shutil
from datetime import datetime

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.config_sheet import build_config_sheet
from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.capacity import build_capacity_sheet
from budget_planner.actuals import build_actuals_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.trends import build_trends_sheet
from budget_planner.meta import build_meta_sheet
from budget_planner.layout import (
    SHEET_META, SHEET_CONFIG, SHEET_PLAN, SHEET_CAPACITY,
    SHEET_ACTUALS, SHEET_TRENDS,
    SECTION_HEADER_LABELS,
)
from budget_planner.styles import BLACK_FORMULA, HEADER_FONT, style_cell, apply_year_groups


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


def _validate_config(project_params, work_packages, team, external_commitments,
                     *, wp_schedule=None):
    """Validate extracted config data.

    Returns:
        (errors, warnings) — lists of strings. Errors abort rebuild.
    """
    errors: list[str] = []
    warnings: list[str] = []

    # Hard errors
    if project_params["months"] <= 0 or project_params["months"] > 120:
        errors.append(f"PROJECT_MONTHS must be 1-120, got {project_params['months']}")
    rate = project_params["rate"]
    # Rate may be a formula string or ArrayFormula (grade-based weighted average) — skip numeric check
    if isinstance(rate, (int, float)) and rate <= 0:
        errors.append(f"FTE rate must be > 0, got {rate}")
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

    # WP-level warnings
    for wp in work_packages:
        if wp["fte"] <= 0:
            warnings.append(f"WP {wp['id']}: FTE budget is {wp['fte']} (expected > 0)")

    # Schedule validation
    if wp_schedule is not None:
        sched_wp_ids = {entry["id"] for entry in wp_schedule}
        for entry in wp_schedule:
            if entry["id"] not in seen_ids:
                errors.append(
                    f"Schedule entry for '{entry['id']}' has no matching WP in budget table"
                )
            if entry["start"] > entry["end"]:
                warnings.append(
                    f"Schedule {entry['id']}: start ({entry['start']}) > end ({entry['end']})"
                )
        for wp in work_packages:
            if wp["id"] not in sched_wp_ids:
                warnings.append(
                    f"WP {wp['id']} has no schedule entry"
                )

    return errors, warnings


def _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      *, wp_schedule=None, assignments=None, preserved_actuals=None,
                      rebuild_warnings=None, reporting_periods=None,
                      grade_rates=None):
    """Build all sheets into the workbook.

    Expects wb to already have a '_META' sheet as the active/first sheet.
    Capacity monthly values are auto-distributed from assignments.
    Actuals month cells are always empty (or restored from preserved data).
    """
    meta_ws = wb[SHEET_META]

    # Config sheet (holds project params + WP budget/schedule + team roster)
    config_ws = wb.create_sheet(SHEET_CONFIG)
    config_layout = build_config_sheet(config_ws, wb, project_params, work_packages,
                                       wp_schedule=wp_schedule,
                                       team=team,
                                       reporting_periods=reporting_periods,
                                       grade_rates=grade_rates,
                                       assignments=assignments,
                                       external_commitments=external_commitments)

    # Plan sheet
    plan_ws = wb.create_sheet(SHEET_PLAN)
    layout = build_plan_sheet(plan_ws, wb,
                              project_params=project_params,
                              work_packages=work_packages,
                              wp_schedule=wp_schedule,
                              config_layout=config_layout)

    # Capacity sheet (uses config_layout for roster person_rows)
    capacity_ws = wb.create_sheet(SHEET_CAPACITY)
    capacity_layout = build_capacity_sheet(
        capacity_ws, config_layout,
        project_params=project_params, work_packages=work_packages,
        wp_schedule=wp_schedule,
        team=team, external_commitments=external_commitments,
        assignments=assignments,
        config_first_wp_row=layout["config_first_wp_row"],
    )

    # Overwrite planned FTE with cross-sheet formulas from Capacity
    _apply_planned_formulas(plan_ws, layout, capacity_layout,
                            work_packages=work_packages,
                            project_params=project_params,
                            wp_schedule=wp_schedule)

    # Trends sheet stub + current_month named range
    trends_ws = wb.create_sheet(SHEET_TRENDS)
    trends_ws.sheet_properties.tabColor = "FFD700"
    style_cell(trends_ws, 1, 1, "Current Month:", font=HEADER_FONT, border=False)
    # Formula: auto-calculate from project start date and TODAY()
    _cm_formula = (
        f"=MAX(1,MIN({SHEET_CONFIG}!B3,"
        f"(YEAR(TODAY())-YEAR(project_start))*12"
        f"+MONTH(TODAY())-MONTH(project_start)+1))"
    )
    style_cell(trends_ws, 1, 2, font=BLACK_FORMULA, border=False)
    trends_ws["B1"].value = _cm_formula
    trends_ws["B1"].comment = Comment(
        "Auto-calculated from project start and today's date. "
        "Override with a static number to freeze.",
        "Programme Budget Tool",
    )
    wb.defined_names.add(DefinedName("current_month", attr_text=f"{SHEET_TRENDS}!$B$1"))

    # Actuals sheet
    actuals_layout = build_actuals_sheet(
        wb, layout,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        preserved=preserved_actuals,
    )

    # Summary sheet
    build_summary_sheet(wb, layout, actuals_layout,
                        project_params=project_params,
                        work_packages=work_packages,
                        wp_schedule=wp_schedule,
                        reporting_periods=reporting_periods,
                        config_layout=config_layout)

    # Trends sheet (populate after Actuals and Summary so layouts are available)
    build_trends_sheet(trends_ws, wb, layout, actuals_layout,
                       capacity_layout, config_layout,
                       project_params=project_params,
                       work_packages=work_packages,
                       wp_schedule=wp_schedule,
                       team=team,
                       reporting_periods=reporting_periods)

    # _META sheet (needs layout for row references)
    build_meta_sheet(meta_ws, layout,
                     project_params=project_params,
                     work_packages=work_packages,
                     team=team,
                     external_commitments=external_commitments,
                     rebuild_warnings=rebuild_warnings,
                     config_layout=config_layout,
                     grade_rates=grade_rates)

    # Tab order: _META, Config, Plan, Capacity, Actuals, Trends, Summary
    wb.move_sheet(SHEET_META, offset=0)
    wb.move_sheet(SHEET_TRENDS, offset=1)  # move after Actuals


def _apply_protection(wb):
    """Protect all sheets with password 'grant'.

    Blue input cells are already unlocked via style_cell auto-detection.
    Explicitly allows column/row resizing so users can auto-fit widths.
    """
    for ws in wb.worksheets:
        ws.protection.set_password("grant")
        ws.protection.formatColumns = False
        ws.protection.formatRows = False


def _apply_all_year_groups(wb, project_params, current_month=None):
    """Apply column grouping by year to sheets with monthly columns."""
    for sheet_name in [SHEET_PLAN, SHEET_CAPACITY, SHEET_ACTUALS]:
        if sheet_name in wb.sheetnames:
            apply_year_groups(wb[sheet_name], 3, project_params["months"],
                              current_month,
                              project_start=project_params["start"])


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
                  assignments: list[dict] | None = None) -> str:
    """Create a fresh workbook from config.py seed data.

    When assignments is None, produces empty Capacity (no WP rows per person).
    Pass DEMO_CAPACITY_ASSIGNMENTS for demo mode.

    Args:
        output_path: Where to save the workbook.
        assignments: Optional list of person-WP assignment dicts.

    Returns:
        The output path.
    """
    from budget_planner import config

    project_params = {
        "title": config.PROJECT_TITLE,
        "start": config.PROJECT_START,
        "months": config.PROJECT_MONTHS,
        "rate": config.DEFAULT_RATE,
        "eu_ceiling_pm": config.EU_ANNUAL_CEILING_PM,
    }
    work_packages = config.WORK_PACKAGES
    wp_schedule = config.WP_SCHEDULE
    team = config.TEAM
    external_commitments = config.EXTERNAL_COMMITMENTS
    reporting_periods = config.REPORTING_PERIODS

    wb = Workbook()
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = SHEET_META

    grade_rates = config.GRADE_RATES

    _build_all_sheets(wb, project_params, work_packages, team, external_commitments,
                      wp_schedule=wp_schedule,
                      assignments=assignments,
                      reporting_periods=reporting_periods,
                      grade_rates=grade_rates)

    # _apply_protection(wb)  # Temporarily disabled
    from datetime import date
    from dateutil.relativedelta import relativedelta as _rd
    _today = date.today()
    _pstart = project_params["start"]
    _delta = _rd(_today, _pstart.date() if hasattr(_pstart, 'date') else _pstart)
    _current_month = max(1, min(_delta.years * 12 + _delta.months + 1,
                                project_params["months"]))
    _apply_all_year_groups(wb, project_params, current_month=_current_month)

    wb.active = wb.sheetnames.index(SHEET_CONFIG)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    _backup_existing(output_path)
    wb.save(output_path)
    print(f"Saved to {output_path}")
    return output_path


def rebuild_workbook(input_path: str = "output/budget_plan.xlsx") -> str:
    """Rebuild workbook from existing file, preserving Capacity and Actuals data.

    Reads Config sheet (params, WPs, team roster) and Capacity sheet from the
    existing workbook. Validates extracted data, regenerates all sheets, and
    re-injects preserved values.

    Args:
        input_path: Path to the existing workbook.

    Returns:
        The output path (same as input_path).
    """
    from openpyxl import load_workbook
    from budget_planner.reader import (
        read_config_sheet, read_team_from_config, read_capacity_sheet,
        read_actuals_values, read_reporting_periods, read_grade_rates,
        read_assignments_from_config, read_externals_from_config,
        _find_header_row,
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
    if SHEET_CONFIG not in existing_wb.sheetnames:
        existing_wb.close()
        raise ValueError(
            "Workbook has no Config sheet. "
            "Run 'python -m budget_planner init' to create a fresh workbook first."
        )
    project_params, work_packages, wp_schedule = read_config_sheet(existing_wb[SHEET_CONFIG])

    # Team roster from Config sheet
    team = read_team_from_config(existing_wb[SHEET_CONFIG])

    # Grade rates from Config sheet
    grade_rates = read_grade_rates(existing_wb[SHEET_CONFIG])

    # Reporting periods from Config sheet
    reporting_periods = read_reporting_periods(
        existing_wb[SHEET_CONFIG],
        fallback_months=project_params["months"],
    )

    # Assignments + external commitments — prefer Config, fall back to Capacity.
    # Fallback is triggered by the *section header* being absent (old workbook),
    # not by the section being present-but-empty (user-cleared). This preserves
    # intentional empty state after migration.
    config_ws = existing_wb[SHEET_CONFIG]
    has_assign_section = _find_header_row(
        config_ws, SECTION_HEADER_LABELS["assignments"], required=False
    ) is not None
    has_ext_section = _find_header_row(
        config_ws, SECTION_HEADER_LABELS["externals"], required=False
    ) is not None

    assignments = read_assignments_from_config(config_ws) if has_assign_section else []
    external_commitments = (
        read_externals_from_config(config_ws) if has_ext_section else []
    )

    if not has_assign_section or not has_ext_section:
        if SHEET_CAPACITY not in existing_wb.sheetnames:
            existing_wb.close()
            raise ValueError("Workbook has no Capacity sheet.")
        capacity_data = read_capacity_sheet(
            existing_wb[SHEET_CAPACITY], work_packages,
            project_months=project_params["months"],
        )
        if not has_assign_section:
            assignments = capacity_data["assignments"]
        if not has_ext_section:
            external_commitments = capacity_data["external_commitments"]

    # Preserve Actuals EUR values (always)
    preserved_actuals = None
    if SHEET_ACTUALS in existing_wb.sheetnames:
        preserved_actuals = read_actuals_values(
            existing_wb[SHEET_ACTUALS], project_params["months"],
        )

    # Calculate current_month from project start and today's date
    from datetime import date
    from dateutil.relativedelta import relativedelta as _rd
    _today = date.today()
    _pstart = project_params["start"]
    _delta = _rd(_today, _pstart.date() if hasattr(_pstart, 'date') else _pstart)
    current_month = max(1, min(_delta.years * 12 + _delta.months + 1,
                               project_params["months"]))

    existing_wb.close()

    # ── Validate ──
    errors, warnings = _validate_config(
        project_params, work_packages, team, external_commitments,
        wp_schedule=wp_schedule,
    )

    # Check for orphaned data
    wp_ids = {wp["id"] for wp in work_packages}
    team_names = {m["name"] for m in team}

    # Validate assignments against current WPs and team
    valid_assignments = []
    for asgn in assignments:
        if asgn["person"] not in team_names:
            warnings.append(
                f"Orphaned Assignment: person '{asgn['person']}' removed from team"
            )
        elif asgn["wp"] not in wp_ids:
            warnings.append(
                f"Orphaned Assignment: WP '{asgn['wp']}' removed for {asgn['person']}"
            )
        else:
            valid_assignments.append(asgn)
    assignments = valid_assignments

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
    meta_ws.title = SHEET_META

    _build_all_sheets(
        wb, project_params, work_packages, team, external_commitments,
        wp_schedule=wp_schedule,
        assignments=assignments,
        preserved_actuals=preserved_actuals,
        rebuild_warnings=warnings if warnings else None,
        reporting_periods=reporting_periods,
        grade_rates=grade_rates,
    )

    # _apply_protection(wb)  # Temporarily disabled
    _apply_all_year_groups(wb, project_params, current_month=current_month)

    wb.active = wb.sheetnames.index(SHEET_CONFIG)

    # Back up then overwrite
    _backup_existing(input_path)
    wb.save(input_path)
    print(f"Rebuilt: {input_path}")
    return input_path


def seed_demo_data(workbook_path: str) -> str:
    """Seed Actuals EUR and current_month after workbook is built.

    Capacity FTE values are now auto-distributed by the builder via
    assignments — this function only seeds Actuals EUR spend and
    sets current_month on Trends.

    Args:
        workbook_path: Path to an existing workbook (from init_workbook).

    Returns:
        The workbook path.
    """
    from openpyxl import load_workbook as _load_wb
    from budget_planner.config import (
        DEMO_CURRENT_MONTH, DEMO_ACTUALS_MULTIPLIERS, WORK_PACKAGES, WP_SCHEDULE,
    )
    from budget_planner.layout import MONTH_COL_START, wp_month_bounds

    wb = _load_wb(workbook_path)

    # ── Seed Actuals EUR values ──
    act_ws = wb[SHEET_ACTUALS]

    # Find first WP row in Actuals (row after "WP ID" header)
    first_wp_row = None
    for r in range(1, act_ws.max_row + 1):
        if act_ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break

    if first_wp_row is not None:
        # Read per-WP rates from Config col D
        from budget_planner.reader import _find_header_row, _data_row_after
        config_ws = wb[SHEET_CONFIG]
        wp_budget_row = _data_row_after(
            _find_header_row(config_ws, SECTION_HEADER_LABELS["wp_budget"])
        )
        for i, wp in enumerate(WORK_PACKAGES):
            row = first_wp_row + i
            wp_rate = config_ws.cell(wp_budget_row + i, 4).value or 6500
            multiplier = DEMO_ACTUALS_MULTIPLIERS[i % len(DEMO_ACTUALS_MULTIPLIERS)]
            wp_start, wp_end = wp_month_bounds(wp["id"], WP_SCHEDULE)
            active_months = wp_end - wp_start + 1
            monthly_planned = (wp["fte"] / active_months) * wp_rate
            for m_idx in range(wp_start - 1, min(wp_end, DEMO_CURRENT_MONTH)):
                col = MONTH_COL_START + m_idx
                eur_value = round(monthly_planned * multiplier)
                act_ws.cell(row, col).value = eur_value

    wb.save(workbook_path)
    wb.close()
    print(f"Demo data seeded: {workbook_path}")
    return workbook_path
