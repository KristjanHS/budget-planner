"""Workbook orchestrator — init, rebuild, snapshot/restore modes.

Build order (load-bearing — do not reorder without reading T5 in
docs/plans/archived/2026-04-12-brittleness-refactor.md)::

    Config ──► Capacity ──► Plan ──► (rewrite Plan FTE)
                               │
                               ├──► Trends(stub = current_month named range)
                               └──► Actuals ──► Summary ──► Trends(full)

Coupling points:
- Capacity consumes ``config_layout`` (param rows, first WP row).
- Plan consumes ``config_layout`` (param rows, first WP row).
- ``_apply_planned_formulas`` overwrites Plan FTE cells using
  ``capacity_layout`` — must run after both Capacity and Plan.
- Actuals formulas reference the ``current_month`` named range — the Trends
  stub that creates it must run before Actuals.
- Summary and full Trends consume every prior layout dict and must
  run last in that order.

``BuildContext`` below threads the accumulated layout dicts through
``_build_all_sheets`` so the DAG is explicit rather than hidden in locals.
"""

import os
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.config_sheet import build_config_sheet
from budget_planner.plan import build_plan_sheet, _apply_planned_formulas
from budget_planner.capacity import (
    build_capacity_sheet, apply_capacity_todate_formulas,
)
from budget_planner.actuals import build_actuals_sheet
from budget_planner.actuals_hours import build_actuals_hours_sheet
from budget_planner.plan_forecast import build_plan_forecast_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.timelog_plan import build_timelog_plan_sheet
from budget_planner.trends import build_trends_sheet
from budget_planner.layout import (
    SHEET_CONFIG,
    SHEET_PLAN,
    SHEET_CAPACITY,
    SHEET_ACTUALS,
    SHEET_ACTUALS_HOURS,
    SHEET_PLAN_FORECAST,
    SHEET_TIMELOG_PLAN,
    SHEET_TRENDS,
    SECTION_HEADER_LABELS,
    cell_ref,
)
from budget_planner.styles import (
    BLACK_FORMULA, HEADER_FONT, style_cell, apply_year_groups, sized_comment,
)


def _resolve_version_stamp() -> str:
    """Return a commit identifier for the running checkout.

    Uses ``git describe --always --dirty`` — so a clean commit yields the tag
    or short SHA, and an uncommitted tree appends ``-dirty`` (an audit-visible
    signal that the builder ran on in-progress code).

    Falls back to ``"unknown"`` when git is unavailable, the cwd isn't a
    checkout, or the subprocess fails for any reason.
    """
    try:
        result = subprocess.run(
            ["git", "describe", "--always", "--dirty"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        return "unknown"
    if result.returncode != 0:
        return "unknown"
    return result.stdout.strip() or "unknown"


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


@dataclass
class BuildContext:
    """Accumulated layout state threaded through ``_build_all_sheets``.

    Each builder contributes a layout dict that downstream builders consume;
    keeping them in one place (instead of named locals scattered through the
    orchestrator) makes the DAG legible and the build order explicit.
    """

    config_layout: dict[str, Any] | None = None
    plan_layout: dict[str, Any] | None = None
    capacity_layout: dict[str, Any] | None = None
    actuals_layout: dict[str, Any] | None = None
    summary_layout: dict[str, Any] | None = None


def _validate_config(
    project_params, work_packages, team, external_commitments, *, wp_schedule=None
):
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
    if rate <= 0:
        errors.append(f"FTE rate must be > 0, got {rate}")
    if not work_packages:
        errors.append("Work packages table is empty")
    if not team:
        errors.append("Team roster is empty")

    # Duplicate WP IDs in budget table (schedule section may repeat the same ID)
    seen_ids: set[str] = set()
    for wp in work_packages:
        if wp["id"] in seen_ids:
            errors.append(
                f"Duplicate WP ID in WP BUDGET table: {wp['id']}. "
                "Use one row per ID there; for multiple active periods use "
                "WP SCHEDULE (same ID, different Start/End), or split into "
                "distinct IDs (e.g. WP4a, WP4b)."
            )
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
                warnings.append(f"WP {wp['id']} has no schedule entry")

    return errors, warnings


def _build_all_sheets(
    wb,
    project_params,
    work_packages,
    team,
    external_commitments,
    *,
    grade_rates,
    wp_schedule=None,
    assignments=None,
    preserved_actuals=None,
    preserved_actuals_hours=None,
    preserved_timelog=None,
    preserved_eur_through_override: int | None = None,
    preserved_hours_through_override: int | None = None,
    reporting_periods=None,
    version_stamp=None,
):
    """Build all sheets into the workbook.

    The caller provides a fresh ``Workbook()``. Any sheets present before
    this function runs (e.g. openpyxl's default ``Sheet``) are treated as
    throwaway scaffolding and removed once the real sheets exist.

    Capacity monthly values are auto-distributed from assignments.
    Actuals month cells are always empty (or restored from preserved data).

    See module docstring for the build-order DAG. The ``BuildContext`` below
    accumulates layout dicts from each builder; downstream builders read from
    it so the dependencies are explicit.
    """
    _initial_sheets = list(wb.sheetnames)
    ctx = BuildContext()

    # Single fallback point: downstream builders (esp. build_summary_sheet)
    # require non-None data. The orchestrator owns the lazy-load so
    # individual builders can stay strict.
    if wp_schedule is None:
        from budget_planner.config import WP_SCHEDULE
        wp_schedule = WP_SCHEDULE
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS

    # Stage 1 — Config (no dependencies)
    config_ws = wb.create_sheet(SHEET_CONFIG)
    ctx.config_layout = build_config_sheet(
        config_ws,
        wb,
        project_params,
        work_packages,
        wp_schedule=wp_schedule,
        team=team,
        reporting_periods=reporting_periods,
        grade_rates=grade_rates,
        assignments=assignments,
        external_commitments=external_commitments,
        version_stamp=version_stamp,
    )

    # Stage 2 — Capacity (consumes config_layout)
    capacity_ws = wb.create_sheet(SHEET_CAPACITY)
    ctx.capacity_layout = build_capacity_sheet(
        capacity_ws,
        ctx.config_layout,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        team=team,
        external_commitments=external_commitments,
        assignments=assignments,
        config_first_wp_row=ctx.config_layout["wp_first_row"],
    )

    # Stage 3 — Plan (consumes config_layout)
    plan_ws = wb.create_sheet(SHEET_PLAN)
    ctx.plan_layout = build_plan_sheet(
        plan_ws,
        wb,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        config_layout=ctx.config_layout,
    )

    # Stage 4 — Rewrite Plan FTE as cross-sheet formulas from Capacity
    # (Plan cells written in Stage 3 get overwritten with SUM(Capacity!…).)
    _apply_planned_formulas(
        plan_ws,
        ctx.plan_layout,
        ctx.capacity_layout,
        work_packages=work_packages,
        project_params=project_params,
        wp_schedule=wp_schedule,
    )

    # Stage 5 — Trends stub: create the ``current_month`` named range so that
    # Actuals formulas (next stage) resolve. Full Trends content lands in
    # Stage 8 after its dependency layouts exist.
    trends_ws = wb.create_sheet(SHEET_TRENDS)
    trends_ws.sheet_properties.tabColor = "FFD700"
    style_cell(trends_ws, 1, 1, "Current Month:", font=HEADER_FONT, border=False)
    _cm_formula = (
        f"=MAX(1,MIN({SHEET_CONFIG}!B3,"
        f"(YEAR(TODAY())-YEAR(project_start))*12"
        f"+MONTH(TODAY())-MONTH(project_start)+1))"
    )
    style_cell(trends_ws, 1, 2, font=BLACK_FORMULA, border=False)
    trends_ws["B1"].value = _cm_formula
    trends_ws["B1"].comment = sized_comment("Auto-calculated from project start and today's date. "
        "Override with a static number to freeze.")
    wb.defined_names.add(
        DefinedName("current_month", attr_text=cell_ref(SHEET_TRENDS, 1, 2))
    )

    # Stage 5b — Timelog Plan (consumes config_layout + Trends-stub
    # ``current_month`` named range used by per-WP "Window months remaining").
    # Independent of Actuals/Hours/Forecast — built here so its tab can sit
    # right after Plan in the user-facing tab order (see move_sheet calls
    # at the end of this function).
    timelog_plan_ws = wb.create_sheet(SHEET_TIMELOG_PLAN)
    _preserved_grid = (preserved_timelog or {}).get("grid") or None
    _resolve_flag = (preserved_timelog or {}).get("resolve_flag", True)
    build_timelog_plan_sheet(
        timelog_plan_ws,
        wb,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        team=team,
        assignments=assignments or [],
        grade_rates=grade_rates,
        config_layout=ctx.config_layout,
        preserved_grid=_preserved_grid,
        resolve_flag=_resolve_flag,
    )

    # Stage 6a — Actuals-Hours (hour timelog; depends only on config_layout).
    # Built BEFORE Actuals so Plan-Forecast can consume its layout, and
    # Actuals can then consume Plan-Forecast's layout for the forecast
    # monitoring columns. Tab order is fixed up at the end of this function.
    actuals_hours_ws = wb.create_sheet(SHEET_ACTUALS_HOURS)
    actuals_hours_layout = build_actuals_hours_sheet(
        actuals_hours_ws,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        team=team,
        external_commitments=external_commitments,
        assignments=assignments,
        config_first_wp_row=ctx.config_layout["wp_first_row"],
        preserved=preserved_actuals_hours,
        wb=wb,
        hours_through_override=preserved_hours_through_override,
    )

    # Back-patch Capacity to-date block now that Actuals-Hours rows exist.
    apply_capacity_todate_formulas(
        capacity_ws,
        ctx.capacity_layout,
        actuals_hours_layout,
        project_params=project_params,
    )

    # Stage 6b — Plan-Forecast (aggregates forecast-EUR leaves from
    # Actuals-Hours into a WP × month grid; mirrors Plan shape).
    plan_forecast_ws = wb.create_sheet(SHEET_PLAN_FORECAST)
    plan_forecast_layout = build_plan_forecast_sheet(
        plan_forecast_ws,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        assignments=assignments,
        config_first_wp_row=ctx.config_layout["wp_first_row"],
        actuals_hours_layout=actuals_hours_layout,
    )

    # Stage 6c — Actuals (depends on plan_layout + Trends stub for
    # current_month + plan_forecast_layout for the forecast block).
    ctx.actuals_layout = build_actuals_sheet(
        wb,
        ctx.plan_layout,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        preserved=preserved_actuals,
        plan_forecast_layout=plan_forecast_layout,
        eur_through_override=preserved_eur_through_override,
    )

    # Stage 7 — Summary (needs plan_layout, actuals_layout, config_layout,
    # plan_forecast_layout for the Forecast EUR column per year)
    ctx.summary_layout = build_summary_sheet(
        wb,
        ctx.plan_layout,
        ctx.actuals_layout,
        project_params=project_params,
        work_packages=work_packages,
        reporting_periods=reporting_periods,
        plan_forecast_layout=plan_forecast_layout,
    )

    # Stage 8 — Trends full (needs every prior layout)
    build_trends_sheet(
        trends_ws,
        ctx.plan_layout,
        ctx.actuals_layout,
        ctx.capacity_layout,
        ctx.config_layout,
        project_params=project_params,
        work_packages=work_packages,
        wp_schedule=wp_schedule,
        team=team,
        reporting_periods=reporting_periods,
        plan_forecast_layout=plan_forecast_layout,
    )

    # Drop the scaffolding sheet(s) the caller handed us — the real sheets
    # are now in place, so the throwaway `Sheet` can go.
    for _name in _initial_sheets:
        if _name in wb.sheetnames:
            del wb[_name]

    # Tab order (user-facing) differs from build order.
    # After build the sheetlist is:
    #   Config, Capacity, Plan, Trends, Timelog Plan, Actuals-Hours,
    #   Plan-Forecast, Actuals, Summary
    # Target tab order:
    #   Config, Capacity, Plan, Timelog Plan, Actuals, Actuals-Hours,
    #   Plan-Forecast, Trends, Summary
    # 1) Push Trends past Timelog Plan + Actuals-Hours + Plan-Forecast +
    #    Actuals (+4).
    # 2) Hop Actuals leftward two slots so it sits just after Timelog Plan (-2).
    wb.move_sheet(SHEET_TRENDS, offset=4)
    wb.move_sheet(SHEET_ACTUALS, offset=-2)


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
            apply_year_groups(
                wb[sheet_name],
                3,
                project_params["months"],
                current_month,
                project_start=project_params["start"],
            )


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


def init_workbook(
    output_path: str = "output/budget_plan.xlsx",
    assignments: list[dict] | None = None,
) -> str:
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

    grade_rates = config.GRADE_RATES

    _build_all_sheets(
        wb,
        project_params,
        work_packages,
        team,
        external_commitments,
        wp_schedule=wp_schedule,
        assignments=assignments,
        reporting_periods=reporting_periods,
        grade_rates=grade_rates,
        version_stamp=_resolve_version_stamp(),
    )

    # _apply_protection(wb)  # Temporarily disabled
    from datetime import date
    from dateutil.relativedelta import relativedelta as _rd

    _today = date.today()
    _pstart = project_params["start"]
    _delta = _rd(_today, _pstart.date() if hasattr(_pstart, "date") else _pstart)
    _current_month = max(
        1, min(_delta.years * 12 + _delta.months + 1, project_params["months"])
    )
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
        read_config_sheet,
        read_team_from_config,
        read_capacity_sheet,
        read_actuals_values,
        read_actuals_hours_sheet,
        read_actuals_override,
        read_reporting_periods,
        read_grade_rates,
        read_assignments_from_config,
        read_externals_from_config,
        read_timelog_plan,
        _find_header_row,
    )

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File not found: {input_path}")

    # Print source file timestamp
    mod_time = datetime.fromtimestamp(os.path.getmtime(input_path))
    print(
        f"Reading from: {os.path.basename(input_path)} "
        f"(modified {mod_time.strftime('%Y-%m-%d %H:%M')})"
    )

    existing_wb = load_workbook(input_path, data_only=False)

    # ── Extract data ──

    # Config sheet → project params + WPs
    if SHEET_CONFIG not in existing_wb.sheetnames:
        existing_wb.close()
        raise ValueError(
            "Workbook has no Config sheet. "
            "Run 'python -m budget_planner init' to create a fresh workbook first."
        )
    project_params, work_packages, wp_schedule = read_config_sheet(
        existing_wb[SHEET_CONFIG]
    )

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
    has_assign_section = (
        _find_header_row(
            config_ws, SECTION_HEADER_LABELS["assignments"], required=False
        )
        is not None
    )
    has_ext_section = (
        _find_header_row(config_ws, SECTION_HEADER_LABELS["externals"], required=False)
        is not None
    )

    assignments = read_assignments_from_config(config_ws) if has_assign_section else []
    external_commitments = (
        read_externals_from_config(config_ws) if has_ext_section else []
    )

    if not has_assign_section or not has_ext_section:
        if SHEET_CAPACITY not in existing_wb.sheetnames:
            existing_wb.close()
            raise ValueError("Workbook has no Capacity sheet.")
        capacity_data = read_capacity_sheet(
            existing_wb[SHEET_CAPACITY],
            work_packages,
            project_months=project_params["months"],
        )
        if not has_assign_section:
            assignments = capacity_data["assignments"]
        if not has_ext_section:
            external_commitments = capacity_data["external_commitments"]

    # Preserve Actuals EUR values (always)
    preserved_actuals = None
    preserved_eur_through_override: int | None = None
    if SHEET_ACTUALS in existing_wb.sheetnames:
        preserved_actuals = read_actuals_values(
            existing_wb[SHEET_ACTUALS],
            project_params["months"],
        )
        preserved_eur_through_override = read_actuals_override(
            existing_wb[SHEET_ACTUALS]
        )

    # Preserve Actuals-Hours blue-input hours across rebuild.
    preserved_actuals_hours: dict[str, dict[str, list[float | None]]] | None = None
    preserved_hours_through_override: int | None = None
    if SHEET_ACTUALS_HOURS in existing_wb.sheetnames:
        preserved_actuals_hours = read_actuals_hours_sheet(
            existing_wb[SHEET_ACTUALS_HOURS],
            project_months=project_params["months"],
        )
        preserved_hours_through_override = read_actuals_override(
            existing_wb[SHEET_ACTUALS_HOURS]
        )

    # Preserve Timelog Plan grid values + resolve flag (v2 surface).
    # Missing sheet → {"grid": {}, "resolve_flag": True} (pre-v2 workbook;
    # builder still renders and the solver will populate the grid).
    preserved_timelog: dict = read_timelog_plan(
        existing_wb, project_params["months"]
    )

    # Calculate current_month from project start and today's date
    from datetime import date
    from dateutil.relativedelta import relativedelta as _rd

    _today = date.today()
    _pstart = project_params["start"]
    _delta = _rd(_today, _pstart.date() if hasattr(_pstart, "date") else _pstart)
    current_month = max(
        1, min(_delta.years * 12 + _delta.months + 1, project_params["months"])
    )

    existing_wb.close()

    # ── Validate ──
    errors, warnings = _validate_config(
        project_params,
        work_packages,
        team,
        external_commitments,
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

    # Prune orphaned Timelog Plan grid cells — a person or WP code removed
    # from config can't round-trip meaningfully.
    _timelog_grid = preserved_timelog.get("grid") if preserved_timelog else None
    if _timelog_grid:
        for key in list(_timelog_grid.keys()):
            person, _month = key
            code = _timelog_grid[key]
            if person not in team_names:
                warnings.append(
                    f"Orphaned Timelog Plan cell: person '{person}' removed from team"
                )
                del _timelog_grid[key]
            elif code not in wp_ids:
                warnings.append(
                    f"Orphaned Timelog Plan cell: WP '{code}' removed (was set for {person})"
                )
                del _timelog_grid[key]

    # Prune orphaned hours (person or WP removed from config). Externals
    # stay keyed by project name — validate against the external list.
    if preserved_actuals_hours:
        ext_projects_by_person: dict[str, set[str]] = {}
        for e in external_commitments:
            ext_projects_by_person.setdefault(e["person"], set()).add(e["project"])
        for person in list(preserved_actuals_hours.keys()):
            if person not in team_names:
                warnings.append(f"Orphaned Hours: person '{person}' removed from team")
                del preserved_actuals_hours[person]
                continue
            per_wp = preserved_actuals_hours[person]
            person_exts = ext_projects_by_person.get(person, set())
            for key in list(per_wp.keys()):
                if key not in wp_ids and key not in person_exts:
                    warnings.append(
                        f"Orphaned Hours: '{key}' removed for {person}"
                    )
                    del per_wp[key]

    if errors:
        raise ValueError(
            "Rebuild validation failed:\n" + "\n".join(f"  ERROR: {e}" for e in errors)
        )

    for w in warnings:
        print(f"  WARNING: {w}")

    # ── Generate new workbook ──
    wb = Workbook()

    _build_all_sheets(
        wb,
        project_params,
        work_packages,
        team,
        external_commitments,
        wp_schedule=wp_schedule,
        assignments=assignments,
        preserved_actuals=preserved_actuals,
        preserved_actuals_hours=preserved_actuals_hours,
        preserved_timelog=preserved_timelog,
        preserved_eur_through_override=preserved_eur_through_override,
        preserved_hours_through_override=preserved_hours_through_override,
        reporting_periods=reporting_periods,
        grade_rates=grade_rates,
        version_stamp=_resolve_version_stamp(),
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
    """Seed Actuals EUR and Actuals-Hours, plus the EUR through-month override.

    Capacity FTE values are now auto-distributed by the builder via
    assignments — this function seeds Actuals EUR spend and per-person
    Actuals-Hours so the reconciliation visuals (greyed forecast-EUR
    band, Capacity to-date block) have something to render. Writes a
    blue-input override to ``Actuals!$B$1`` at DEMO_CURRENT_MONTH − 3
    to model the accountant's 3-month finance-close lag (EUR cadence
    trails hours). Hours through-month is left auto-derived — the
    row-1 scan on Actuals-Hours naturally returns DEMO_CURRENT_MONTH
    because hours are seeded through that month.

    Args:
        workbook_path: Path to an existing workbook (from init_workbook).

    Returns:
        The workbook path.
    """
    from openpyxl import load_workbook as _load_wb
    from budget_planner.config import (
        DEMO_CURRENT_MONTH,
        DEMO_ACTUALS_MULTIPLIERS,
        DEMO_CAPACITY_ASSIGNMENTS,
        TEAM,
        WORK_PACKAGES,
        WP_SCHEDULE,
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

    # ── Seed EUR through-month override on Actuals!$B$1 ──
    # Models a 3-month finance-close lag (EUR cadence trails hours). Hours
    # through-month resolves via the auto-derived scan on Actuals-Hours;
    # no override needed because hours are seeded through DEMO_CURRENT_MONTH.
    # See design doc
    # ``docs/plans/archived/2026-04-18-config-through-month-cleanup-design.md``.
    hours_through = DEMO_CURRENT_MONTH
    act_ws.cell(row=1, column=2, value=max(1, DEMO_CURRENT_MONTH - 3))

    # ── Seed Actuals-Hours per-person monthly hours ──
    # Reconstruct per-person WP ordering from DEMO_CAPACITY_ASSIGNMENTS
    # (dedup by person+wp preserving first occurrence) — mirrors the
    # order actuals_hours.py builds rows in.
    from budget_planner.layout import wp_active_months as _wp_active_months
    HOURS_PER_FTE_MONTH = 143
    per_person: dict[str, list[dict]] = {}
    seen: set[tuple[str, str]] = set()
    for asgn in DEMO_CAPACITY_ASSIGNMENTS:
        key = (asgn["person"], asgn["wp"])
        if key in seen:
            continue
        seen.add(key)
        per_person.setdefault(asgn["person"], []).append(asgn)

    team_names = {m["name"] for m in TEAM}
    wp_index = {w["id"]: i for i, w in enumerate(WORK_PACKAGES)}
    wp_active_set = {
        w["id"]: _wp_active_months(w["id"], WP_SCHEDULE)
        for w in WORK_PACKAGES
    }

    if SHEET_ACTUALS_HOURS in wb.sheetnames:
        hours_ws = wb[SHEET_ACTUALS_HOURS]
        name_to_header_row: dict[str, int] = {}
        for r in range(1, hours_ws.max_row + 1):
            v = hours_ws.cell(r, 1).value
            if isinstance(v, str) and v in team_names:
                name_to_header_row[v] = r

        for member in TEAM:
            name = member["name"]
            header_row = name_to_header_row.get(name)
            if header_row is None:
                continue
            for i, asgn in enumerate(per_person.get(name, [])):
                wp_id = asgn["wp"]
                if wp_id not in wp_index:
                    continue
                row = header_row + 1 + i
                wp_start, wp_end = wp_month_bounds(wp_id, WP_SCHEDULE)
                a_start = max(wp_start, asgn["start"])
                a_end = min(wp_end, asgn["end"])
                active_months = a_end - a_start + 1
                if active_months <= 0:
                    continue
                monthly_fte = asgn["total_fte"] / active_months
                multiplier = DEMO_ACTUALS_MULTIPLIERS[
                    wp_index[wp_id] % len(DEMO_ACTUALS_MULTIPLIERS)
                ]
                active_set = wp_active_set[wp_id]
                for m_idx in range(a_start - 1, min(a_end, hours_through)):
                    month_num = m_idx + 1
                    if month_num not in active_set:
                        continue
                    col = MONTH_COL_START + m_idx
                    hours = round(monthly_fte * HOURS_PER_FTE_MONTH * multiplier)
                    hours_ws.cell(row, col).value = hours

    wb.save(workbook_path)
    wb.close()
    print(f"Demo data seeded: {workbook_path}")
    return workbook_path
