"""Workbook orchestrator — init, rebuild, snapshot/restore modes.

## Builder interface

Each sheet is produced by a stateless builder function. Builders receive
data + upstream layout dicts as parameters and return a layout dict of
row/column positions consumed by downstream builders::

    def build_<sheet>_sheet(ws, wb, ..., *, project_params=None,
                            work_packages=None, ...):
        # ws: openpyxl Worksheet
        # wb: parent Workbook (named ranges, cross-sheet access)
        # ...: upstream layout dicts
        # Returns: layout dict with row/column numbers for downstream builders

## Builder dependency chain

| Builder        | Needs from upstream                                                                                             |
|----------------|-----------------------------------------------------------------------------------------------------------------|
| config_sheet   | (none)                                                                                                          |
| plan           | config_layout.wp_first_row                                                                                      |
| capacity       | config_layout.person_rows, plan_layout                                                                          |
| plan (patch)   | capacity_layout.person_groups.wp_rows                                                                           |
| timelog_plan   | config_layout + Trends-stub ``current_month`` named range + ``team_priorities`` / ``roster_lookup`` named ranges |
| actuals        | plan_layout                                                                                                     |
| summary        | plan_layout + actuals_layout                                                                                    |
| trends         | all upstream layouts                                                                                            |

## Build order (load-bearing — downstream builders consume layout dicts
produced by upstream ones, so reordering silently breaks cross-sheet
formulas)::

    Config ──► Capacity ──► Plan ──► (rewrite Plan FTE)
                               │
                               ├──► Trends(stub = current_month named range)
                               │       │
                               │       └──► Timelog Plan
                               └──► Actuals-Hours
                                       └──► Actuals ──► Summary ──► Trends(full)

Enforced by ``test_build_order_is_stable``. Build order differs from
user-facing tab order — some builders are ordered by dependency, not
display position. Tab order (set after build via ``wb.move_sheet``) is in
``docs/spec.md`` § Cross-sheet formula graph.

### Build-order rationale

- Capacity consumes ``config_layout`` (param rows, first WP row).
- Plan consumes ``config_layout`` (param rows, first WP row).
- ``_apply_planned_formulas`` overwrites Plan FTE cells using
  ``capacity_layout`` — must run after both Capacity and Plan.
- Trends stub creates the ``current_month`` named range that Actuals
  formulas depend on — must run before Actuals.
- Timelog Plan is independent of the forecast chain; it only needs the
  Config layout and the ``current_month`` stub, so it slots right after
  Trends-stub.
- Trends is populated last so every upstream layout dict (including
  ``helper_row_forecast``) is known.

## Sheet role contract (one-line)

Each sheet has a single role; derived views must not write to input
surfaces and vice versa. Roles are encoded in the font-color rule
(blue=input, black=formula, green=cross-sheet ref) plus the tab-color
markers in ``docs/spec.md`` § Tab colors.

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
from budget_planner.summary import build_summary_sheet
from budget_planner.timelog_plan import build_timelog_plan_sheet
from budget_planner.trends import build_trends_sheet
from budget_planner.layout import (
    SHEET_CONFIG,
    SHEET_PLAN,
    SHEET_CAPACITY,
    SHEET_ACTUALS,
    SHEET_ACTUALS_HOURS,
    SHEET_TIMELOG_PLAN,
    SHEET_TRENDS,
    SECTION_HEADER_LABELS,
    cell_ref,
)
from budget_planner.styles import (
    BLACK_FORMULA, HEADER_FONT, style_cell, apply_year_groups, sized_comment,
)
from budget_planner import config
# Re-exported for backwards compatibility — actual implementation in demo_seed.
from budget_planner.demo_seed import seed_demo_data as seed_demo_data  # noqa: F401


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
    project_params, work_packages, team, external_commitments
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

    # Duplicate WP IDs in budget table — each WP must appear exactly once.
    seen_ids: set[str] = set()
    for wp in work_packages:
        if wp["id"] in seen_ids:
            errors.append(
                f"Duplicate WP ID in WP BUDGET table: {wp['id']}. "
                "Each WP must appear exactly once. If you need separable "
                "phases, split into distinct IDs (e.g. WP4a, WP4b)."
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
        start = wp.get("start")
        end = wp.get("end")
        if start is not None and end is not None and start > end:
            warnings.append(
                f"WP {wp['id']}: start ({start}) > end ({end})"
            )

    return errors, warnings


def _build_all_sheets(
    wb,
    project_params,
    work_packages,
    team,
    external_commitments,
    *,
    grade_rates,
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
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS

    # Stage 1 — Config (no dependencies)
    # `Rebuild from month #` (Config!B9) is BLUE_INPUT — preserve the user's
    # last value across rebuilds rather than resetting to 1, otherwise the
    # second rebuild silently collapses the freeze point even though the
    # first rebuild's solver honored the value.
    _config_rebuild_from_month = (preserved_timelog or {}).get("rebuild_from_month", 1)
    config_ws = wb.create_sheet(SHEET_CONFIG)
    ctx.config_layout = build_config_sheet(
        config_ws,
        wb,
        project_params,
        work_packages,
        team=team,
        reporting_periods=reporting_periods,
        grade_rates=grade_rates,
        assignments=assignments,
        external_commitments=external_commitments,
        version_stamp=version_stamp,
        rebuild_from_month=_config_rebuild_from_month,
    )

    # Stage 2 — Capacity (consumes config_layout)
    capacity_ws = wb.create_sheet(SHEET_CAPACITY)
    ctx.capacity_layout = build_capacity_sheet(
        capacity_ws,
        ctx.config_layout,
        project_params=project_params,
        work_packages=work_packages,
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
    # ``current_month`` named range used by per-WP "Burn rate needed" wallclock
    # denominator). Independent of Actuals/Hours/Forecast — built here so its
    # tab can sit right after Config in the user-facing tab order (see
    # move_sheet calls at the end of this function).
    timelog_plan_ws = wb.create_sheet(SHEET_TIMELOG_PLAN)
    # `preserved_timelog["grid"]` was pruned in `rebuild_workbook` before
    # reaching here. Cells in months < rebuild_from_month are frozen and kept
    # verbatim by the builder; cells ≥ rebuild_from_month are discarded and
    # the solver re-derives them. See `docs/spec.md` § Timelog Plan merged grid.
    _preserved_grid = (preserved_timelog or {}).get("grid") or None
    _rebuild_from_month = (preserved_timelog or {}).get("rebuild_from_month", 1)
    assert ctx.config_layout is not None
    build_timelog_plan_sheet(
        timelog_plan_ws,
        wb,
        project_params=project_params,
        work_packages=work_packages,
        team=team,
        assignments=assignments or [],
        grade_rates=grade_rates,
        config_layout=ctx.config_layout,
        preserved_grid=_preserved_grid,
        rebuild_from_month=_rebuild_from_month,
    )

    # Stage 6a — Actuals-Hours (hour timelog; depends only on config_layout).
    # Tab order is fixed up at the end of this function.
    actuals_hours_ws = wb.create_sheet(SHEET_ACTUALS_HOURS)
    actuals_hours_layout = build_actuals_hours_sheet(
        actuals_hours_ws,
        project_params=project_params,
        work_packages=work_packages,
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

    # Stage 6b — Actuals (depends on plan_layout + Trends stub for
    # current_month).
    ctx.actuals_layout = build_actuals_sheet(
        wb,
        ctx.plan_layout,
        project_params=project_params,
        work_packages=work_packages,
        preserved=preserved_actuals,
        eur_through_override=preserved_eur_through_override,
    )

    # Stage 7 — Summary (needs plan_layout, actuals_layout, config_layout;
    # consumes Actuals!{posted_idx,posted_value,slope} for the project-end
    # Forecast EUR column).
    ctx.summary_layout = build_summary_sheet(
        wb,
        ctx.plan_layout,
        ctx.actuals_layout,
        project_params=project_params,
        work_packages=work_packages,
        reporting_periods=reporting_periods,
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
        team=team,
        reporting_periods=reporting_periods,
    )

    # Drop the scaffolding sheet(s) the caller handed us — the real sheets
    # are now in place, so the throwaway `Sheet` can go.
    for _name in _initial_sheets:
        if _name in wb.sheetnames:
            del wb[_name]

    # Tab order (user-facing) differs from build order. Target positions:
    #   0: Config, 1: Timelog Plan, 2: Actuals-Hours  (the frequently-edited
    #   tabs sit up-front). Remaining sheets keep their relative build order.
    # Idempotent: computed offsets collapse to 0 when already in place.
    for target_idx, name in enumerate(
        [SHEET_TIMELOG_PLAN, SHEET_ACTUALS_HOURS], start=1,
    ):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=target_idx - wb.sheetnames.index(name))


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
    project_params = {
        "title": config.PROJECT_TITLE,
        "start": config.PROJECT_START,
        "months": config.PROJECT_MONTHS,
        "rate": config.DEFAULT_RATE,
        "eu_ceiling_pm": config.EU_ANNUAL_CEILING_PM,
    }
    work_packages = config.WORK_PACKAGES
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
    project_params, work_packages = read_config_sheet(
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

    # Preserve Timelog Plan grid values + rebuild-from-month N (freeze-point).
    # Missing sheet → {"grid": {}, "rebuild_from_month": 1} (pre-freeze-point
    # workbook; builder still renders and the solver will populate the grid).
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
