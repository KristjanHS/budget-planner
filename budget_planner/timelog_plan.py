"""Timelog Plan sheet builder — prescriptive 1.0-FTE WP allocation grid.

A separate allocation surface from the Plan sheet. For each (person, month)
the sheet prescribes which WP they should log 100% against.

Layout (v2 + 2026-04-19 enhancements — supersedes v1's PROPOSED + ASSIGNED twin bands):

  1. Banner (row 1)
  2. Traffic lights (row 2)
  3. blank spacer
  4. PER-WP SUMMARY band (9 base cols + one EUR col per calendar year)
     blank spacer
  5. TIMELOG GRID — single editable band, literal WP codes + empty strings
     blank spacer
  6. PER-PERSON SUMMARY band (9 cols: Person + 6 metrics + Rate -i + Rate +i)

Each data band is wrapped in an Excel row-group (``outline_level=1``) so the
user can fold away sections not under review. Section-title rows stay at
level 0 and remain visible when the group collapses.

Col A is Person (cross-ref into Config Team Roster) across the grid and the
PER-PERSON SUMMARY band. No per-row metric column — the merged grid starts
at col B (TIMELOG_MONTH_COL_START). The per-person Rate helpers live on the
PER-PERSON SUMMARY band: col H is Rate -i (VLOOKUP into roster_lookup, the
canonical -i source) and col I is Rate +i (= H × 1.25). PER-WP SUMMARY
Assigned-€ SUMPRODUCT consumes the +i col so its result lands in +i units,
matching Budget +i in col B.

Cross-sheet invariants: ``docs/spec.md`` § Timelog Plan merged grid.

Stateless builder — receives all config via kwargs from the orchestrator. Per
``.claude/rules/python-refactors.md``: no ``from budget_planner import config``
fallback inside this builder; the orchestrator owns the lazy-load.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass

from dateutil.relativedelta import relativedelta
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from budget_planner.layout import (
    COL_WP_BUDGET_PLUS_I,
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
    TIMELOG_WP_SUMMARY_BASE_LAST_COL,
    TIMELOG_YEARLY_LABEL_COL,
    cell_ref,
    month_range_ref,
)
from budget_planner.styles import (
    BLACK_FORMULA,
    BLUE_INPUT,
    CENTER,
    CF_GREEN_FILL,
    CF_PASTEL_FILLS,
    CF_RED_FILL,
    DARK_GREY_FILL,
    GREEN_XREF,
    GREY_FILL,
    HEADER_FONT,
    LIGHT_BLUE_FILL,
    SECTION_FILL,
    SECTION_FONT,
    TITLE_FONT,
    VACATION_FILL,
    WITH_INDIRECT_HEADER_FILL,
    WITHOUT_INDIRECT_HEADER_FILL,
    col_letter,
    sized_comment,
    style_cell,
)
from budget_planner.timelog_solver import (
    Person,
    PersonWPWindow,
    WP,
    classify_wp_stop_reasons,
    solve,
)

# Section-title strings — exported so the reader can anchor on the grid title
# via the same literal the builder writes. Update here, not at the callsite.
SECTION_TITLE_WP_SUMMARY = "PER-WP SUMMARY"
SECTION_TITLE_GRID = (
    "TIMELOG GRID (edit freely — rebuild re-solves from Config!B9 onward)"
)
SECTION_TITLE_PERSON_SUMMARY = "PER-PERSON SUMMARY"

# PER-PERSON SUMMARY layout: col A = Person, cols B..I = 8 metric cols.
# Metric headers start at col B. The two rate cols sit at the right edge:
# col H = "Rate (€/PM) -i" (VLOOKUP from roster_lookup, the canonical -i
# source); col I = "Rate (€/PM) +i" (= H × 1.25). Assigned-€ SUMPRODUCT on
# PER-WP SUMMARY consumes the +i col so its result is +i, matching Budget +i.
# The Monthly total -i row consumes the -i col directly.
_PERSON_SUMMARY_METRIC_HEADERS = [
    "Assigned", "Pool", "Over-pool?",
    "Switches", "First log", "Last log",
    "Rate (€/PM) -i", "Rate (€/PM) +i",
]
_PERSON_SUMMARY_FIRST_METRIC_COL = 2   # col B (Assigned)
_PERSON_SUMMARY_RATE_MINUS_I_COL = (
    _PERSON_SUMMARY_FIRST_METRIC_COL + len(_PERSON_SUMMARY_METRIC_HEADERS) - 2
)  # col H
_PERSON_SUMMARY_RATE_PLUS_I_COL = (
    _PERSON_SUMMARY_FIRST_METRIC_COL + len(_PERSON_SUMMARY_METRIC_HEADERS) - 1
)  # col I
# Backwards-compat alias: existing call sites that consumed the single-rate
# column treated it as -i. Keep the alias pointing at -i so Monthly total
# rows + burn-rate CF stay anchored on -i (their math depends on it). The
# Assigned-€ rate-vector (the only consumer that needs +i) targets
# _PERSON_SUMMARY_RATE_PLUS_I_COL explicitly.
_PERSON_SUMMARY_RATE_COL = _PERSON_SUMMARY_RATE_MINUS_I_COL

# PER-WP SUMMARY metric columns live at A..I (WP code + 8 metrics).
# Phase 3 appends dynamic "Y{yyyy} +i" cols right of col I — one per calendar
# year spanned by the project. See `_year_col_specs`. Aiko-port T3 inserted
# Buffer % at col E; Assigned months/WP sched months/Burn rate/OOW shifted
# one col right (E→F, F→G, G→H, H→I).
_WP_SUMMARY_BASE_HEADERS = [
    "WP code", "Budget +i", "Assigned €", "Buffer €", "Buffer %",
    "Assigned months", "WP sched months", "Burn rate needed",
    "Any out-of-window?", "Stop reason",
]
_WP_SUMMARY_BASE_LAST_COL = len(_WP_SUMMARY_BASE_HEADERS)  # 10
# Drift guard: layout.py exposes this boundary publicly for tests and future
# cross-module callers; the header-count derivation here must stay in sync.
assert _WP_SUMMARY_BASE_LAST_COL == TIMELOG_WP_SUMMARY_BASE_LAST_COL

_LEGACY_RESOLVE_FLAG_NAMED_RANGE = "timelog_resolve_flag"


# ---------------------------------------------------------------------------
# Public builder
# ---------------------------------------------------------------------------


def build_timelog_plan_sheet(
    ws: Worksheet,
    wb: Workbook,
    *,
    project_params: dict,
    work_packages: list[dict],
    team: list[dict],
    assignments: list[dict],
    grade_rates: Mapping[str, float],
    config_layout: dict,
    preserved_grid: Mapping[tuple[str, int], str] | None = None,
    rebuild_from_month: int = 1,
) -> dict:
    """Build the Timelog Plan sheet with merged grid + summary bands.

    Args:
        ws: Worksheet to populate (created by orchestrator, named "Timelog Plan").
        wb: Parent workbook — legacy ``timelog_resolve_flag`` defined name is
            removed here if present (no BC shim).
        project_params: dict with keys 'title', 'start', 'months', 'rate', etc.
        work_packages: list of WP dicts with keys 'id', 'name', 'fte', 'rate',
            'start', 'end', 'type' (single contiguous active range).
        team: list of team-member dicts with keys 'name', 'max_fte', 'grade',
            'priority' (1-3, defaults to 2 from the reader).
        assignments: list of person-WP assignment dicts (person, wp,
            total_fte, start, end). Treated as the solver's hard
            (person, wp) → window constraints.
        grade_rates: dict grade_name → EUR/PM. Used by the solver for
            €-accounting: each Person's ``rate_eur_per_pm`` is resolved
            from this mapping via the roster's grade.
        config_layout: dict returned by build_config_sheet() — must include
            ``param_rows`` (incl. "Rebuild from month #"), ``wp_first_row``,
            ``person_rows``.
        preserved_grid: Optional mapping ``(person_name, month_num)`` →
            WP code. Cells whose ``month_num`` is strictly below
            ``rebuild_from_month`` are treated as FROZEN: kept verbatim
            and subtracted from WP ceilings + Person pools before solving.
            Cells at or above are discarded (solver re-derives them).
        rebuild_from_month: 1-indexed month boundary. Cells in months
            ``< rebuild_from_month`` are frozen (preserved from
            ``preserved_grid``); cells from ``rebuild_from_month`` onward
            are re-solved. Default ``1`` = solve-all (pre-v2 migration and
            init path). A value > project_months preserves every cell.

    Returns:
        Layout dict with keys:
          sheet, month_col_start, month_col_end,
          grid_person_first_row, grid_person_last_row,
          summary_person_first_row, summary_person_last_row,
          wp_summary_first_row, wp_summary_last_row,
          person_rows_by_band (dict band → {name: row}),
          rebuild_from_month_cell (Config!B{row} coordinate string).
    """
    project_months: int = project_params["months"]
    project_start = project_params["start"]
    vacation_month_num = int(project_params.get("vacation_month_num", 7))
    vacation_months_project = _vacation_months_in_span(
        vacation_month_num, project_start, project_months,
    )

    # ── Tab appearance ──
    ws.sheet_properties.tabColor = "FF8C00"  # orange — distinguishes from existing tabs

    # ── Column geometry ──
    month_col_start = TIMELOG_MONTH_COL_START
    month_col_end = month_col_start + project_months - 1
    canvas_last_col = month_col_end

    # ── Solver input prep ──
    # Freeze semantics: months strictly below ``rebuild_from_month`` are
    # preserved verbatim from ``preserved_grid`` and their € / pool usage
    # is subtracted from WP ceilings + Person pools before solving. See
    # docs/spec.md § Timelog Plan merged grid.
    freeze_boundary = max(1, int(rebuild_from_month))
    frozen_months: frozenset[int] = frozenset(range(1, freeze_boundary))

    # Partition preserved_grid: keep only frozen-month, non-vacation cells.
    # Vacation-month entries are dropped defensively (vacation cells are
    # always empty per invariant; mirrors _write_grid_cell's migration).
    frozen_placements: dict[str, dict[int, str]] = {}
    for (pname, mi), wp_code in (preserved_grid or {}).items():
        if mi in frozen_months and mi not in vacation_months_project:
            frozen_placements.setdefault(pname, {})[mi] = wp_code

    sorted_team = sorted(
        team, key=lambda m: (int(m.get("priority", 2)), m["name"])
    )
    rate_by_person: dict[str, int] = {
        m["name"]: int(round(grade_rates[m["grade"]]))
        for m in sorted_team
    }
    # Per-person frozen-month count shrinks the logging pool.
    frozen_count_by_person: dict[str, int] = {
        pname: len(month_wp) for pname, month_wp in frozen_placements.items()
    }
    # Per-WP frozen €-consumption narrows the solver's fill target. Negative
    # residuals (frozen > budget) surface as negative Buffer € in the per-WP
    # summary — that cell reads Config!E directly, so the full ceiling is
    # still displayed; only the solver sees the narrowed remainder.
    frozen_eur_by_wp: dict[str, int] = {}
    for pname, month_wp in frozen_placements.items():
        rate = rate_by_person.get(pname, 0)
        for wp_code in month_wp.values():
            frozen_eur_by_wp[wp_code] = (
                frozen_eur_by_wp.get(wp_code, 0) + rate
            )

    persons = [
        Person(
            name=m["name"],
            priority=int(m.get("priority", 2)),
            pool=max(
                0,
                round(
                    float(m["max_fte"])
                    * (project_months - len(vacation_months_project))
                )
                - frozen_count_by_person.get(m["name"], 0),
            ),
            rate_eur_per_pm=int(round(grade_rates[m["grade"]])),
            vacation_months=vacation_months_project,
            frozen_months=frozen_months,
        )
        for m in sorted_team
    ]
    wps_full = _build_wps(work_packages, project_params)
    wps = [
        WP(
            code=w.code,
            active_start=w.active_start,
            active_end=w.active_end,
            ceiling_eur=max(
                0, w.ceiling_eur - frozen_eur_by_wp.get(w.code, 0)
            ),
        )
        for w in wps_full
    ]
    windows = [
        PersonWPWindow(
            person_name=a["person"],
            wp_code=a["wp"],
            start_month=int(a["start"]),
            end_month=int(a["end"]),
        )
        for a in (assignments or [])
    ]
    assigned_map = solve(persons, wps, windows, project_months)
    # Overlay frozen placements — freeze wins in frozen months. The solver
    # is already guarded against touching frozen months, so this is a no-op
    # for correctly-behaving solvers; keep as a belt-and-suspenders write.
    for pname, month_wp in frozen_placements.items():
        for mi, wp_code in month_wp.items():
            assigned_map.setdefault(pname, {})[mi] = wp_code
    stop_reasons = classify_wp_stop_reasons(
        assigned_map, persons, wps, windows, project_months,
    )
    person_windows: dict[str, list[tuple[str, int, int]]] = {}
    for w in windows:
        person_windows.setdefault(w.person_name, []).append(
            (w.wp_code, w.start_month, w.end_month)
        )
    num_persons = len(persons)
    num_wps = len(wps)

    # ── Row layout ──
    banner_row = 1
    tl_row = 2
    # row 3: blank spacer

    wp_summary_title_row = 4
    wp_summary_header_row = 5
    wp_summary_first_row = 6
    wp_summary_last_row = (
        wp_summary_first_row + num_wps - 1 if num_wps > 0
        else wp_summary_header_row
    )
    # Yearly Assigned +i / Yearly Budget +i rows sit directly below the last WP row
    # (no spacer, same outline group). Only emitted when there are WP rows.
    yearly_total_row = wp_summary_last_row + 1 if num_wps > 0 else None
    yearly_budget_row = wp_summary_last_row + 2 if num_wps > 0 else None
    wp_summary_band_end_row = (
        yearly_budget_row if yearly_budget_row is not None
        else wp_summary_last_row
    )

    grid_title_row = wp_summary_band_end_row + 2  # +1 spacer

    grid_year_row = grid_title_row + 1
    grid_month_row = grid_year_row + 1
    grid_midx_row = grid_month_row + 1
    grid_person_first_row = grid_midx_row + 1
    grid_person_last_row = (
        grid_person_first_row + num_persons - 1 if num_persons > 0
        else grid_person_first_row - 1
    )
    # Monthly total -i row sits directly under the last person row (no spacer,
    # same outline group). Only emitted when there are person rows.
    grid_total_row = grid_person_last_row + 1 if num_persons > 0 else None
    # Monthly total +i row sits directly below the -i row, outside the grid
    # outline group so it stays visible when the grid collapses. A blank
    # spacer row separates it from PER-PERSON SUMMARY below.
    grid_person_cost_row = (
        grid_total_row + 1 if grid_total_row is not None else None
    )

    summary_title_row = (
        (grid_total_row if grid_total_row is not None else grid_midx_row) + 3
    )
    summary_header_row = summary_title_row + 1
    summary_person_first_row = summary_header_row + 1
    summary_person_last_row = (
        summary_person_first_row + num_persons - 1 if num_persons > 0
        else summary_person_first_row - 1
    )

    person_rows_by_band: dict[str, dict[str, int]] = {
        "grid": {
            p.name: grid_person_first_row + i for i, p in enumerate(persons)
        },
        "summary": {
            p.name: summary_person_first_row + i for i, p in enumerate(persons)
        },
    }

    # ── Row 1: title banner (merged across canvas) ──
    ws.merge_cells(
        start_row=banner_row, start_column=1,
        end_row=banner_row, end_column=canvas_last_col,
    )
    style_cell(
        ws, banner_row, 1,
        value="TIMELOG PLAN — prescriptive 1.0-FTE WP allocations",
        font=TITLE_FONT, fill=SECTION_FILL, alignment=CENTER, border=False,
    )
    for c in range(2, canvas_last_col + 1):
        style_cell(ws, banner_row, c, fill=SECTION_FILL, border=False)
    ws.cell(row=banner_row, column=1).font = Font(
        name="Arial", bold=True, size=14, color="FFFFFF"
    )

    # ── Row 2: traffic-light labels + toggle; values patched in below ──
    tl_label_cols = [1, 3, 5, 7]
    tl_value_cols = [2, 4, 6, 8]
    tl_labels = [
        "Any undershoot?",
        "Any out-of-window cell?",
        "Any over-pool person?",
        "Total switches",
    ]
    for c, lbl in zip(tl_label_cols, tl_labels):
        style_cell(ws, tl_row, c, value=lbl, font=HEADER_FONT, alignment=CENTER)
    for c in tl_value_cols:
        style_cell(ws, tl_row, c, font=BLACK_FORMULA, alignment=CENTER, fmt="0")
    # Toggle-row legacy cells (cols 9-10) are no longer written. The rebuild
    # control is now the Config!B{Rebuild from month #} integer param (see
    # config_sheet.py + docs/spec.md § Timelog Plan merged grid).
    # Any pre-v2 workbook carrying the old ``timelog_resolve_flag`` defined
    # name is actively unregistered on rebuild — no BC shim.
    if _LEGACY_RESOLVE_FLAG_NAMED_RANGE in wb.defined_names:
        del wb.defined_names[_LEGACY_RESOLVE_FLAG_NAMED_RANGE]

    # row 3: blank spacer (no write)

    # ── PER-WP SUMMARY band ──
    year_specs = _year_col_specs(
        project_params["start"], project_months, month_col_start,
    )
    wp_summary_last_col = _WP_SUMMARY_BASE_LAST_COL + len(year_specs)
    _write_section_title(
        ws, wp_summary_title_row, SECTION_TITLE_WP_SUMMARY,
        last_col=wp_summary_last_col,
    )
    wp_summary_headers = _WP_SUMMARY_BASE_HEADERS + [
        f"Y{spec.year} +i" for spec in year_specs
    ]
    for i, hdr in enumerate(wp_summary_headers):
        # Year-col headers (i ≥ _WP_SUMMARY_BASE_LAST_COL) carry
        # WITH_INDIRECT_HEADER_FILL — per-WP year cells display +i
        # via ×1.25 wrap at the SUMPRODUCT leaf (T7.3). Among the base
        # headers, only "Budget +i" (col B, retargeted to Config!F by
        # Aiko-port Q-B) carries the +i token; it gets the same green
        # fill so the T9 header-fill invariant holds.
        is_year_col = i >= _WP_SUMMARY_BASE_LAST_COL
        carries_plus_i = hdr.endswith("+i")
        style_cell(
            ws, wp_summary_header_row, 1 + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER,
            fill=(WITH_INDIRECT_HEADER_FILL
                  if is_year_col or carries_plus_i else None),
        )
    cfg_wp_first_row = config_layout["wp_first_row"]
    # Rate vector for Assigned-€ targets the +i col so the SUMPRODUCT result
    # is +i (matches Budget +i in col B). Monthly total rows below construct
    # their own ranges off the -i col directly.
    rate_col_range = (
        f"{get_column_letter(_PERSON_SUMMARY_RATE_PLUS_I_COL)}{summary_person_first_row}"
        f":{get_column_letter(_PERSON_SUMMARY_RATE_PLUS_I_COL)}{summary_person_last_row}"
    ) if num_persons > 0 else ""
    grid_range = _grid_range_ref(
        month_col_start, month_col_end,
        grid_person_first_row, grid_person_last_row,
    ) if num_persons > 0 else ""
    midx_range = (
        f"{get_column_letter(month_col_start)}{grid_midx_row}"
        f":{get_column_letter(month_col_end)}{grid_midx_row}"
    )
    for i, wp in enumerate(wps):
        _write_wp_summary_row(
            ws, wp_summary_first_row + i, wp,
            cfg_wp_row=cfg_wp_first_row + i,
            grid_range=grid_range,
            rate_col_range=rate_col_range,
            midx_range=midx_range,
            has_persons=num_persons > 0,
            year_specs=year_specs,
            grid_person_first_row=grid_person_first_row,
            grid_person_last_row=grid_person_last_row,
            stop_reason=stop_reasons.get(wp.code, "Underfilled"),
        )
    if num_wps > 0:
        assert yearly_total_row is not None and yearly_budget_row is not None
        _write_yearly_rows(
            ws,
            yearly_total_row=yearly_total_row,
            yearly_budget_row=yearly_budget_row,
            wp_summary_first_row=wp_summary_first_row,
            wp_summary_last_row=wp_summary_last_row,
            wps=wps,
            year_specs=year_specs,
            cfg_wp_first_row=cfg_wp_first_row,
        )
        for r in range(wp_summary_header_row, wp_summary_band_end_row + 1):
            ws.row_dimensions[r].outline_level = 1

    # ── TIMELOG GRID band ──
    _write_section_title(
        ws, grid_title_row, SECTION_TITLE_GRID, last_col=canvas_last_col,
    )
    _write_band_axis_headers(
        ws, grid_year_row, grid_month_row, project_params, month_col_start,
    )
    _write_month_index_row(
        ws, grid_midx_row, project_months, month_col_start,
    )
    if num_persons > 0:
        assert grid_total_row is not None
        for r in (
            grid_year_row, grid_month_row,
            grid_midx_row,
            *range(grid_person_first_row, grid_person_last_row + 1),
            grid_total_row,
        ):
            ws.row_dimensions[r].outline_level = 1
    else:
        for r in (grid_year_row, grid_month_row, grid_midx_row):
            ws.row_dimensions[r].outline_level = 1

    # ── PER-PERSON SUMMARY band ──
    _write_section_title(
        ws, summary_title_row, SECTION_TITLE_PERSON_SUMMARY,
        last_col=_PERSON_SUMMARY_RATE_PLUS_I_COL,
    )
    style_cell(
        ws, summary_header_row, 1, value="Person",
        font=HEADER_FONT, alignment=CENTER,
    )
    for i, hdr in enumerate(_PERSON_SUMMARY_METRIC_HEADERS):
        if hdr.endswith("+i"):
            hdr_fill = WITH_INDIRECT_HEADER_FILL
        elif hdr.endswith("-i"):
            hdr_fill = WITHOUT_INDIRECT_HEADER_FILL
        else:
            hdr_fill = None
        style_cell(
            ws, summary_header_row, _PERSON_SUMMARY_FIRST_METRIC_COL + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER, fill=hdr_fill,
        )
    if num_persons > 0:
        for r in range(summary_header_row, summary_person_last_row + 1):
            ws.row_dimensions[r].outline_level = 1

    # ── Fill person rows (grid + summary) ──
    cfg_person_rows: dict[str, int] = config_layout.get("person_rows", {})
    months_row_cfg = config_layout["param_rows"]["Project Months"]
    project_months_ref = cell_ref(SHEET_CONFIG, months_row_cfg, 2)

    for i, person in enumerate(persons):
        name = person.name
        grid_row = person_rows_by_band["grid"][name]
        summary_row = person_rows_by_band["summary"][name]
        cfg_row = cfg_person_rows.get(name)
        person_assignment = assigned_map.get(name, {})
        windows_for_person = person_windows.get(name, [])

        # Merged grid: Person xref + month cells (literals only).
        _write_person_col_a(ws, grid_row, name, cfg_row)
        for mi in range(project_months):
            month_num = mi + 1
            col = month_col_start + mi
            is_vacation = month_num in vacation_months_project
            # Defensive: if a preserved-grid cell lands inside a vacation
            # month (pre-enhancement workbook migrating forward), drop the
            # WP code — vacation cells are always empty + locked.
            wp_code = (
                None if is_vacation else person_assignment.get(month_num)
            )
            in_any_window = any(
                start <= month_num <= end for (_wp, start, end)
                in windows_for_person
            )
            _write_grid_cell(
                ws, grid_row, col, wp_code, in_any_window,
                is_vacation=is_vacation,
            )

        # PER-PERSON SUMMARY: Person xref + 8 metric cols.
        _write_person_col_a(ws, summary_row, name, cfg_row)
        _write_person_summary_metrics(
            ws, summary_row,
            grid_row=grid_row,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            month_header_row=grid_month_row,
            month_idx_row=grid_midx_row,
            project_months_ref=project_months_ref,
            cfg_row=cfg_row,
        )

    # ── Monthly total -i row (under TIMELOG GRID, same outline group) ──
    if grid_total_row is not None:
        _write_monthly_total_row(
            ws, row=grid_total_row,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            grid_person_first_row=grid_person_first_row,
            grid_person_last_row=grid_person_last_row,
            summary_person_first_row=summary_person_first_row,
            summary_person_last_row=summary_person_last_row,
            vacation_months_project=vacation_months_project,
            label="Monthly total -i",
            with_indirect=False,
        )

    # ── Monthly total +i row (outside outline group, always visible) ──
    if grid_person_cost_row is not None:
        _write_monthly_total_row(
            ws, row=grid_person_cost_row,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            grid_person_first_row=grid_person_first_row,
            grid_person_last_row=grid_person_last_row,
            summary_person_first_row=summary_person_first_row,
            summary_person_last_row=summary_person_last_row,
            vacation_months_project=vacation_months_project,
            label="Monthly total +i",
        )

    # ── Patch traffic-light formulas (downstream rows now exist) ──
    if num_wps > 0:
        # 1) Any undershoot? — Buffer € < 0 at col D of PER-WP SUMMARY.
        style_cell(
            ws, tl_row, tl_value_cols[0], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f'=COUNTIF(D{wp_summary_first_row}:D{wp_summary_last_row},"<0")'
        )
        # 2) Any out-of-window cell? — sum col I of PER-WP SUMMARY (shifted
        # from col H when Buffer % was inserted at col E, Aiko-port T3).
        style_cell(
            ws, tl_row, tl_value_cols[1], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = f"=SUM(I{wp_summary_first_row}:I{wp_summary_last_row})"
    else:
        for c in tl_value_cols[:2]:
            ws.cell(row=tl_row, column=c, value=0)

    if num_persons > 0:
        # 3) Any over-pool person? — Over-pool? at col D of PER-PERSON SUMMARY.
        op_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 2)  # D
        style_cell(
            ws, tl_row, tl_value_cols[2], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f"=COUNTIF({op_letter}{summary_person_first_row}"
            f":{op_letter}{summary_person_last_row},TRUE)"
        )
        # 4) Total switches — Switches at col E of PER-PERSON SUMMARY.
        sw_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 3)  # E
        style_cell(
            ws, tl_row, tl_value_cols[3], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f"=SUM({sw_letter}{summary_person_first_row}"
            f":{sw_letter}{summary_person_last_row})"
        )
    else:
        for c in tl_value_cols[2:]:
            ws.cell(row=tl_row, column=c, value=0)

    # CF on first three traffic-light value cells: red when > 0. Toggle cell
    # (col 10) is a TRUE/FALSE control — intentionally not flagged.
    for c in tl_value_cols[:3]:
        cell_addr = f"{get_column_letter(c)}{tl_row}"
        ws.conditional_formatting.add(
            cell_addr,
            FormulaRule(formula=[f"{cell_addr}>0"], fill=CF_RED_FILL),
        )

    # ── CF rules on canvas ──
    # Per ``openpyxl-patterns.md``: first rule added gets highest priority.
    if num_persons > 0 and num_wps > 0:
        # Rule 1: per-WP cell-level red for WP code outside its active window.
        # Targets the merged grid's month-idx helper row.
        top_left = grid_range.split(":")[0]
        tl_col_letter = "".join(ch for ch in top_left if ch.isalpha())
        tl_row_num = "".join(ch for ch in top_left if ch.isdigit())
        idx_ref = f"{tl_col_letter}${grid_midx_row}"
        cell_ref_inline = f"{tl_col_letter}{tl_row_num}"
        for wp in wps:
            formula = (
                f'AND({cell_ref_inline}="{wp.code}",'
                f"OR({idx_ref}<{wp.active_start},{idx_ref}>{wp.active_end}))"
            )
            ws.conditional_formatting.add(
                grid_range,
                FormulaRule(formula=[formula], fill=CF_RED_FILL),
            )

        # Rule 2: pastel fill per WP across the merged grid.
        for i, wp in enumerate(wps):
            pastel = CF_PASTEL_FILLS[i % len(CF_PASTEL_FILLS)]
            ws.conditional_formatting.add(
                grid_range,
                FormulaRule(
                    formula=[f'{cell_ref_inline}="{wp.code}"'],
                    fill=pastel,
                ),
            )

    # Rule 3 (PER-PERSON SUMMARY CF): Over-pool? = TRUE → red. Col D.
    if num_persons > 0:
        op_col_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 2)  # D
        op_range = (
            f"{op_col_letter}{summary_person_first_row}"
            f":{op_col_letter}{summary_person_last_row}"
        )
        ws.conditional_formatting.add(
            op_range,
            FormulaRule(
                formula=[f"{op_col_letter}{summary_person_first_row}=TRUE"],
                fill=CF_RED_FILL,
            ),
        )

    # Rule 4 (PER-WP SUMMARY CF): Buffer € (col D) — green>0, red<0.
    if num_wps > 0:
        buffer_range = f"D{wp_summary_first_row}:D{wp_summary_last_row}"
        # Red first (higher-priority conflict wins), then green.
        ws.conditional_formatting.add(
            buffer_range,
            FormulaRule(
                formula=[f"D{wp_summary_first_row}<0"], fill=CF_RED_FILL,
            ),
        )
        ws.conditional_formatting.add(
            buffer_range,
            FormulaRule(
                formula=[f"D{wp_summary_first_row}>0"], fill=CF_GREEN_FILL,
            ),
        )

        # Rule 5 (PER-WP SUMMARY CF): Burn rate (col H) exceeds max team rate.
        # Burn rate is in +i units (= -D{row}/wallclock, where D = C − B with
        # both C and B in +i), so the threshold must read from the +i rate
        # col on PER-PERSON SUMMARY (col I) — not the -i alias.
        burn_range = f"H{wp_summary_first_row}:H{wp_summary_last_row}"
        if num_persons > 0:
            rate_col_letter = get_column_letter(_PERSON_SUMMARY_RATE_PLUS_I_COL)
            max_rate_expr = (
                f"MAX({rate_col_letter}{summary_person_first_row}"
                f":{rate_col_letter}{summary_person_last_row})"
            )
            ws.conditional_formatting.add(
                burn_range,
                FormulaRule(
                    formula=[
                        f"AND(ISNUMBER(H{wp_summary_first_row}),"
                        f"H{wp_summary_first_row}>{max_rate_expr})"
                    ],
                    fill=CF_RED_FILL,
                ),
            )

        # Rule 6 (PER-WP SUMMARY CF): Any out-of-window? > 0 → red. Col I
        # (shifted from col H by Aiko-port T3).
        oow_range = f"I{wp_summary_first_row}:I{wp_summary_last_row}"
        ws.conditional_formatting.add(
            oow_range,
            FormulaRule(
                formula=[f"I{wp_summary_first_row}>0"],
                fill=CF_RED_FILL,
            ),
        )

    # ── Freeze panes + column widths + outline ──
    # Freeze rows 1–2 (banner + traffic lights) and col A (Person). Month grid
    # and summary metrics scroll — col A stays visible.
    ws.freeze_panes = ws.cell(
        row=3, column=TIMELOG_MONTH_COL_START,
    ).coordinate  # "B3"

    ws.column_dimensions["A"].width = 25          # Person
    # PER-PERSON SUMMARY metric cols: B..H. Slightly wider than default.
    for i in range(len(_PERSON_SUMMARY_METRIC_HEADERS)):
        c_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + i)
        ws.column_dimensions[c_letter].width = 12

    return {
        "sheet": SHEET_TIMELOG_PLAN,
        "month_col_start": month_col_start,
        "month_col_end": month_col_end,
        "grid_person_first_row": grid_person_first_row,
        "grid_person_last_row": grid_person_last_row,
        "summary_person_first_row": summary_person_first_row,
        "summary_person_last_row": summary_person_last_row,
        "person_rows_by_band": person_rows_by_band,
        "wp_summary_first_row": wp_summary_first_row,
        "wp_summary_last_row": wp_summary_last_row,
        "rebuild_from_month_cell": (
            f"B{config_layout['param_rows']['Rebuild from month #']}"
        ),
    }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_wps(
    work_packages: list[dict],
    project_params: dict,
) -> list[WP]:
    """Convert config WPs into solver WP records.

    Single-range schema: each WP carries its own ``start``/``end``/``type``
    directly. ``ceiling_eur`` (solver-internal field, pre-T1 naming) mirrors
    the Budget -i position at ``Config!E{wp_row}`` = ``FTE × Rate``. Per-WP
    rate is taken from ``wp["rate"]`` when present (post-reader rebuild
    path); on init the seed WP dicts may lack the key, so we fall back to
    ``project_params["rate"]`` — the same default used by
    ``config_sheet.py`` to populate the sheet.
    """
    default_rate = project_params.get("rate", 10000)
    if not isinstance(default_rate, (int, float)):
        default_rate = 10000
    wps: list[WP] = []
    for wp in work_packages:
        if "start" not in wp or "end" not in wp:
            continue
        active_start = int(wp["start"])
        active_end = int(wp["end"])
        wp_rate = wp.get("rate", default_rate)
        if not isinstance(wp_rate, (int, float)):
            wp_rate = default_rate
        ceiling_eur = int(round(float(wp["fte"]) * float(wp_rate)))
        wps.append(WP(
            code=wp["id"],
            active_start=active_start,
            active_end=active_end,
            ceiling_eur=ceiling_eur,
        ))
    return wps


def _upsert_defined_name(wb: Workbook, name: str, attr_text: str) -> None:
    """Idempotently add/replace a workbook-scoped defined name."""
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names.add(DefinedName(name, attr_text=attr_text))


def _write_section_title(
    ws: Worksheet, row: int, title: str, last_col: int,
) -> None:
    """Merge cols A..last_col on ``row`` and write ``title`` styled as a section bar."""
    ws.merge_cells(
        start_row=row, start_column=1, end_row=row, end_column=last_col,
    )
    style_cell(
        ws, row, 1, value=title, font=SECTION_FONT, fill=SECTION_FILL,
        alignment=CENTER, border=False,
    )
    for c in range(2, last_col + 1):
        style_cell(ws, row, c, fill=SECTION_FILL, border=False)


def _write_band_axis_headers(
    ws: Worksheet,
    year_row: int,
    month_row: int,
    project_params: dict,
    month_col_start: int,
) -> None:
    """Write the year + month headers for the merged grid.

    Col A is filled with LIGHT_BLUE_FILL on the year row and styled with a
    "Person" header on the month row — the grid starts at col B.
    """
    project_months = project_params["months"]
    project_start = project_params["start"]

    # Year row: LIGHT_BLUE_FILL on col A, then calendar-year blocks.
    style_cell(ws, year_row, 1, value="", fill=LIGHT_BLUE_FILL)
    for group_start, span in _calendar_year_groups(project_start, project_months):
        year_start_col = month_col_start + group_start
        year_end_col = year_start_col + span - 1
        cal_date = project_start + relativedelta(months=group_start)
        year_label = str(cal_date.year)
        if span > 1:
            ws.merge_cells(
                start_row=year_row, start_column=year_start_col,
                end_row=year_row, end_column=year_end_col,
            )
        style_cell(
            ws, year_row, year_start_col, value=year_label, font=HEADER_FONT,
            fill=LIGHT_BLUE_FILL, alignment=CENTER,
        )
        for c2 in range(year_start_col + 1, year_end_col + 1):
            style_cell(ws, year_row, c2, fill=LIGHT_BLUE_FILL)

    # Month row: "Person" header in A, then per-month labels.
    style_cell(
        ws, month_row, 1, value="Person", font=HEADER_FONT,
        fill=SECTION_FILL, alignment=CENTER,
    ).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")

    for mi in range(project_months):
        c = month_col_start + mi
        cal_date = project_start + relativedelta(months=mi)
        mlabel = f"M{mi + 1}-{cal_date.strftime('%b%y')}"
        style_cell(
            ws, month_row, c, value=mlabel,
            font=Font(name="Arial", size=8),
            fill=LIGHT_BLUE_FILL, alignment=CENTER,
        )


def _calendar_year_groups(project_start, project_months: int):
    """Yield (group_start_month_index, span) tuples — mirrors layout.py's helper."""
    m = 0
    while m < project_months:
        cal_date = project_start + relativedelta(months=m)
        months_left_in_year = 12 - cal_date.month + 1
        span = min(months_left_in_year, project_months - m)
        yield m, span
        m += span


@dataclass(frozen=True)
class _YearColSpec:
    """One calendar-year column on PER-WP SUMMARY.

    ``year`` is the calendar year (e.g. 2026). ``first_month`` / ``last_month``
    are 1-indexed month positions within the project timeline — used both to
    build the grid subrange reference and to check overlap against each WP's
    active window at build time.
    """
    year: int
    first_month: int
    last_month: int
    col_start: int
    col_end: int


def _year_col_specs(
    project_start, project_months: int, month_col_start: int,
) -> list[_YearColSpec]:
    """Compute the dynamic Y{yyyy} +i column specs for PER-WP SUMMARY."""
    specs: list[_YearColSpec] = []
    for group_start, span in _calendar_year_groups(project_start, project_months):
        cal_date = project_start + relativedelta(months=group_start)
        specs.append(
            _YearColSpec(
                year=cal_date.year,
                first_month=group_start + 1,
                last_month=group_start + span,
                col_start=month_col_start + group_start,
                col_end=month_col_start + group_start + span - 1,
            )
        )
    return specs


def _write_month_index_row(
    ws: Worksheet,
    row: int,
    project_months: int,
    month_col_start: int,
) -> None:
    """Write a hidden 1..N month-index helper row for SUMPRODUCT OOW checks."""
    style_cell(
        ws, row, 1, value="(month idx)",
        font=Font(name="Arial", italic=True, color="808080", size=8),
        border=False,
    )
    for mi in range(project_months):
        c = month_col_start + mi
        style_cell(
            ws, row, c, value=mi + 1,
            font=Font(name="Arial", color="808080", size=8),
            alignment=CENTER, border=False,
        )
    ws.row_dimensions[row].hidden = True


def _write_person_col_a(
    ws: Worksheet, row: int, name: str, cfg_row: int | None,
) -> None:
    """Write col A on a person row: cross-ref into Config Team Roster."""
    if cfg_row is not None:
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{cfg_row}"
    else:
        style_cell(ws, row, 1, value=name, font=GREEN_XREF)


def _grid_range_ref(
    col_start: int, col_end: int,
    row_start: int, row_end: int,
) -> str:
    """Return ``"B7:AV20"`` for a rectangular range."""
    return (
        f"{get_column_letter(col_start)}{row_start}"
        f":{get_column_letter(col_end)}{row_end}"
    )


def _vacation_months_in_span(
    vac_month_num: int, project_start, project_months: int,
) -> frozenset[int]:
    """Return project-1-indexed month numbers whose calendar month == vac_month_num.

    E.g. project_start = Nov 2026, project_months = 14, vac_month_num = 7 →
    {9} (every July inside the span).
    """
    out: set[int] = set()
    for i in range(project_months):
        d = project_start + relativedelta(months=i)
        if d.month == vac_month_num:
            out.add(i + 1)
    return frozenset(out)


def _write_grid_cell(
    ws: Worksheet,
    row: int,
    col: int,
    wp_code: str | None,
    in_any_window: bool,
    is_vacation: bool = False,
) -> None:
    """Write a single merged-grid cell.

    Cells hold LITERAL values (v2: the v1 `=<Proposed>` default is gone).
    Empty cells outside any person-WP window carry DARK_GREY_FILL; in-window-
    but-empty cells carry GREY_FILL. All cells are unlocked and BLUE_INPUT —
    this band IS the edit surface.

    Vacation-month cells are an exception: always empty string, VACATION_FILL
    (pale yellow), Protection(locked=True), default font (not BLUE_INPUT).
    Solver never places WP codes there; on rebuild any preserved WP code in
    a vacation cell (frozen or not) is discarded before solve, per the
    freeze-point design (vacation cells are always empty + locked).

    Rule 3 carve-out: cells hold literal strings (WP codes), not formulas,
    but styled BLUE_INPUT because the user types here.
    """
    if is_vacation:
        cell = style_cell(
            ws, row, col, value="",
            fill=VACATION_FILL, alignment=CENTER,
        )
        cell.protection = Protection(locked=True)
        return
    if wp_code is not None:
        cell = style_cell(
            ws, row, col, value=wp_code,
            font=BLUE_INPUT, alignment=CENTER,
        )
    elif in_any_window:
        cell = style_cell(
            ws, row, col, value="",
            font=BLUE_INPUT, fill=GREY_FILL, alignment=CENTER,
        )
    else:
        cell = style_cell(
            ws, row, col, value="",
            font=BLUE_INPUT, fill=DARK_GREY_FILL, alignment=CENTER,
        )
    cell.protection = Protection(locked=False)


def _write_yearly_rows(
    ws: Worksheet,
    *,
    yearly_total_row: int,
    yearly_budget_row: int,
    wp_summary_first_row: int,
    wp_summary_last_row: int,
    wps: list[WP],
    year_specs: list["_YearColSpec"],
    cfg_wp_first_row: int,
) -> None:
    """Write the Yearly Assigned +i and Yearly Budget +i rows under PER-WP SUMMARY.

    - Col A: blank. (Pre-D9 layout placed the row label here; shifted to col I
      via TIMELOG_YEARLY_LABEL_COL so the label sits adjacent-left to the
      per-year breakdown. Aiko-port T3 shifted col H → I when Buffer % was
      inserted at col E.)
    - Col I (``TIMELOG_YEARLY_LABEL_COL``): row label (HEADER_FONT) with
      ``WITH_INDIRECT_HEADER_FILL`` — the per-year cells on this row are +i,
      so the label carries the +i header fill.
    - Col C: cross-check sum (across-WP SUM of the PER-WP SUMMARY column
      that feeds the row's aggregate). Aiko-port Q-B dropped the inline
      ``*1.25`` from both rows:
        * ``yearly_total_row`` col C = ``=SUM(C{first}:C{last})``. PER-WP
          SUMMARY col C (Assigned €) is -i; this cross-check is -i by
          construction. The per-year cells to its right stay +i via the
          ×1.25 wrap at the leaf SUMPRODUCT (T7.3 display-layer wrap).
        * ``yearly_budget_row`` col C = ``=SUM(B{first}:B{last})``. Q-B
          flipped PER-WP SUMMARY col B to direct Config!F (Budget +i,
          ``COL_WP_BUDGET_PLUS_I``), so SUM(B) is already +i.
    - Cols B, D-H: intentionally blank.
    - Per-year cols (K…):
      * yearly_total_row: =SUM(<year_col>{W1}:<year_col>{Wn}) — stacks the
        per-WP Y{yyyy} +i cells for that year (each cell is +i via the T7.3
        ×1.25 wrap at the display layer).
      * yearly_budget_row: build-time term-sum
        ``Σ prop_ky × Config!F{wp_row}`` — Config col F = Budget +i
        (``COL_WP_BUDGET_PLUS_I``). Budget +i is the authoritative +i source
        per D1/D4; no ×1.25 applied.
        ``prop_ky = months_of_wpk_in_year_y / total_sched_months_for_wpk``
        for WPs overlapping year y; zero-term WPs omitted (design §5 D16).
        Formula stays live to Config edits; rebuild re-derives the proportions
        when the WP schedule changes.
    """
    style_cell(
        ws, yearly_total_row, TIMELOG_YEARLY_LABEL_COL, value="Yearly Assigned +i",
        font=HEADER_FONT, fill=WITH_INDIRECT_HEADER_FILL, alignment=CENTER,
    )
    style_cell(
        ws, yearly_budget_row, TIMELOG_YEARLY_LABEL_COL, value="Yearly Budget +i",
        font=HEADER_FONT, fill=WITH_INDIRECT_HEADER_FILL, alignment=CENTER,
    )

    # Col C cross-checks (Aiko-port Q-B: raw SUMs, no inline ×1.25):
    #   - yearly_total_row col C = SUM(C6:C13). Matches Aiko's raw total;
    #     PER-WP SUMMARY col C (Assigned €) is -i, so this cross-check is
    #     -i by construction. +i parity with the per-year cells on this
    #     row is no longer enforced at col C — the per-year SUMs stay +i
    #     via the ×1.25 wrap at the leaf (T7.3).
    #   - yearly_budget_row col C = SUM(B6:B13). Aiko-port Q-B flipped
    #     PER-WP SUMMARY col B to direct Config!F (Budget +i via
    #     COL_WP_BUDGET_PLUS_I), so SUM(B) is already +i — no ×1.25.
    yt_c = style_cell(
        ws, yearly_total_row, 3, font=BLACK_FORMULA, fmt="€#,##0",
    )
    yt_c.value = f"=SUM(C{wp_summary_first_row}:C{wp_summary_last_row})"
    yt_c.comment = sized_comment(
        "Cross-check SUM of PER-WP SUMMARY col C (Assigned €, -i). "
        "Row label 'Yearly Assigned +i' refers to the per-year cells to "
        "the right (all +i via ×1.25 at the SUMPRODUCT leaf). Compare "
        "this cell × 1.25 against the per-year sum for drift detection.",
        author="Programme Budget Tool",
    )
    style_cell(
        ws, yearly_budget_row, 3, font=BLACK_FORMULA, fmt="€#,##0",
    ).value = f"=SUM(B{wp_summary_first_row}:B{wp_summary_last_row})"

    # Precompute per-WP schedule-month sets for build-time overlap math.
    # Single-range schema: each WP is one contiguous interval.
    wp_month_sets: list[set[int]] = []
    wp_sched_totals: list[int] = []
    for wp in wps:
        ms = set(range(wp.active_start, wp.active_end + 1))
        wp_month_sets.append(ms)
        wp_sched_totals.append(len(ms))

    for yi, spec in enumerate(year_specs):
        col = _WP_SUMMARY_BASE_LAST_COL + 1 + yi
        col_letter_str = get_column_letter(col)

        # Yearly Assigned +i: stack the per-WP year-col cells.
        style_cell(
            ws, yearly_total_row, col, font=BLACK_FORMULA, fmt="€#,##0",
        ).value = (
            f"=SUM({col_letter_str}{wp_summary_first_row}"
            f":{col_letter_str}{wp_summary_last_row})"
        )

        # Yearly budget €: build-time term-sum with literal proportions.
        terms: list[str] = []
        for wi, wp in enumerate(wps):
            total_months = wp_sched_totals[wi]
            if total_months <= 0:
                continue
            months_in_year = sum(
                1 for m in wp_month_sets[wi]
                if spec.first_month <= m <= spec.last_month
            )
            if months_in_year <= 0:
                continue
            prop = months_in_year / total_months
            budget_ref = cell_ref(
                SHEET_CONFIG, cfg_wp_first_row + wi, COL_WP_BUDGET_PLUS_I,
            )
            terms.append(f"{prop}*{budget_ref}")
        if terms:
            style_cell(
                ws, yearly_budget_row, col, font=BLACK_FORMULA, fmt="€#,##0",
            ).value = "=" + "+".join(terms)
        else:
            style_cell(
                ws, yearly_budget_row, col, value="",
                font=BLACK_FORMULA, fmt="€#,##0",
            )


def _write_monthly_total_row(
    ws: Worksheet,
    *,
    row: int,
    month_col_start: int,
    month_col_end: int,
    grid_person_first_row: int,
    grid_person_last_row: int,
    summary_person_first_row: int,
    summary_person_last_row: int,
    vacation_months_project: frozenset[int],
    label: str = "Monthly total +i",
    with_indirect: bool = True,
) -> None:
    """Write a per-month aggregate row.

    Each month cell sums non-empty grid cells × the corresponding person's
    rate. Rates are -i; ``with_indirect=True`` wraps the result in ×1.25 to
    flip the display to +i. Vacation-month cells additionally carry
    VACATION_FILL for visual column continuity (formula still evaluates to 0).

    Used for both the "Monthly total -i" row (under the grid, inside the
    outline group) and the "Monthly total +i" row (just above PER-PERSON
    SUMMARY, always visible). Callers select via ``label`` + ``with_indirect``.
    """
    header_fill = (
        WITH_INDIRECT_HEADER_FILL if with_indirect
        else WITHOUT_INDIRECT_HEADER_FILL
    )
    style_cell(
        ws, row, 1, value=label,
        font=HEADER_FONT, fill=header_fill, alignment=CENTER,
    )
    rate_col_letter = get_column_letter(_PERSON_SUMMARY_RATE_COL)
    rate_range = (
        f"{rate_col_letter}{summary_person_first_row}"
        f":{rate_col_letter}{summary_person_last_row}"
    )
    indirect_wrap = "*1.25" if with_indirect else ""
    for mi in range(month_col_end - month_col_start + 1):
        month_num = mi + 1
        col = month_col_start + mi
        col_letter_str = get_column_letter(col)
        grid_col_range = (
            f"{col_letter_str}{grid_person_first_row}"
            f":{col_letter_str}{grid_person_last_row}"
        )
        is_vacation = month_num in vacation_months_project
        cell = style_cell(
            ws, row, col,
            font=BLACK_FORMULA, alignment=CENTER, fmt="€#,##0",
            fill=VACATION_FILL if is_vacation else None,
        )
        cell.value = (
            f'=ROUND(SUMPRODUCT(({grid_col_range}<>"")*{rate_range}){indirect_wrap},0)'
        )


def _write_wp_summary_row(
    ws: Worksheet,
    row: int,
    wp: WP,
    *,
    cfg_wp_row: int,
    grid_range: str,
    rate_col_range: str,
    midx_range: str,
    has_persons: bool,
    year_specs: list["_YearColSpec"],
    grid_person_first_row: int,
    grid_person_last_row: int,
    stop_reason: str,
) -> None:
    """Fill the per-WP-summary cells on a single WP row.

    Cols A..J are the 10 base metrics (col J = solver "Stop reason"); cols
    K.. hold one ``Y{yyyy} +i`` column per calendar year spanned by the
    project. Year cells where the WP's active window does not overlap the
    year are blank (build-time check).
    """
    wp_code = wp.code

    # Col A: WP code (literal).
    style_cell(ws, row, 1, value=wp_code, font=BLACK_FORMULA, alignment=CENTER)

    # Col B: Budget +i — green xref to Config!F{wp_row} (COL_WP_BUDGET_PLUS_I).
    # Aiko-port Q-B: B-col on PER-WP SUMMARY is Budget +i direct from Config,
    # so the TOTAL-row col C cross-check on the yearly-budget row becomes a
    # plain SUM(B6:B13) without ×1.25 (pairs with _write_yearly_rows change).
    budget_ref = cell_ref(SHEET_CONFIG, cfg_wp_row, COL_WP_BUDGET_PLUS_I)
    style_cell(ws, row, 2, font=GREEN_XREF, fmt="€#,##0").value = f"={budget_ref}"

    # Col C: Assigned € = ROUND(SUMPRODUCT((grid="WP") * rate_vector), 0).
    # Rate vector is the PER-PERSON SUMMARY's col I (Rate +i) over the
    # person-row range; grid rows match one-to-one with summary rows. Result
    # is +i, so col B (Budget +i) − col C lands in matching units.
    if has_persons:
        style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="€#,##0").value = (
            f'=ROUND(SUMPRODUCT(({grid_range}="{wp_code}")'
            f"*{rate_col_range}),0)"
        )
    else:
        style_cell(ws, row, 3, value=0, font=BLACK_FORMULA, fmt="€#,##0")

    # Col D: Buffer € = Assigned − Budget. Positive = buffer (good);
    # negative = shortfall (bad). CF green>0 / red<0.
    style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=C{row}-B{row}"
    )

    # Col E: Buffer % = Buffer / Budget. Aiko-port T3 (2026-04-23) inserted
    # this col; everything right of col D shifted one right.
    style_cell(ws, row, 5, font=BLACK_FORMULA, fmt="0.0%").value = (
        f"=D{row}/B{row}"
    )

    # Col F: Assigned months = SUMPRODUCT(1*(grid="WP")).
    if has_persons:
        style_cell(ws, row, 6, font=BLACK_FORMULA, fmt="0").value = (
            f'=SUMPRODUCT(1*({grid_range}="{wp_code}"))'
        )
    else:
        style_cell(ws, row, 6, value=0, font=BLACK_FORMULA, fmt="0")

    # Col G: WP sched months — build-time literal (end − start + 1 under the
    # single-range schema). Rule 3 carve-out: BLACK_FORMULA font on a literal
    # — signals "builder-written, do not edit." Cell comment echoes the
    # deviation.
    sched_cell = style_cell(
        ws, row, 7, value=int(wp.active_end - wp.active_start + 1),
        font=BLACK_FORMULA, fmt="0", alignment=CENTER,
    )
    sched_cell.comment = sized_comment(
        "Number of months this WP is active (end − start + 1). Literal "
        "(not a formula) — recomputed on rebuild from Config.",
        author="Programme Budget Tool",
    )

    # Col H: Burn rate needed — IF-guarded; "" when buffer≥0 or wallclock=0.
    # wallclock = MAX(0, wp.active_end − current_month); buffer = D{row}.
    wallclock_expr = f"MAX(0,{wp.active_end}-current_month)"
    style_cell(ws, row, 8, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=IF(AND(D{row}<0,{wallclock_expr}>0),"
        f'-D{row}/{wallclock_expr},"")'
    )

    # Col I: Any out-of-window? = SUMPRODUCT((grid="WP") * OOW).
    if has_persons:
        style_cell(ws, row, 9, font=BLACK_FORMULA, fmt="0").value = (
            f'=SUMPRODUCT(({grid_range}="{wp_code}")*'
            f"(({midx_range}<{wp.active_start})+"
            f"({midx_range}>{wp.active_end})))"
        )
    else:
        style_cell(ws, row, 9, value=0, font=BLACK_FORMULA, fmt="0")

    # Col J: Stop reason — 1–2-word literal from the solver post-pass
    # classifier. Rule 3 carve-out: BLACK_FORMULA on a literal — signals
    # "builder-written, do not edit." Recomputed on rebuild.
    style_cell(
        ws, row, 10, value=stop_reason,
        font=BLACK_FORMULA, alignment=CENTER,
    )

    # Cols K..: Y{yyyy} +i — EUR placed by the solver in this WP × year slice.
    # (Col K on PER-WP SUMMARY — this band has 10 base cols A..J; year
    # cols start at col K.) ``rate_col_range`` already targets the +i col on
    # PER-PERSON SUMMARY, so the SUMPRODUCT result is +i without an extra
    # ×1.25 wrap. Rounded per the programme leaf-rounding rule.
    for i, spec in enumerate(year_specs):
        col = _WP_SUMMARY_BASE_LAST_COL + 1 + i
        overlaps = (
            spec.first_month <= wp.active_end
            and spec.last_month >= wp.active_start
        )
        if has_persons and overlaps:
            year_subrange = _grid_range_ref(
                spec.col_start, spec.col_end,
                grid_person_first_row, grid_person_last_row,
            )
            style_cell(ws, row, col, font=BLACK_FORMULA, fmt="€#,##0").value = (
                f'=ROUND(SUMPRODUCT(({year_subrange}="{wp_code}")'
                f"*{rate_col_range}),0)"
            )
        else:
            style_cell(ws, row, col, value="", font=BLACK_FORMULA, fmt="€#,##0")


def _write_person_summary_metrics(
    ws: Worksheet,
    row: int,
    *,
    grid_row: int,
    month_col_start: int,
    month_col_end: int,
    month_header_row: int,
    month_idx_row: int,
    project_months_ref: str,
    cfg_row: int | None,
) -> None:
    """Fill the 8 per-person-summary metric cells on cols B..I.

    All metrics target the merged grid's row for the same person — that's
    the canonical edit surface. Cols B..G are non-rate metrics; col H is
    Rate (€/PM) -i (VLOOKUP); col I is Rate (€/PM) +i (= H × 1.25).
    """
    c0 = _PERSON_SUMMARY_FIRST_METRIC_COL

    grid_row_range = month_range_ref(grid_row, month_col_start, month_col_end)
    grid_left_range = month_range_ref(grid_row, month_col_start, month_col_end - 1)
    grid_right_range = month_range_ref(grid_row, month_col_start + 1, month_col_end)
    month_headers_range = month_range_ref(
        month_header_row, month_col_start, month_col_end,
    )
    month_idx_range = month_range_ref(
        month_idx_row, month_col_start, month_col_end,
    )

    # Col B: Assigned = COUNTIF(row, "WP*").
    style_cell(ws, row, c0, font=BLACK_FORMULA, fmt="0").value = (
        f'=COUNTIF({grid_row_range},"WP*")'
    )

    # Col C: Pool = ROUND(max_fte * (project_months − num_vacation_months), 0).
    # Green xref. num_vacation_months is a Config named range (builder-written
    # literal, Rule 3 carve-out — see docs/spec.md § Future
    # invariants → Named-range contract).
    if cfg_row is not None:
        max_fte_ref = cell_ref(SHEET_CONFIG, cfg_row, 2)
        style_cell(ws, row, c0 + 1, font=GREEN_XREF, fmt="0").value = (
            f"=ROUND({max_fte_ref}*({project_months_ref}-num_vacation_months),0)"
        )
    else:
        style_cell(ws, row, c0 + 1, value=0, font=BLACK_FORMULA, fmt="0")

    # Col D: Over-pool? = Assigned > Pool. Same-row ref.
    style_cell(
        ws, row, c0 + 2, font=BLACK_FORMULA, alignment=CENTER,
    ).value = f"={col_letter(c0)}{row}>{col_letter(c0 + 1)}{row}"

    # Col E: Switches = SUMPRODUCT(1*(left <> right)).
    style_cell(ws, row, c0 + 3, font=BLACK_FORMULA, fmt="0").value = (
        f"=SUMPRODUCT(1*({grid_left_range}<>{grid_right_range}))"
    )

    # Col F: First log (LO-stable SUMPRODUCT-MIN pattern).
    # MIN picks the smallest (month index of first non-empty cell); empty
    # cells contribute a 99999 sentinel so they're never selected. All-empty
    # row → MIN = 99999 → INDEX out of range → IFERROR returns "".
    style_cell(
        ws, row, c0 + 4, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f"=IFERROR(INDEX({month_headers_range},"
        f'SUMPRODUCT(MIN(({grid_row_range}<>"")*{month_idx_range}'
        f'+({grid_row_range}="")*99999))),"")'
    )

    # Col G: Last log (LO-stable SUMPRODUCT-MAX pattern).
    # MAX picks the largest (month index of last non-empty cell); empty cells
    # contribute 0. All-empty row → MAX = 0, but INDEX(range, 0) in LO does
    # NOT raise — it returns the full range and implicit-intersects at the
    # formula cell's column, leaking a neighbour's header into Last log.
    # Explicit IF(MAX=0, "", …) guard is required — asymmetric to First log's
    # 99999 sentinel because MAX can't naturally produce an out-of-range
    # sentinel (any value > N would dominate valid midx values).
    # The SUMPRODUCT(MAX(…)) subexpression is evaluated twice (IF test + INDEX
    # index). LET() would de-duplicate but requires Excel 365 / LO 7.4+; keep
    # the duplicated form until the project raises its recalc floor.
    last_log_max = (
        f'SUMPRODUCT(MAX(({grid_row_range}<>"")*{month_idx_range}))'
    )
    style_cell(
        ws, row, c0 + 5, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f'=IFERROR(IF({last_log_max}=0,"",'
        f"INDEX({month_headers_range},{last_log_max})),\"\")"
    )

    # Col H: Rate (€/PM) -i — single-source VLOOKUP into roster_lookup;
    # roster rates are -i.
    rate_minus_i_letter = col_letter(c0 + 6)
    style_cell(ws, row, c0 + 6, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=VLOOKUP(A{row},roster_lookup,4,FALSE)"
    )
    # Col I: Rate (€/PM) +i = -i × 1.25. Consumed by PER-WP SUMMARY
    # Assigned-€ SUMPRODUCT so the result lands in +i (matches Budget +i).
    style_cell(ws, row, c0 + 7, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"={rate_minus_i_letter}{row}*1.25"
    )
