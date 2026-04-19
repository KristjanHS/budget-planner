"""Timelog Plan sheet builder — prescriptive 1.0-FTE WP allocation grid.

A separate allocation surface from the Plan sheet. For each (person, month)
the sheet prescribes which WP they should log 100% against.

Layout (v2 + 2026-04-19 enhancements — supersedes v1's PROPOSED + ASSIGNED twin bands):

  1. Banner (row 1)
  2. Traffic lights + resolve-flag toggle (row 2)
  3. blank spacer
  4. PER-WP SUMMARY band (8 base cols + one EUR col per calendar year)
     blank spacer
  5. TIMELOG GRID — single editable band, literal WP codes + empty strings
     blank spacer
  6. PER-PERSON SUMMARY band (8 cols: Person + 6 metrics + Rate)

Each data band is wrapped in an Excel row-group (``outline_level=1``) so the
user can fold away sections not under review. Section-title rows stay at
level 0 and remain visible when the group collapses.

Col A is Person (cross-ref into Config Team Roster) across the grid and the
PER-PERSON SUMMARY band. No per-row metric column — the merged grid starts
at col B (TIMELOG_MONTH_COL_START). The per-person Rate helper lives only
once, on col H of the PER-PERSON SUMMARY band, and is consumed by the
PER-WP SUMMARY Assigned-€ SUMPRODUCT via row-aligned vectors.

Spec: ``docs/plans/2026-04-19-timelog-plan-enhancements-design.md`` (builds
on ``docs/plans/archived/2026-04-19-timelog-plan-v2-design.md``).

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
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
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
    col_letter,
    sized_comment,
    style_cell,
)
from budget_planner.timelog_solver import Person, PersonWPWindow, WP, solve

# Section-title strings — exported so the reader can anchor on the grid title
# via the same literal the builder writes. Update here, not at the callsite.
SECTION_TITLE_WP_SUMMARY = "PER-WP SUMMARY"
SECTION_TITLE_GRID = "TIMELOG GRID (edit freely — toggle re-solve at J2)"
SECTION_TITLE_PERSON_SUMMARY = "PER-PERSON SUMMARY"

# PER-PERSON SUMMARY layout: col A = Person, cols B..H = 7 metric cols.
# Metric headers start at col B. Rate is the last metric (col H) — single
# source of the per-person rate, consumed by PER-WP Assigned-€ SUMPRODUCT.
_PERSON_SUMMARY_METRIC_HEADERS = [
    "Assigned", "Pool", "Over-pool?",
    "Switches", "First log", "Last log",
    "Rate",
]
_PERSON_SUMMARY_FIRST_METRIC_COL = 2   # col B (Assigned)
_PERSON_SUMMARY_RATE_COL = (
    _PERSON_SUMMARY_FIRST_METRIC_COL + len(_PERSON_SUMMARY_METRIC_HEADERS) - 1
)  # col H

# PER-WP SUMMARY metric columns live at A..H (WP code + 7 metrics).
# Phase 3 appends dynamic "Y{yyyy} €" cols right of col H — one per calendar
# year spanned by the project. See `_year_col_specs`.
_WP_SUMMARY_BASE_HEADERS = [
    "WP code", "Ceiling €", "Assigned €", "Buffer €",
    "Assigned months", "WP sched months", "Burn rate needed",
    "Any out-of-window?",
]
_WP_SUMMARY_BASE_LAST_COL = len(_WP_SUMMARY_BASE_HEADERS)  # 8

# Toggle cell coordinates on row 2. Pair (label col, value col) = (9, 10).
_TOGGLE_LABEL_COL = 9
_TOGGLE_VALUE_COL = 10
_TOGGLE_NAMED_RANGE = "timelog_resolve_flag"


# ---------------------------------------------------------------------------
# Public builder
# ---------------------------------------------------------------------------


def build_timelog_plan_sheet(
    ws: Worksheet,
    wb: Workbook,
    *,
    project_params: dict,
    work_packages: list[dict],
    wp_schedule: list[dict],
    team: list[dict],
    assignments: list[dict],
    grade_rates: Mapping[str, float],
    config_layout: dict,
    preserved_grid: Mapping[tuple[str, int], str] | None = None,
    resolve_flag: bool = True,
) -> dict:
    """Build the Timelog Plan sheet with merged grid + summary bands.

    Args:
        ws: Worksheet to populate (created by orchestrator, named "Timelog Plan").
        wb: Parent workbook — receives the ``timelog_resolve_flag`` named range.
        project_params: dict with keys 'title', 'start', 'months', 'rate', etc.
        work_packages: list of WP dicts with keys 'id', 'name', 'fte', 'rate'.
        wp_schedule: list of WP schedule entries (id, start, end, type).
        team: list of team-member dicts with keys 'name', 'max_fte', 'grade',
            'priority' (1-3, defaults to 2 from the reader).
        assignments: list of person-WP assignment dicts (person, wp,
            total_fte, start, end). Treated as the solver's hard
            (person, wp) → window constraints.
        grade_rates: dict grade_name → EUR/PM. Used by the solver for
            €-accounting: each Person's ``rate_eur_per_pm`` is resolved
            from this mapping via the roster's grade.
        config_layout: dict returned by build_config_sheet() — must include
            ``param_rows``, ``wp_first_row``, ``person_rows``.
        preserved_grid: Optional mapping ``(person_name, month_num)`` →
            WP code. When ``resolve_flag`` is False, the grid is written
            from this dict verbatim (user edits preserved). When
            ``resolve_flag`` is True, the solver runs and this dict is
            ignored.
        resolve_flag: Gate for the solver. True → solver runs, grid is
            re-derived from assignments. False → grid is written from
            ``preserved_grid``, preserving user edits. On init the caller
            passes True; on rebuild the caller passes the user's value
            from the toggle cell.

    Returns:
        Layout dict with keys:
          sheet, month_col_start, month_col_end,
          grid_person_first_row, grid_person_last_row,
          summary_person_first_row, summary_person_last_row,
          wp_summary_first_row, wp_summary_last_row,
          person_rows_by_band (dict band → {name: row}),
          resolve_flag_cell (coordinate string).
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
    sorted_team = sorted(
        team, key=lambda m: (int(m.get("priority", 2)), m["name"])
    )
    persons = [
        Person(
            name=m["name"],
            priority=int(m.get("priority", 2)),
            pool=round(
                float(m["max_fte"])
                * (project_months - len(vacation_months_project))
            ),
            rate_eur_per_pm=int(round(grade_rates[m["grade"]])),
            vacation_months=vacation_months_project,
        )
        for m in sorted_team
    ]
    wps = _build_wps(work_packages, wp_schedule, project_params)
    windows = [
        PersonWPWindow(
            person_name=a["person"],
            wp_code=a["wp"],
            start_month=int(a["start"]),
            end_month=int(a["end"]),
        )
        for a in (assignments or [])
    ]
    if resolve_flag:
        assigned_map = solve(persons, wps, windows, project_months)
    else:
        assigned_map = {}
        for (pname, month_num), wp_code in (preserved_grid or {}).items():
            assigned_map.setdefault(pname, {})[month_num] = wp_code
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

    grid_title_row = wp_summary_last_row + 2  # +1 spacer

    grid_year_row = grid_title_row + 1
    grid_month_row = grid_year_row + 1
    grid_midx_row = grid_month_row + 1
    grid_person_first_row = grid_midx_row + 1
    grid_person_last_row = (
        grid_person_first_row + num_persons - 1 if num_persons > 0
        else grid_person_first_row - 1
    )

    summary_title_row = (
        (grid_person_last_row if num_persons > 0 else grid_midx_row) + 2
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
    # Toggle (cols 9-10): label + BLUE_INPUT TRUE/FALSE cell.
    style_cell(
        ws, tl_row, _TOGGLE_LABEL_COL,
        value="Re-run solver on next rebuild?",
        font=HEADER_FONT, alignment=CENTER,
    )
    toggle_cell = style_cell(
        ws, tl_row, _TOGGLE_VALUE_COL,
        value=bool(resolve_flag),
        font=BLUE_INPUT, alignment=CENTER,
    )
    toggle_cell.comment = sized_comment(
        "TRUE: next rebuild re-solves the grid (user edits discarded). "
        "FALSE: next rebuild preserves grid values verbatim. "
        "No auto-reset — flip manually.",
        author="Programme Budget Tool",
    )
    # Register named range (hyphen-/space-safe single-quoted sheet ref).
    toggle_ref = f"'{SHEET_TIMELOG_PLAN}'!${get_column_letter(_TOGGLE_VALUE_COL)}${tl_row}"
    _upsert_defined_name(wb, _TOGGLE_NAMED_RANGE, toggle_ref)

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
        f"Y{spec.year} €" for spec in year_specs
    ]
    for i, hdr in enumerate(wp_summary_headers):
        style_cell(
            ws, wp_summary_header_row, 1 + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER,
        )
    cfg_wp_first_row = config_layout["wp_first_row"]
    rate_col_range = (
        f"{get_column_letter(_PERSON_SUMMARY_RATE_COL)}{summary_person_first_row}"
        f":{get_column_letter(_PERSON_SUMMARY_RATE_COL)}{summary_person_last_row}"
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
            wp_schedule=wp_schedule,
            grid_range=grid_range,
            rate_col_range=rate_col_range,
            midx_range=midx_range,
            has_persons=num_persons > 0,
            year_specs=year_specs,
            grid_person_first_row=grid_person_first_row,
            grid_person_last_row=grid_person_last_row,
        )
    if num_wps > 0:
        for r in range(wp_summary_header_row, wp_summary_last_row + 1):
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
        for r in (
            grid_year_row, grid_month_row,
            grid_midx_row,
            *range(grid_person_first_row, grid_person_last_row + 1),
        ):
            ws.row_dimensions[r].outline_level = 1
    else:
        for r in (grid_year_row, grid_month_row, grid_midx_row):
            ws.row_dimensions[r].outline_level = 1

    # ── PER-PERSON SUMMARY band ──
    _write_section_title(
        ws, summary_title_row, SECTION_TITLE_PERSON_SUMMARY,
        last_col=_PERSON_SUMMARY_RATE_COL,
    )
    style_cell(
        ws, summary_header_row, 1, value="Person",
        font=HEADER_FONT, alignment=CENTER,
    )
    for i, hdr in enumerate(_PERSON_SUMMARY_METRIC_HEADERS):
        style_cell(
            ws, summary_header_row, _PERSON_SUMMARY_FIRST_METRIC_COL + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER,
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

    # ── Patch traffic-light formulas (downstream rows now exist) ──
    if num_wps > 0:
        # 1) Any undershoot? — Buffer € < 0 at col D of PER-WP SUMMARY.
        style_cell(
            ws, tl_row, tl_value_cols[0], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f'=COUNTIF(D{wp_summary_first_row}:D{wp_summary_last_row},"<0")'
        )
        # 2) Any out-of-window cell? — sum col H of PER-WP SUMMARY.
        style_cell(
            ws, tl_row, tl_value_cols[1], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = f"=SUM(H{wp_summary_first_row}:H{wp_summary_last_row})"
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

        # Rule 5 (PER-WP SUMMARY CF): Burn rate (col G) exceeds max team rate.
        burn_range = f"G{wp_summary_first_row}:G{wp_summary_last_row}"
        if num_persons > 0:
            rate_col_letter = get_column_letter(_PERSON_SUMMARY_RATE_COL)
            max_rate_expr = (
                f"MAX({rate_col_letter}{summary_person_first_row}"
                f":{rate_col_letter}{summary_person_last_row})"
            )
            ws.conditional_formatting.add(
                burn_range,
                FormulaRule(
                    formula=[
                        f"AND(ISNUMBER(G{wp_summary_first_row}),"
                        f"G{wp_summary_first_row}>{max_rate_expr})"
                    ],
                    fill=CF_RED_FILL,
                ),
            )

        # Rule 6 (PER-WP SUMMARY CF): Any out-of-window? > 0 → red. Col H.
        oow_range = f"H{wp_summary_first_row}:H{wp_summary_last_row}"
        ws.conditional_formatting.add(
            oow_range,
            FormulaRule(
                formula=[f"H{wp_summary_first_row}>0"],
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
        "resolve_flag_cell": (
            f"{get_column_letter(_TOGGLE_VALUE_COL)}{tl_row}"
        ),
    }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_wps(
    work_packages: list[dict],
    wp_schedule: list[dict],
    project_params: dict,
) -> list[WP]:
    """Convert config WP + schedule rows into solver WP records.

    Multi-segment WPs (e.g. WP3 with a mid-WP gap) are represented by their
    union span (min start, max end) — the solver treats the WP active window
    as contiguous because in-window-but-empty cells are still legal placements;
    the WP SCHEDULE gap remains visible on other sheets (Capacity, Plan).

    ``ceiling_eur`` mirrors the sheet's displayed `Config!E{wp_row}` =
    ``FTE × Rate``. Per-WP rate is taken from ``wp["rate"]`` when present
    (post-reader rebuild path); on init the seed WP dicts lack the key, so
    we fall back to ``project_params["rate"]`` — the same default used by
    ``config_sheet.py`` to populate the sheet.
    """
    default_rate = project_params.get("rate", 10000)
    if not isinstance(default_rate, (int, float)):
        default_rate = 10000
    wps: list[WP] = []
    for wp in work_packages:
        entries = [s for s in wp_schedule if s["id"] == wp["id"]]
        if not entries:
            continue
        active_start = min(int(s["start"]) for s in entries)
        active_end = max(int(s["end"]) for s in entries)
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


def _wp_sched_months(wp_code: str, wp_schedule: list[dict]) -> int:
    """Sum active months across every schedule segment for this WP code.

    Multi-segment WPs (e.g. a mid-project gap) count only the active
    segments — the gap is NOT a schedule month.
    """
    return sum(
        int(s["end"]) - int(s["start"]) + 1
        for s in wp_schedule if s["id"] == wp_code
    )


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
    """Compute the dynamic Y{yyyy} € column specs for PER-WP SUMMARY."""
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
    Solver never places WP codes there; on rebuild with resolve_flag=FALSE,
    any preserved WP code in a vacation cell is discarded (one-shot migration).

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


def _write_wp_summary_row(
    ws: Worksheet,
    row: int,
    wp: WP,
    *,
    cfg_wp_row: int,
    wp_schedule: list[dict],
    grid_range: str,
    rate_col_range: str,
    midx_range: str,
    has_persons: bool,
    year_specs: list["_YearColSpec"],
    grid_person_first_row: int,
    grid_person_last_row: int,
) -> None:
    """Fill the per-WP-summary cells on a single WP row.

    Cols A..H are the 8 base metrics; cols I.. hold one ``Y{yyyy} €`` column
    per calendar year spanned by the project. Year cells where the WP's
    active window does not overlap the year are blank (build-time check).
    """
    wp_code = wp.code

    # Col A: WP code (literal).
    style_cell(ws, row, 1, value=wp_code, font=BLACK_FORMULA, alignment=CENTER)

    # Col B: Ceiling € — green xref to Config!E{wp_row}.
    ceiling_ref = cell_ref(SHEET_CONFIG, cfg_wp_row, 5)
    style_cell(ws, row, 2, font=GREEN_XREF, fmt="€#,##0").value = f"={ceiling_ref}"

    # Col C: Assigned € = ROUND(SUMPRODUCT((grid="WP") * rate_vector), 0).
    # Rate vector is the PER-PERSON SUMMARY's col H over the person-row range;
    # grid rows match one-to-one with summary rows (same sort order).
    if has_persons:
        style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="€#,##0").value = (
            f'=ROUND(SUMPRODUCT(({grid_range}="{wp_code}")'
            f"*{rate_col_range}),0)"
        )
    else:
        style_cell(ws, row, 3, value=0, font=BLACK_FORMULA, fmt="€#,##0")

    # Col D: Buffer € = Assigned − Ceiling. Positive = buffer (good);
    # negative = shortfall (bad). CF green>0 / red<0.
    style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=C{row}-B{row}"
    )

    # Col E: Assigned months = SUMPRODUCT(1*(grid="WP")).
    if has_persons:
        style_cell(ws, row, 5, font=BLACK_FORMULA, fmt="0").value = (
            f'=SUMPRODUCT(1*({grid_range}="{wp_code}"))'
        )
    else:
        style_cell(ws, row, 5, value=0, font=BLACK_FORMULA, fmt="0")

    # Col F: WP sched months — build-time literal. Multi-segment WPs sum their
    # active segments only (gap months don't count).
    # Rule 3 carve-out: BLACK_FORMULA font on a literal — signals "builder-
    # written, do not edit." Cell comment echoes the deviation.
    sched_cell = style_cell(
        ws, row, 6, value=_wp_sched_months(wp_code, wp_schedule),
        font=BLACK_FORMULA, fmt="0", alignment=CENTER,
    )
    sched_cell.comment = sized_comment(
        "Number of months the WP SCHEDULE marks as active for this WP. "
        "Literal (not a formula) — recomputed on rebuild from Config.",
        author="Programme Budget Tool",
    )

    # Col G: Burn rate needed — IF-guarded; "" when buffer≥0 or wallclock=0.
    # wallclock = MAX(0, wp.active_end − current_month); buffer = D{row}.
    wallclock_expr = f"MAX(0,{wp.active_end}-current_month)"
    style_cell(ws, row, 7, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=IF(AND(D{row}<0,{wallclock_expr}>0),"
        f'-D{row}/{wallclock_expr},"")'
    )

    # Col H: Any out-of-window? = SUMPRODUCT((grid="WP") * OOW).
    if has_persons:
        style_cell(ws, row, 8, font=BLACK_FORMULA, fmt="0").value = (
            f'=SUMPRODUCT(({grid_range}="{wp_code}")*'
            f"(({midx_range}<{wp.active_start})+"
            f"({midx_range}>{wp.active_end})))"
        )
    else:
        style_cell(ws, row, 8, value=0, font=BLACK_FORMULA, fmt="0")

    # Cols I..: Y{yyyy} € — EUR placed by the solver in this WP × year slice.
    # (Still col I on PER-WP SUMMARY — this band has 8 base cols A..H; year
    # cols start at col I. Independent of the PER-PERSON SUMMARY Rate col.)
    # Overlap with the WP's active window is checked at build time; non-
    # overlap cells are blank. Rounded per the programme leaf-rounding rule.
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
    """Fill the 7 per-person-summary metric cells on cols B..H.

    All metrics target the merged grid's row for the same person — that's
    the canonical edit surface. Cols B..G are metrics; col H is Rate.
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
    # literal, Rule 3 carve-out — see `docs/plans/2026-04-19-timelog-plan-
    # enhancements-design.md` §3.4).
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

    # Col H: Rate — single-source VLOOKUP into roster_lookup. Consumed by
    # PER-WP SUMMARY Assigned-€ SUMPRODUCT as the rate vector.
    style_cell(ws, row, c0 + 6, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=VLOOKUP(A{row},roster_lookup,4,FALSE)"
    )
