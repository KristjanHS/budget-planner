"""Timelog Plan sheet builder — prescriptive 1.0-FTE WP allocation grid.

A separate allocation surface from the Plan sheet. For each (person, month)
the sheet prescribes which WP they should log 100% against.

Layout — vertical-stacked bands (one above the other, all on the same column
range ``D..D+months-1`` for the month grid). Top-to-bottom:

  1. Banner (row 1)
  2. Traffic lights (row 2)
  3. blank spacer
  4. PER-WP SUMMARY band
     blank spacer
  5. PROPOSED band (solver-written, locked)
     blank spacer
  6. ASSIGNED band (defaults to ``=<Proposed>``; blue-input overrides accepted)
     blank spacer
  7. PER-PERSON SUMMARY band

Each data band is wrapped in an Excel row-group (``outline_level=1``) so the
user can fold away sections not under review. Section-title rows stay at
level 0 and remain visible when the group collapses.

Col A is Person (cross-ref into Config Team Roster) across all three person-
indexed bands. Col B is Underload (Pool − Assigned) — always visible in the
frozen pane so the user can eyeball utilisation while scrolling months.
Col C is the rate helper (VLOOKUP into ``roster_lookup``) on PROPOSED and
ASSIGNED bands; on the PER-PERSON SUMMARY band col C is the Assigned count
metric instead.

Spec: ``docs/plans/2026-04-18-timelog-plan-vertical-layout-design.md``.

Stateless builder — receives all config via kwargs from the orchestrator. Per
``.claude/rules/python-refactors.md``: no ``from budget_planner import config``
fallback inside this builder; the orchestrator owns the lazy-load.

Known deferred work (also flagged in design §9 follow-ups):
- Per-person ``Largest gap`` summary column is a placeholder ``""`` cell with
  a comment. Implementing it cleanly needs a hidden streak-length helper row
  and a MAX over it; deferred until there's a concrete user need.
"""

from __future__ import annotations

from collections.abc import Mapping

from dateutil.relativedelta import relativedelta
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from budget_planner.layout import (
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
    TIMELOG_RATE_HELPER_COL,
    cell_ref,
    month_range_ref,
)
from budget_planner.styles import (
    BLACK_FORMULA,
    BLUE_INPUT,
    CENTER,
    CF_PASTEL_FILLS,
    CF_RED_FILL,
    DARK_GREY_FILL,
    GREEN_XREF,
    GREY_FILL,
    HEADER_FONT,
    LIGHT_BLUE_FILL,
    LIGHT_GREEN_FILL,
    SECTION_FILL,
    SECTION_FONT,
    TITLE_FONT,
    col_letter,
    sized_comment,
    style_cell,
)
from budget_planner.timelog_solver import Person, PersonWPWindow, WP, solve

# Section-title strings — exported so the reader can anchor on "ASSIGNED…"
# via the same literal the builder writes. Update here, not at the callsite.
SECTION_TITLE_WP_SUMMARY = "PER-WP SUMMARY"
SECTION_TITLE_PROPOSED = "PROPOSED (solver-written, locked)"
SECTION_TITLE_ASSIGNED = "ASSIGNED (edit here — default =Proposed)"
SECTION_TITLE_PERSON_SUMMARY = "PER-PERSON SUMMARY"

# PER-PERSON SUMMARY metric columns live at C..I (7 cols). Col A is Person,
# col B is Underload — the same layout roles they carry in PROPOSED/ASSIGNED.
# Dropping the duplicated ``Remaining`` column vs. the prior horizontal layout
# makes col B (Underload) the single source of truth for Pool − Assigned.
_PERSON_SUMMARY_METRIC_HEADERS = [
    "Assigned", "Pool", "Over-pool?",
    "Switches", "First log", "Last log", "Largest gap",
]
_PERSON_SUMMARY_FIRST_METRIC_COL = 3   # col C (Assigned)

# PER-WP SUMMARY metric columns live at A..H (WP code + 7 metrics).
_WP_SUMMARY_HEADERS = [
    "WP code", "Ceiling €", "Assigned €", "Overshoot €",
    "Undershoot €", "Assigned months", "Window months remaining",
    "Any out-of-window?",
]

# Font colours for the Underload CF (red < 0, green > 0, black default).
# Fonts constructed inline here rather than shared in styles.py — they are
# CF-local and don't follow the blue/black/green cell-convention.
_UNDERLOAD_RED_FONT = Font(name="Arial", color="C00000", size=10, bold=True)
_UNDERLOAD_GREEN_FONT = Font(name="Arial", color="006100", size=10)


# ---------------------------------------------------------------------------
# Public builder
# ---------------------------------------------------------------------------


def build_timelog_plan_sheet(
    ws: Worksheet,
    wb: Workbook,  # noqa: ARG001 — kept for parity with sibling builders
    *,
    project_params: dict,
    work_packages: list[dict],
    wp_schedule: list[dict],
    team: list[dict],
    assignments: list[dict],
    grade_rates: Mapping[str, float],  # noqa: ARG001 — reserved for future per-grade tinting
    config_layout: dict,
    assigned_overrides: Mapping[tuple[str, int], str] | None = None,
) -> dict:
    """Build the Timelog Plan sheet with vertical-stacked bands.

    Args:
        ws: Worksheet to populate (created by orchestrator, named "Timelog Plan").
        wb: Parent workbook (kept for parity with sibling builders; this
            builder does not currently add named ranges).
        project_params: dict with keys 'title', 'start', 'months', 'rate', etc.
        work_packages: list of WP dicts with keys 'id', 'name', 'fte', 'rate'.
        wp_schedule: list of WP schedule entries (id, start, end, type).
        team: list of team-member dicts with keys 'name', 'max_fte', 'grade',
            'priority' (1-3, defaults to 2 from the reader).
        assignments: list of person-WP assignment dicts (person, wp,
            total_fte, start, end). Treated as the solver's hard
            (person, wp) → window constraints.
        grade_rates: dict grade_name → EUR/PM. Currently unused by this
            builder (per-person rate is looked up from roster_lookup via
            the rate helper column); accepted for orchestrator-call parity.
        config_layout: dict returned by build_config_sheet() — must include
            ``param_rows``, ``wp_first_row``, ``person_rows``.
        assigned_overrides: Optional mapping ``(person_name, month_num)`` →
            WP code. Cells present here are written as BLUE_INPUT literals
            (preserving round-trip user edits); cells absent default to
            ``=<Proposed cell>`` formulas. Rebuild flow: reader extracts
            literals from the previous workbook via ``read_timelog_plan``.

    Returns:
        Layout dict with keys:
          sheet, month_col_start, month_col_end,
          proposed_person_first_row, proposed_person_last_row,
          assigned_person_first_row, assigned_person_last_row,
          summary_person_first_row, summary_person_last_row,
          person_rows_by_band (dict band → {name: row}),
          wp_summary_first_row, wp_summary_last_row.
    """
    project_months: int = project_params["months"]

    # ── Tab appearance ──
    ws.sheet_properties.tabColor = "FF8C00"  # orange — distinguishes from existing tabs

    # ── Column geometry (shared by PROPOSED and ASSIGNED bands) ──
    month_col_start = TIMELOG_MONTH_COL_START
    month_col_end = month_col_start + project_months - 1
    canvas_last_col = month_col_end

    # ── Solver input prep ──
    # Sort persons (priority asc, name asc). Persists the solver's pass order
    # into the sheet so row order matches solver iteration order.
    sorted_team = sorted(
        team, key=lambda m: (int(m.get("priority", 2)), m["name"])
    )
    persons = [
        Person(
            name=m["name"],
            priority=int(m.get("priority", 2)),
            pool=round(float(m["max_fte"]) * project_months),
        )
        for m in sorted_team
    ]
    wps = _build_wps(work_packages, wp_schedule)
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
    person_windows: dict[str, list[tuple[str, int, int]]] = {}
    for w in windows:
        person_windows.setdefault(w.person_name, []).append(
            (w.wp_code, w.start_month, w.end_month)
        )
    override_map: Mapping[tuple[str, int], str] = assigned_overrides or {}
    num_persons = len(persons)
    num_wps = len(wps)

    # ── Row layout — pre-computed so forward cross-band refs resolve ──
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
    proposed_title_row = wp_summary_last_row + 2  # +1 spacer

    proposed_year_row = proposed_title_row + 1
    proposed_month_row = proposed_year_row + 1
    proposed_midx_row = proposed_month_row + 1
    proposed_person_first_row = proposed_midx_row + 1
    proposed_person_last_row = (
        proposed_person_first_row + num_persons - 1 if num_persons > 0
        else proposed_person_first_row - 1
    )
    assigned_title_row = (
        (proposed_person_last_row if num_persons > 0 else proposed_midx_row) + 2
    )

    assigned_year_row = assigned_title_row + 1
    assigned_month_row = assigned_year_row + 1
    assigned_midx_row = assigned_month_row + 1
    assigned_person_first_row = assigned_midx_row + 1
    assigned_person_last_row = (
        assigned_person_first_row + num_persons - 1 if num_persons > 0
        else assigned_person_first_row - 1
    )
    summary_title_row = (
        (assigned_person_last_row if num_persons > 0 else assigned_midx_row) + 2
    )

    summary_header_row = summary_title_row + 1
    summary_person_first_row = summary_header_row + 1
    summary_person_last_row = (
        summary_person_first_row + num_persons - 1 if num_persons > 0
        else summary_person_first_row - 1
    )

    person_rows_by_band: dict[str, dict[str, int]] = {
        "proposed": {
            p.name: proposed_person_first_row + i for i, p in enumerate(persons)
        },
        "assigned": {
            p.name: assigned_person_first_row + i for i, p in enumerate(persons)
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

    # ── Row 2: traffic-light labels (values patched in after WP rows exist) ──
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

    # row 3: blank spacer (no write)

    # ── PER-WP SUMMARY band ──
    _write_section_title(
        ws, wp_summary_title_row, SECTION_TITLE_WP_SUMMARY, last_col=8,
    )
    for i, hdr in enumerate(_WP_SUMMARY_HEADERS):
        style_cell(
            ws, wp_summary_header_row, 1 + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER,
        )
    cfg_wp_first_row = config_layout["wp_first_row"]
    for i, wp in enumerate(wps):
        _write_wp_summary_row(
            ws, wp_summary_first_row + i, wp,
            cfg_wp_row=cfg_wp_first_row + i,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            assigned_first_row=assigned_person_first_row,
            assigned_last_row=assigned_person_last_row,
            assigned_midx_row=assigned_midx_row,
        )
    if num_wps > 0:
        for r in range(wp_summary_header_row, wp_summary_last_row + 1):
            ws.row_dimensions[r].outline_level = 1

    # ── PROPOSED band ──
    _write_section_title(
        ws, proposed_title_row, SECTION_TITLE_PROPOSED, last_col=canvas_last_col,
    )
    _write_band_axis_headers(
        ws, proposed_year_row, proposed_month_row, project_params, month_col_start,
    )
    _write_month_index_row(
        ws, proposed_midx_row, project_months, month_col_start,
    )
    if num_persons > 0:
        for r in (
            proposed_year_row, proposed_month_row,
            proposed_midx_row, *range(proposed_person_first_row, proposed_person_last_row + 1),
        ):
            ws.row_dimensions[r].outline_level = 1
    else:
        for r in (proposed_year_row, proposed_month_row, proposed_midx_row):
            ws.row_dimensions[r].outline_level = 1

    # ── ASSIGNED band ──
    _write_section_title(
        ws, assigned_title_row, SECTION_TITLE_ASSIGNED, last_col=canvas_last_col,
    )
    _write_band_axis_headers(
        ws, assigned_year_row, assigned_month_row, project_params, month_col_start,
    )
    _write_month_index_row(
        ws, assigned_midx_row, project_months, month_col_start,
    )
    if num_persons > 0:
        for r in (
            assigned_year_row, assigned_month_row,
            assigned_midx_row, *range(assigned_person_first_row, assigned_person_last_row + 1),
        ):
            ws.row_dimensions[r].outline_level = 1
    else:
        for r in (assigned_year_row, assigned_month_row, assigned_midx_row):
            ws.row_dimensions[r].outline_level = 1

    # ── PER-PERSON SUMMARY band ──
    _write_section_title(
        ws, summary_title_row, SECTION_TITLE_PERSON_SUMMARY,
        last_col=_PERSON_SUMMARY_FIRST_METRIC_COL + len(_PERSON_SUMMARY_METRIC_HEADERS) - 1,
    )
    # Col A + B headers (Person + Underload) on the summary-metric header row.
    style_cell(
        ws, summary_header_row, 1, value="Person",
        font=HEADER_FONT, alignment=CENTER,
    )
    style_cell(
        ws, summary_header_row, 2, value="Underload",
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

    # ── Fill person rows in all three bands ──
    cfg_person_rows: dict[str, int] = config_layout.get("person_rows", {})
    months_row_cfg = config_layout["param_rows"]["Project Months"]
    project_months_ref = cell_ref(SHEET_CONFIG, months_row_cfg, 2)

    for i, person in enumerate(persons):
        name = person.name
        proposed_row = person_rows_by_band["proposed"][name]
        assigned_row = person_rows_by_band["assigned"][name]
        summary_row = person_rows_by_band["summary"][name]
        cfg_row = cfg_person_rows.get(name)
        person_assignment = assigned_map.get(name, {})
        windows_for_person = person_windows.get(name, [])

        # PROPOSED band: Person xref + Underload + rate helper + month cells.
        _write_person_col_a(ws, proposed_row, name, cfg_row)
        _write_underload_cell(
            ws, proposed_row,
            summary_row=summary_row,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
        )
        # Rate helper on PROPOSED: informational/cosmetic (row-aligned with
        # the ASSIGNED grid consumer, which SUMPRODUCTs across the ASSIGNED
        # band's rate column — this PROPOSED-band rate cell is duplicated so
        # the band reads self-contained).
        _write_rate_helper_cell(ws, proposed_row)
        for mi in range(project_months):
            month_num = mi + 1
            col = month_col_start + mi
            wp_code = person_assignment.get(month_num)
            in_any_window = any(
                start <= month_num <= end for (_wp, start, end)
                in windows_for_person
            )
            _write_proposed_cell(ws, proposed_row, col, wp_code, in_any_window)

        # ASSIGNED band: Person xref + Underload + rate helper + month cells.
        _write_person_col_a(ws, assigned_row, name, cfg_row)
        _write_underload_cell(
            ws, assigned_row,
            summary_row=summary_row,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
        )
        _write_rate_helper_cell(ws, assigned_row)
        for mi in range(project_months):
            month_num = mi + 1
            col = month_col_start + mi
            wp_code = person_assignment.get(month_num)
            in_any_window = any(
                start <= month_num <= end for (_wp, start, end)
                in windows_for_person
            )
            _write_assigned_cell(
                ws, assigned_row, col, proposed_row, wp_code, in_any_window,
                override=override_map.get((name, month_num)),
            )

        # PER-PERSON SUMMARY band: Person xref + Underload + 7 metric cols.
        _write_person_col_a(ws, summary_row, name, cfg_row)
        _write_summary_underload_cell(ws, summary_row)
        _write_person_summary_metrics(
            ws, summary_row,
            assigned_row=assigned_row,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            month_header_row=assigned_month_row,
            project_months_ref=project_months_ref,
            cfg_row=cfg_row,
        )

    # ── Patch traffic-light formulas (values now that downstream rows exist) ──
    if num_wps > 0:
        # 1) Any undershoot? — count WPs where Undershoot € > 0 (col E).
        style_cell(
            ws, tl_row, tl_value_cols[0], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f'=COUNTIF(E{wp_summary_first_row}:E{wp_summary_last_row},">0")'
        )
        # 2) Any out-of-window cell? — sum across per-WP out-of-window counts (col H).
        style_cell(
            ws, tl_row, tl_value_cols[1], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = f"=SUM(H{wp_summary_first_row}:H{wp_summary_last_row})"
    else:
        # No WPs → no signal; leave the value cells blank but styled.
        for c in tl_value_cols[:2]:
            ws.cell(row=tl_row, column=c, value=0)

    if num_persons > 0:
        # 3) Any over-pool person? — count persons where Over-pool? = TRUE.
        # Over-pool? lives in col E of the PER-PERSON SUMMARY band.
        over_pool_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 2)
        style_cell(
            ws, tl_row, tl_value_cols[2], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f"=COUNTIF({over_pool_letter}{summary_person_first_row}"
            f":{over_pool_letter}{summary_person_last_row},TRUE)"
        )
        # 4) Total switches — sum of per-person switch counts (col F of summary).
        switches_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 3)
        style_cell(
            ws, tl_row, tl_value_cols[3], font=BLACK_FORMULA,
            alignment=CENTER, fmt="0",
        ).value = (
            f"=SUM({switches_letter}{summary_person_first_row}"
            f":{switches_letter}{summary_person_last_row})"
        )
    else:
        for c in tl_value_cols[2:]:
            ws.cell(row=tl_row, column=c, value=0)

    # CF on first three traffic-light value cells: red when > 0.
    for c in tl_value_cols[:3]:
        cell_addr = f"{get_column_letter(c)}{tl_row}"
        ws.conditional_formatting.add(
            cell_addr,
            FormulaRule(formula=[f"{cell_addr}>0"], fill=CF_RED_FILL),
        )

    # ── CF rules (canvas) — most-critical first ──
    # Per ``openpyxl-patterns.md``: first rule added gets highest priority.
    if num_persons > 0 and num_wps > 0:
        proposed_grid_range = _grid_range_ref(
            month_col_start, month_col_end,
            proposed_person_first_row, proposed_person_last_row,
        )
        assigned_grid_range = _grid_range_ref(
            month_col_start, month_col_end,
            assigned_person_first_row, assigned_person_last_row,
        )

        # Rule 1: per-WP cell-level red for WP code outside its active window.
        # One rule per WP × (proposed, assigned) grid. Each grid's CF formula
        # references that grid's own month-idx helper row.
        for wp in wps:
            for grid_range, midx_row in (
                (proposed_grid_range, proposed_midx_row),
                (assigned_grid_range, assigned_midx_row),
            ):
                top_left = grid_range.split(":")[0]
                tl_col_letter = "".join(ch for ch in top_left if ch.isalpha())
                tl_row_num = "".join(ch for ch in top_left if ch.isdigit())
                idx_ref = f"{tl_col_letter}${midx_row}"
                cell_ref_inline = f"{tl_col_letter}{tl_row_num}"
                formula = (
                    f'AND({cell_ref_inline}="{wp.code}",'
                    f"OR({idx_ref}<{wp.active_start},{idx_ref}>{wp.active_end}))"
                )
                ws.conditional_formatting.add(
                    grid_range,
                    FormulaRule(formula=[formula], fill=CF_RED_FILL),
                )

        # Rule 2: pastel fill per WP across both grids.
        for i, wp in enumerate(wps):
            pastel = CF_PASTEL_FILLS[i % len(CF_PASTEL_FILLS)]
            for grid_range in (proposed_grid_range, assigned_grid_range):
                top_left = grid_range.split(":")[0]
                tl_col_letter = "".join(ch for ch in top_left if ch.isalpha())
                tl_row_num = "".join(ch for ch in top_left if ch.isdigit())
                cell_ref_inline = f"{tl_col_letter}{tl_row_num}"
                ws.conditional_formatting.add(
                    grid_range,
                    FormulaRule(
                        formula=[f'{cell_ref_inline}="{wp.code}"'],
                        fill=pastel,
                    ),
                )

    # Rule 3 (per-person summary CF): Over-pool? = TRUE → red.
    if num_persons > 0:
        op_col_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 2)
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

    # Rule 4 (per-WP summary CF): Undershoot € > 0 → red.
    if num_wps > 0:
        us_range = f"E{wp_summary_first_row}:E{wp_summary_last_row}"
        ws.conditional_formatting.add(
            us_range,
            FormulaRule(
                formula=[f"E{wp_summary_first_row}>0"],
                fill=CF_RED_FILL,
            ),
        )
        # Rule 5 (per-WP summary CF): Any out-of-window? > 0 → red.
        oow_range = f"H{wp_summary_first_row}:H{wp_summary_last_row}"
        ws.conditional_formatting.add(
            oow_range,
            FormulaRule(
                formula=[f"H{wp_summary_first_row}>0"],
                fill=CF_RED_FILL,
            ),
        )

    # Rule 6 (Underload col B font colour per band — red < 0, green > 0).
    # Default font colour (BLACK_FORMULA) shows through at 0.
    if num_persons > 0:
        for first, last in (
            (proposed_person_first_row, proposed_person_last_row),
            (assigned_person_first_row, assigned_person_last_row),
            (summary_person_first_row, summary_person_last_row),
        ):
            band_range = f"B{first}:B{last}"
            ws.conditional_formatting.add(
                band_range,
                FormulaRule(
                    formula=[f"B{first}<0"], font=_UNDERLOAD_RED_FONT,
                ),
            )
            ws.conditional_formatting.add(
                band_range,
                FormulaRule(
                    formula=[f"B{first}>0"], font=_UNDERLOAD_GREEN_FONT,
                ),
            )

    # ── Freeze panes + column widths + outline ──
    # Freeze rows 1–2 (banner + traffic lights) and cols A–C (Person + Underload
    # + Rate helper). Month grid scrolls horizontally; Underload stays visible.
    ws.freeze_panes = ws.cell(
        row=3, column=TIMELOG_MONTH_COL_START,
    ).coordinate  # "D3"

    ws.column_dimensions["A"].width = 25          # Person
    ws.column_dimensions["B"].width = 10          # Underload
    ws.column_dimensions["C"].width = 6           # Rate helper
    ws.column_dimensions["C"].outline_level = 1   # collapsible
    # Metric cols on PER-PERSON SUMMARY are at D..J — a little wider for headers.
    for i in range(len(_PERSON_SUMMARY_METRIC_HEADERS)):
        c_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + i)
        ws.column_dimensions[c_letter].width = 12

    return {
        "sheet": SHEET_TIMELOG_PLAN,
        "month_col_start": month_col_start,
        "month_col_end": month_col_end,
        "proposed_person_first_row": proposed_person_first_row,
        "proposed_person_last_row": proposed_person_last_row,
        "assigned_person_first_row": assigned_person_first_row,
        "assigned_person_last_row": assigned_person_last_row,
        "summary_person_first_row": summary_person_first_row,
        "summary_person_last_row": summary_person_last_row,
        "person_rows_by_band": person_rows_by_band,
        "wp_summary_first_row": wp_summary_first_row,
        "wp_summary_last_row": wp_summary_last_row,
    }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_wps(
    work_packages: list[dict],
    wp_schedule: list[dict],
) -> list[WP]:
    """Convert config WP + schedule rows into solver WP records.

    Multi-segment WPs (e.g. WP3 with a mid-WP gap) are represented by their
    union span (min start, max end) — the solver treats the WP active window
    as contiguous because in-window-but-empty cells are still legal placements;
    the WP SCHEDULE gap remains visible on other sheets (Capacity, Plan).
    """
    wps: list[WP] = []
    for wp in work_packages:
        entries = [s for s in wp_schedule if s["id"] == wp["id"]]
        if not entries:
            continue
        active_start = min(int(s["start"]) for s in entries)
        active_end = max(int(s["end"]) for s in entries)
        ceiling_months = round(float(wp["fte"]))
        wps.append(WP(
            code=wp["id"],
            active_start=active_start,
            active_end=active_end,
            ceiling_months=ceiling_months,
        ))
    return wps


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
    # Fill rest of the merge so styling reads cleanly if merge target cells
    # lose styling (openpyxl quirk on re-open).
    for c in range(2, last_col + 1):
        style_cell(ws, row, c, fill=SECTION_FILL, border=False)


def _write_band_axis_headers(
    ws: Worksheet,
    year_row: int,
    month_row: int,
    project_params: dict,
    month_col_start: int,
) -> None:
    """Write the year + month headers for one band (PROPOSED or ASSIGNED).

    Self-contained per band — so collapsing one band still shows the other's
    axis. Col A/B/C (Person / Underload / Rate) header cells are styled on
    the month row; the year row has LIGHT_BLUE_FILL across cols A..C.
    """
    project_months = project_params["months"]
    project_start = project_params["start"]

    # Year row: LIGHT_BLUE_FILL on cols A-C, then calendar-year blocks from
    # ``month_col_start`` onward.
    for c in range(1, month_col_start):
        style_cell(ws, year_row, c, value="", fill=LIGHT_BLUE_FILL)
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

    # Month row: "Person" / "Underload" / "Rate" headers in A/B/C, then
    # ``M1-Jan26``-style per-month labels.
    style_cell(
        ws, month_row, 1, value="Person", font=HEADER_FONT,
        fill=SECTION_FILL, alignment=CENTER,
    ).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    style_cell(
        ws, month_row, 2, value="Underload", font=HEADER_FONT,
        fill=SECTION_FILL, alignment=CENTER,
    ).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    style_cell(
        ws, month_row, TIMELOG_RATE_HELPER_COL, value="Rate",
        font=HEADER_FONT, fill=SECTION_FILL, alignment=CENTER,
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
    """Yield (group_start_month_index, span) tuples — mirrors layout.py's helper.

    Inlined (small + tightly coupled to the band-axis header writer) rather
    than imported from layout.py to keep the builder self-contained for this
    one use site.
    """
    m = 0
    while m < project_months:
        cal_date = project_start + relativedelta(months=m)
        months_left_in_year = 12 - cal_date.month + 1
        span = min(months_left_in_year, project_months - m)
        yield m, span
        m += span


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
        # Fallback: literal name (should not happen if orchestrator passes a
        # config_layout aligned with `team`).
        style_cell(ws, row, 1, value=name, font=GREEN_XREF)


def _write_rate_helper_cell(ws: Worksheet, row: int) -> None:
    """Write the per-row VLOOKUP into roster_lookup col 4 (EUR/PM)."""
    style_cell(
        ws, row, TIMELOG_RATE_HELPER_COL,
        font=BLACK_FORMULA, fmt="€#,##0",
    ).value = f"=VLOOKUP(A{row},roster_lookup,4,FALSE)"


def _write_underload_cell(
    ws: Worksheet,
    row: int,
    *,
    summary_row: int,
    month_col_start: int,
    month_col_end: int,
) -> None:
    """Write col B on a PROPOSED or ASSIGNED person row.

    Formula: ``=$D{summary_row} − COUNTIF(<grid-row-range>, "WP*")``.
    Forward-references the PER-PERSON SUMMARY band's Pool cell (col D =
    ``_PERSON_SUMMARY_FIRST_METRIC_COL + 1``). Excel resolves forward refs
    at recalc — no chain issue.
    """
    pool_col_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 1)  # D
    grid_range = month_range_ref(row, month_col_start, month_col_end)
    style_cell(ws, row, 2, font=BLACK_FORMULA, fmt="0").value = (
        f'=${pool_col_letter}{summary_row}-COUNTIF({grid_range},"WP*")'
    )


def _write_summary_underload_cell(ws: Worksheet, row: int) -> None:
    """Write col B on a PER-PERSON SUMMARY row: Pool − Assigned (same-row cells)."""
    assigned_col_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL)      # C
    pool_col_letter = get_column_letter(_PERSON_SUMMARY_FIRST_METRIC_COL + 1)      # D
    style_cell(ws, row, 2, font=BLACK_FORMULA, fmt="0").value = (
        f"={pool_col_letter}{row}-{assigned_col_letter}{row}"
    )


def _grid_range_ref(
    col_start: int, col_end: int,
    row_start: int, row_end: int,
) -> str:
    """Return ``"D7:AV20"`` for a rectangular range."""
    return (
        f"{get_column_letter(col_start)}{row_start}"
        f":{get_column_letter(col_end)}{row_end}"
    )


def _write_proposed_cell(
    ws: Worksheet,
    row: int,
    col: int,
    wp_code: str | None,
    in_any_window: bool,
) -> None:
    """Write a single Proposed-grid cell.

    Locked. Pre-fills cells outside any person-WP window with DARK_GREY_FILL
    and in-window-but-empty cells with GREY_FILL — deterministic visual that
    doesn't depend on CF or LO recalc.
    """
    if wp_code is not None:
        cell = style_cell(
            ws, row, col, value=wp_code,
            font=BLACK_FORMULA, alignment=CENTER,
        )
    elif in_any_window:
        cell = style_cell(
            ws, row, col, value="",
            font=BLACK_FORMULA, fill=GREY_FILL, alignment=CENTER,
        )
    else:
        cell = style_cell(
            ws, row, col, value="",
            font=BLACK_FORMULA, fill=DARK_GREY_FILL, alignment=CENTER,
        )
    cell.protection = Protection(locked=True)


def _write_assigned_cell(
    ws: Worksheet,
    row: int,
    col: int,
    proposed_row: int,
    wp_code: str | None,
    in_any_window: bool,
    *,
    override: str | None = None,
) -> None:
    """Write a single Assigned-grid cell.

    Defaults to ``=<ColLetter>{proposed_row}`` — a cross-band ref into the
    PROPOSED band's same-column cell for this person. Cells outside any
    person-WP window are pre-greyed (matching Proposed) and unlocked so the
    user can override. When ``override`` is supplied (round-trip from a prior
    user edit), the literal value is written as BLUE_INPUT.

    Rule 3 carve-out: the default-formula cells are coloured BLUE_INPUT even
    though they hold formulas, because this band IS the editable input
    surface — the colour tells the user "you may overwrite this cell." The
    formula default is a scaffolding default, not a derived value.
    """
    col_ltr = get_column_letter(col)
    if override is not None:
        if not in_any_window:
            cell = style_cell(
                ws, row, col, value=override,
                font=BLUE_INPUT, fill=DARK_GREY_FILL, alignment=CENTER,
            )
        else:
            cell = style_cell(
                ws, row, col, value=override,
                font=BLUE_INPUT, alignment=CENTER,
            )
        cell.protection = Protection(locked=False)
        return

    formula = f"={col_ltr}{proposed_row}"
    if wp_code is None and not in_any_window:
        cell = style_cell(
            ws, row, col, value=formula,
            font=BLUE_INPUT, fill=DARK_GREY_FILL, alignment=CENTER,
        )
    elif wp_code is None:
        cell = style_cell(
            ws, row, col, value=formula,
            font=BLUE_INPUT, fill=GREY_FILL, alignment=CENTER,
        )
    else:
        cell = style_cell(
            ws, row, col, value=formula,
            font=BLUE_INPUT, alignment=CENTER,
        )
    cell.protection = Protection(locked=False)


def _write_wp_summary_row(
    ws: Worksheet,
    row: int,
    wp: WP,
    *,
    cfg_wp_row: int,
    month_col_start: int,
    month_col_end: int,
    assigned_first_row: int,
    assigned_last_row: int,
    assigned_midx_row: int,
) -> None:
    """Fill the 8 per-WP-summary cells on a single WP row (cols A..H)."""
    wp_code = wp.code
    rate_helper_letter = get_column_letter(TIMELOG_RATE_HELPER_COL)

    # Col A: WP code (literal).
    style_cell(ws, row, 1, value=wp_code, font=BLACK_FORMULA, alignment=CENTER)

    # Col B: Ceiling € — green xref to Config!E{wp_row}.
    ceiling_ref = cell_ref(SHEET_CONFIG, cfg_wp_row, 5)
    style_cell(ws, row, 2, font=GREEN_XREF, fmt="€#,##0").value = f"={ceiling_ref}"

    # Col C: Assigned € = SUMPRODUCT((assigned="WP") * rate_vector_broadcast).
    # Rate vector is the ASSIGNED band's col C over the person-row range —
    # row-aligned with the assigned grid, so SUMPRODUCT broadcasts rate across
    # months for each person row.
    assigned_grid_range = _grid_range_ref(
        month_col_start, month_col_end,
        assigned_first_row, assigned_last_row,
    )
    rate_col_range = (
        f"{rate_helper_letter}{assigned_first_row}"
        f":{rate_helper_letter}{assigned_last_row}"
    )
    # ROUND-wrap EUR leaves so fractional cents from non-integer PM rates
    # don't accumulate into auditor mismatches downstream.
    style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=ROUND(SUMPRODUCT(({assigned_grid_range}=\"{wp_code}\")"
        f"*{rate_col_range}),0)"
    )

    # Col D: Overshoot € = MAX(0, Assigned - Ceiling).
    style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=MAX(0,C{row}-B{row})"
    )
    # Col E: Undershoot € = MAX(0, Ceiling - Assigned).
    style_cell(ws, row, 5, font=BLACK_FORMULA, fmt="€#,##0").value = (
        f"=MAX(0,B{row}-C{row})"
    )
    # Col F: Assigned months = SUMPRODUCT(1*(assigned="WP")).
    style_cell(ws, row, 6, font=BLACK_FORMULA, fmt="0").value = (
        f'=SUMPRODUCT(1*({assigned_grid_range}="{wp_code}"))'
    )
    # Col G: Window months remaining = MAX(0, wp_end - current_month).
    style_cell(ws, row, 7, font=BLACK_FORMULA, fmt="0").value = (
        f"=MAX(0,{wp.active_end}-current_month)"
    )
    # Col H: Any out-of-window? = SUMPRODUCT((assigned="WP") * OOW).
    month_idx_range = (
        f"{get_column_letter(month_col_start)}{assigned_midx_row}"
        f":{get_column_letter(month_col_end)}{assigned_midx_row}"
    )
    style_cell(ws, row, 8, font=BLACK_FORMULA, fmt="0").value = (
        f'=SUMPRODUCT(({assigned_grid_range}="{wp_code}")*'
        f"(({month_idx_range}<{wp.active_start})+"
        f"({month_idx_range}>{wp.active_end})))"
    )


def _write_person_summary_metrics(
    ws: Worksheet,
    row: int,
    *,
    assigned_row: int,
    month_col_start: int,
    month_col_end: int,
    month_header_row: int,
    project_months_ref: str,
    cfg_row: int | None,
) -> None:
    """Fill the 7 per-person-summary metric cells on cols C..I.

    All metrics target the ASSIGNED band's row for the same person ({a})
    since that's the canonical edit surface. Col B (Underload) is written
    by ``_write_summary_underload_cell``.
    """
    c0 = _PERSON_SUMMARY_FIRST_METRIC_COL

    assigned_row_range = month_range_ref(
        assigned_row, month_col_start, month_col_end,
    )
    assigned_left_range = month_range_ref(
        assigned_row, month_col_start, month_col_end - 1,
    )
    assigned_right_range = month_range_ref(
        assigned_row, month_col_start + 1, month_col_end,
    )
    month_headers_range = month_range_ref(
        month_header_row, month_col_start, month_col_end,
    )

    # Col C: Assigned = COUNTIF(row, "WP*").
    style_cell(ws, row, c0, font=BLACK_FORMULA, fmt="0").value = (
        f'=COUNTIF({assigned_row_range},"WP*")'
    )

    # Col D: Pool = ROUND(max_fte * project_months, 0). Green xref.
    if cfg_row is not None:
        max_fte_ref = cell_ref(SHEET_CONFIG, cfg_row, 2)
        style_cell(ws, row, c0 + 1, font=GREEN_XREF, fmt="0").value = (
            f"=ROUND({max_fte_ref}*{project_months_ref},0)"
        )
    else:
        style_cell(ws, row, c0 + 1, value=0, font=BLACK_FORMULA, fmt="0")

    # Col E: Over-pool? = Assigned > Pool.
    style_cell(
        ws, row, c0 + 2, font=BLACK_FORMULA, alignment=CENTER,
    ).value = f"={col_letter(c0)}{row}>{col_letter(c0 + 1)}{row}"

    # Col F: Switches = SUMPRODUCT(1*(left <> right)).
    style_cell(ws, row, c0 + 3, font=BLACK_FORMULA, fmt="0").value = (
        f"=SUMPRODUCT(1*({assigned_left_range}<>{assigned_right_range}))"
    )

    # Col G: First log = IFERROR(INDEX(headers, MATCH(1, (row<>"")*1, 0)), "").
    # *1 coercion (not literal TRUE) — works as regular (non-CSE) entry in
    # pre-365 Excel; see openpyxl-patterns.md on silent-CSE MATCH pitfall.
    style_cell(
        ws, row, c0 + 4, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f"=IFERROR(INDEX({month_headers_range},"
        f'MATCH(1,({assigned_row_range}<>"")*1,0)),"")'
    )

    # Col H: Last log = IFERROR(LOOKUP(2, 1/(row<>""), headers), "").
    style_cell(
        ws, row, c0 + 5, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f"=IFERROR(LOOKUP(2,1/({assigned_row_range}<>\"\"),"
        f'{month_headers_range}),"")'
    )

    # Col I: Largest gap — placeholder, see module docstring.
    cell = style_cell(
        ws, row, c0 + 6, value="",
        font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, alignment=CENTER,
    )
    cell.comment = sized_comment(
        "Largest gap calc deferred — see solver design Q follow-ups "
        "(needs hidden streak-length helper row + MAX). Placeholder cell.",
        author="Claude",
    )
