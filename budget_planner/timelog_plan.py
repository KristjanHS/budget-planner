"""Timelog Plan sheet builder — prescriptive 1.0-FTE WP allocation grid.

A separate allocation surface from the Plan sheet. For each (person, month)
the sheet prescribes which WP they should log 100% against. The Proposed grid
is written by the solver and locked; the Assigned grid mirrors Proposed via
``=<cell>`` formulas and accepts blue-input overrides.

Spec: ``docs/plans/2026-04-18-timelog-plan-design.md`` (§2 sheet layout, §5
feedback formulas + CF, §6 build/read/rebuild).

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

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
    TIMELOG_RATE_HELPER_COL,
    build_year_month_headers,
    cell_ref,
    month_range_ref,
    timelog_assigned_col_end,
    timelog_assigned_col_start,
    timelog_summary_col_start,
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
    LIGHT_GREEN_FILL,
    SECTION_FILL,
    SECTION_FONT,
    TITLE_FONT,
    col_letter,
    sized_comment,
    style_cell,
)
from budget_planner.timelog_solver import Person, PersonWPWindow, WP, solve

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
    grade_rates: Mapping[str, float],  # noqa: ARG001 — reserved for future per-grade tinting
    config_layout: dict,
) -> dict:
    """Build the Timelog Plan sheet.

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
            ``param_rows``, ``wp_first_row``, ``person_rows``,
            ``roster_first_row``.

    Returns:
        Layout dict with keys:
          sheet, person_first_row, person_last_row, person_rows,
          proposed_col_start, proposed_col_end,
          assigned_col_start, assigned_col_end,
          summary_col_start, wp_summary_first_row, wp_summary_last_row.
    """
    project_months: int = project_params["months"]

    # ── Tab appearance ──
    ws.sheet_properties.tabColor = "FF8C00"  # orange — distinguishes from existing tabs

    # Resolve column geometry once.
    proposed_col_start = MONTH_COL_START
    proposed_col_end = MONTH_COL_START + project_months - 1
    assigned_col_start = timelog_assigned_col_start(project_months)
    assigned_col_end = timelog_assigned_col_end(project_months)
    summary_col_start = timelog_summary_col_start(project_months)
    summary_col_end = summary_col_start + 8 - 1  # 8 summary columns
    canvas_last_col = summary_col_end

    # ── Row 1: title banner (merged across canvas) ──
    ws.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=canvas_last_col
    )
    style_cell(
        ws, 1, 1,
        value="TIMELOG PLAN — prescriptive 1.0-FTE WP allocations",
        font=TITLE_FONT, fill=SECTION_FILL, alignment=CENTER, border=False,
    )
    # Fill remaining merged-row cells so the banner reads cleanly even when
    # openpyxl drops styling on merge target cells.
    for c in range(2, canvas_last_col + 1):
        style_cell(ws, 1, c, fill=SECTION_FILL, border=False)
    # Title font on the merged anchor must remain TITLE_FONT colour-overridden:
    # the SECTION_FILL is dark, so use white text.
    ws.cell(row=1, column=1).font = Font(
        name="Arial", bold=True, size=14, color="FFFFFF"
    )

    # ── Row 2: traffic-light cells ──
    # Layout: 4 (label, value) pairs across the canvas. Use the first 8 cols
    # of the canvas for readability.
    tl_row = 2
    # Cell positions for the four traffic-light VALUE cells. We compute these
    # before the WP-summary block exists, so the formulas reference cells we
    # haven't written yet — that's fine, Excel resolves at recalc.
    # We'll fill in the actual formulas after we know wp_summary rows.
    # For now, write labels and reserve value cells.
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
    # Pre-style empty value cells; formulas patched in after summary rows known.
    for c in tl_value_cols:
        style_cell(ws, tl_row, c, font=BLACK_FORMULA, alignment=CENTER, fmt="0")

    # Row 3: blank separator
    # (Nothing to write; openpyxl leaves blank.)

    # ── Rows 4-5: year/month header band ──
    # build_year_month_headers writes "WP ID"/"WP Name" into A/B of the month
    # row. Overwrite A with "Person" and clear B (it's our rate helper, not a
    # month). Year row is also overwritten in cols A/B so they don't carry the
    # LIGHT_BLUE_FILL "WP ID" suggestion.
    year_row, month_row, _next_after_headers = build_year_month_headers(
        ws, 4, project_params, MONTH_COL_START,
    )
    # Overwrite col A (year row + month row) — Person column header.
    style_cell(ws, year_row, 1, value="", font=HEADER_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    style_cell(ws, month_row, 1, value="Person", font=HEADER_FONT,
               fill=SECTION_FILL, alignment=CENTER).font = Font(
        name="Arial", bold=True, size=10, color="FFFFFF"
    )
    # Overwrite col B (rate helper).
    style_cell(ws, year_row, TIMELOG_RATE_HELPER_COL, value="",
               font=HEADER_FONT, fill=SECTION_FILL)
    style_cell(ws, month_row, TIMELOG_RATE_HELPER_COL, value="Rate",
               font=HEADER_FONT, fill=SECTION_FILL,
               alignment=CENTER).font = Font(
        name="Arial", bold=True, size=10, color="FFFFFF"
    )

    # Now build the Assigned-grid year/month headers as a mirror (so that
    # the summary block's references to month_headers work for both sides).
    # We rewrite the Assigned grid cells in the same year_row/month_row.
    _write_assigned_headers(
        ws, year_row, month_row, project_params,
        assigned_col_start, assigned_col_end,
    )

    # Hidden numeric month-index helper row (1..N) for SUMPRODUCT
    # out-of-window checks. Sits directly below the month header, hidden via
    # outline level so it remains inspectable.
    month_index_row = month_row + 1
    style_cell(ws, month_index_row, 1, value="(month idx)",
               font=Font(name="Arial", italic=True, color="808080", size=8),
               border=False)
    for mi in range(project_months):
        # Proposed grid copy
        c_p = MONTH_COL_START + mi
        style_cell(ws, month_index_row, c_p, value=mi + 1,
                   font=Font(name="Arial", color="808080", size=8),
                   alignment=CENTER, border=False)
        # Assigned grid copy
        c_a = assigned_col_start + mi
        style_cell(ws, month_index_row, c_a, value=mi + 1,
                   font=Font(name="Arial", color="808080", size=8),
                   alignment=CENTER, border=False)
    ws.row_dimensions[month_index_row].outline_level = 1
    ws.row_dimensions[month_index_row].hidden = True

    # ── Section header row (PROPOSED / ASSIGNED) ──
    section_row = month_index_row + 1
    # PROPOSED merge across rate-helper + months
    ws.merge_cells(
        start_row=section_row, start_column=1,
        end_row=section_row, end_column=proposed_col_end,
    )
    style_cell(
        ws, section_row, 1, value="PROPOSED (Python-computed, locked)",
        font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False,
    )
    # ASSIGNED merge across the assigned grid
    ws.merge_cells(
        start_row=section_row, start_column=assigned_col_start,
        end_row=section_row, end_column=assigned_col_end,
    )
    style_cell(
        ws, section_row, assigned_col_start,
        value="ASSIGNED (edit here — start as =Proposed)",
        font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False,
    )
    # Summary block header
    ws.merge_cells(
        start_row=section_row, start_column=summary_col_start,
        end_row=section_row, end_column=summary_col_end,
    )
    style_cell(
        ws, section_row, summary_col_start,
        value="PER-PERSON SUMMARY",
        font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False,
    )

    # Per-person summary column headers, on the row just below section_row
    summary_header_row = section_row + 1
    summary_headers = [
        "Assigned", "Pool", "Remaining", "Over-pool?",
        "Switches", "First log", "Last log", "Largest gap",
    ]
    for i, hdr in enumerate(summary_headers):
        style_cell(
            ws, summary_header_row, summary_col_start + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER,
        )

    # Person rows start after the summary header row.
    person_first_row = summary_header_row + 1

    # ── Build solver inputs and run solver ──
    # Sort persons (priority asc, name asc). The solver itself sorts by name
    # within each priority pass; we sort the input list so row order on the
    # sheet matches the solver's iteration order.
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

    # Index assignments by person → list of (wp_code, start, end) for the
    # in-window/out-of-window pre-fill.
    person_windows: dict[str, list[tuple[str, int, int]]] = {}
    for w in windows:
        person_windows.setdefault(w.person_name, []).append(
            (w.wp_code, w.start_month, w.end_month)
        )

    # ── Person rows ──
    cfg_person_rows: dict[str, int] = config_layout.get("person_rows", {})
    months_row_cfg = config_layout["param_rows"]["Project Months"]
    project_months_ref = cell_ref(SHEET_CONFIG, months_row_cfg, 2)

    person_rows: dict[str, int] = {}
    r = person_first_row
    for person in persons:
        name = person.name
        person_rows[name] = r

        # Col A: cross-ref to Config Team Roster row
        cfg_row = cfg_person_rows.get(name)
        if cfg_row is not None:
            style_cell(
                ws, r, 1, font=GREEN_XREF,
            ).value = f"={SHEET_CONFIG}!A{cfg_row}"
        else:
            # Fallback: literal name (should not happen if orchestrator
            # passes a config_layout aligned with `team`).
            style_cell(ws, r, 1, value=name, font=GREEN_XREF)

        # Col B: per-person rate helper — VLOOKUP into roster_lookup col 4.
        # Single-cell VLOOKUP per row (the safe pattern from openpyxl-patterns
        # rule on SUMPRODUCT + range lookups).
        style_cell(
            ws, r, TIMELOG_RATE_HELPER_COL, font=BLACK_FORMULA, fmt="€#,##0",
        ).value = f"=VLOOKUP(A{r},roster_lookup,4,FALSE)"

        # Proposed + Assigned cells per month
        person_assignment = assigned_map.get(name, {})
        windows_for_person = person_windows.get(name, [])
        for mi in range(project_months):
            month_num = mi + 1
            proposed_col = MONTH_COL_START + mi
            assigned_col = assigned_col_start + mi

            wp_code = person_assignment.get(month_num)
            in_any_window = any(
                start <= month_num <= end for (_wp, start, end)
                in windows_for_person
            )
            _write_proposed_cell(ws, r, proposed_col, wp_code, in_any_window)
            _write_assigned_cell(
                ws, r, assigned_col, proposed_col, wp_code, in_any_window,
            )

        # Per-person summary formulas
        _write_person_summary(
            ws, r,
            assigned_col_start=assigned_col_start,
            assigned_col_end=assigned_col_end,
            summary_col_start=summary_col_start,
            project_months_ref=project_months_ref,
            month_header_row=month_row,
            cfg_row=cfg_row,
        )

        r += 1
    person_last_row = r - 1

    # ── Per-WP summary block (below grids) ──
    wp_summary_section_row = person_last_row + 2
    ws.merge_cells(
        start_row=wp_summary_section_row, start_column=1,
        end_row=wp_summary_section_row, end_column=8,
    )
    style_cell(
        ws, wp_summary_section_row, 1,
        value="PER-WP SUMMARY", font=SECTION_FONT, fill=SECTION_FILL,
        alignment=CENTER, border=False,
    )
    wp_summary_header_row = wp_summary_section_row + 1
    wp_summary_headers = [
        "WP code", "Ceiling €", "Assigned €", "Overshoot €",
        "Undershoot €", "Assigned months", "Window months remaining",
        "Any out-of-window?",
    ]
    for i, hdr in enumerate(wp_summary_headers):
        style_cell(
            ws, wp_summary_header_row, 1 + i,
            value=hdr, font=HEADER_FONT, alignment=CENTER,
        )

    wp_summary_first_row = wp_summary_header_row + 1
    cfg_wp_first_row = config_layout["wp_first_row"]
    rate_helper_col_letter = get_column_letter(TIMELOG_RATE_HELPER_COL)

    for i, wp in enumerate(wps):
        row = wp_summary_first_row + i
        wp_code = wp.code
        # Col 1: WP code (literal label — black text).
        style_cell(ws, row, 1, value=wp_code, font=BLACK_FORMULA,
                   alignment=CENTER)
        # Col 2: Ceiling € — green xref to Config!E{wp_row}
        cfg_wp_row = cfg_wp_first_row + i
        ceiling_ref = cell_ref(SHEET_CONFIG, cfg_wp_row, 5)
        style_cell(ws, row, 2, font=GREEN_XREF, fmt="€#,##0",
                   ).value = f"={ceiling_ref}"

        # Col 3: Assigned € = SUMPRODUCT((assigned_grid="WPx") * person_rates)
        # Per-WP, summed across all (person, month) cells.
        assigned_range = _grid_range_ref(
            assigned_col_start, assigned_col_end,
            person_first_row, person_last_row,
        )
        # Per-person rate column repeated across the same row count.
        # SUMPRODUCT shape: assigned_range is (rows × months), rate column is
        # (rows × 1). To make SUMPRODUCT work we sum across months explicitly:
        # use SUMPRODUCT((assigned="WP")*rate_vector_broadcast) — but Excel
        # doesn't broadcast. The clean pattern: a per-row sum first.
        # We instead build a sum of per-row contributions:
        # = SUMPRODUCT( (assigned="WP")*1 , rate_col_repeated )
        # Approach: SUMPRODUCT with two same-shaped arrays. We multiply the
        # COUNTIF-style mask of assigned by the rate column expanded across
        # months via broadcasting — Excel/LO supports this when one operand
        # is column-shaped (rows × 1) and the other is rectangular
        # (rows × months): the column is reused across months automatically.
        rate_col_range = (
            f"{rate_helper_col_letter}{person_first_row}"
            f":{rate_helper_col_letter}{person_last_row}"
        )
        style_cell(
            ws, row, 3, font=BLACK_FORMULA, fmt="€#,##0",
        ).value = (
            f'=SUMPRODUCT(({assigned_range}="{wp_code}")*{rate_col_range})'
        )

        # Col 4: Overshoot € = MAX(0, Assigned - Ceiling)
        style_cell(
            ws, row, 4, font=BLACK_FORMULA, fmt="€#,##0",
        ).value = f"=MAX(0,C{row}-B{row})"

        # Col 5: Undershoot € = MAX(0, Ceiling - Assigned)
        style_cell(
            ws, row, 5, font=BLACK_FORMULA, fmt="€#,##0",
        ).value = f"=MAX(0,B{row}-C{row})"

        # Col 6: Assigned months = SUMPRODUCT(1*(assigned_grid="WPx"))
        style_cell(
            ws, row, 6, font=BLACK_FORMULA, fmt="0",
        ).value = f'=SUMPRODUCT(1*({assigned_range}="{wp_code}"))'

        # Col 7: Window months remaining = MAX(0, wp_end - current_month)
        # current_month named range is created by the Trends stub which the
        # orchestrator (T6) will run before this builder.
        style_cell(
            ws, row, 7, font=BLACK_FORMULA, fmt="0",
        ).value = f"=MAX(0,{wp.active_end}-current_month)"

        # Col 8: Any out-of-window? = SUMPRODUCT((assigned="WP") *
        #   ((month_idx<wp_start) + (month_idx>wp_end)))
        # month_idx_row spans the assigned grid columns.
        month_idx_range = (
            f"{get_column_letter(assigned_col_start)}{month_index_row}"
            f":{get_column_letter(assigned_col_end)}{month_index_row}"
        )
        style_cell(
            ws, row, 8, font=BLACK_FORMULA, fmt="0",
        ).value = (
            f'=SUMPRODUCT(({assigned_range}="{wp_code}")*'
            f"(({month_idx_range}<{wp.active_start})+"
            f"({month_idx_range}>{wp.active_end})))"
        )

    wp_summary_last_row = wp_summary_first_row + len(wps) - 1

    # ── Patch traffic-light formulas (now that wp summary rows exist) ──
    # 1) Any undershoot? — count WPs where Undershoot € > 0
    style_cell(
        ws, tl_row, tl_value_cols[0], font=BLACK_FORMULA, alignment=CENTER, fmt="0",
    ).value = (
        f"=COUNTIF(E{wp_summary_first_row}:E{wp_summary_last_row},\">0\")"
    )
    # 2) Any out-of-window cell? — sum across per-WP out-of-window counts
    style_cell(
        ws, tl_row, tl_value_cols[1], font=BLACK_FORMULA, alignment=CENTER, fmt="0",
    ).value = (
        f"=SUM(H{wp_summary_first_row}:H{wp_summary_last_row})"
    )
    # 3) Any over-pool person? — count persons where Assigned > Pool
    over_pool_col_letter = get_column_letter(summary_col_start + 3)
    style_cell(
        ws, tl_row, tl_value_cols[2], font=BLACK_FORMULA, alignment=CENTER, fmt="0",
    ).value = (
        f"=COUNTIF({over_pool_col_letter}{person_first_row}"
        f":{over_pool_col_letter}{person_last_row},TRUE)"
    )
    # 4) Total switches — sum of per-person switch counts
    switches_col_letter = get_column_letter(summary_col_start + 4)
    style_cell(
        ws, tl_row, tl_value_cols[3], font=BLACK_FORMULA, alignment=CENTER, fmt="0",
    ).value = (
        f"=SUM({switches_col_letter}{person_first_row}"
        f":{switches_col_letter}{person_last_row})"
    )

    # CF on traffic-light cells: red on > 0 for the first three (skip switches).
    for c in tl_value_cols[:3]:
        cell_addr = f"{get_column_letter(c)}{tl_row}"
        ws.conditional_formatting.add(
            cell_addr,
            FormulaRule(formula=[f"{cell_addr}>0"], fill=CF_RED_FILL),
        )

    # ── CF rules (canvas) — most-critical first ──
    # Per ``openpyxl-patterns.md``: first rule added gets highest priority.
    proposed_grid_range = _grid_range_ref(
        proposed_col_start, proposed_col_end,
        person_first_row, person_last_row,
    )
    assigned_grid_range = _grid_range_ref(
        assigned_col_start, assigned_col_end,
        person_first_row, person_last_row,
    )

    # Rule 1 (highest priority): Per-WP cell-level red marker for cells with
    # the WP code BUT in a column outside the WP's active window. We add one
    # rule per WP, applied to the union of Proposed + Assigned grids.
    # The CF formula uses the month-index helper row so the column-position
    # comparison is data-driven and survives row inserts/shifts.
    proposed_idx_range = (
        f"{get_column_letter(proposed_col_start)}{month_index_row}"
        f":{get_column_letter(proposed_col_end)}{month_index_row}"
    )
    assigned_idx_range = (
        f"{get_column_letter(assigned_col_start)}{month_index_row}"
        f":{get_column_letter(assigned_col_end)}{month_index_row}"
    )
    for wp in wps:
        # Per-grid (proposed and assigned each get their own rule with their
        # own month-index helper range).
        for grid_range, idx_range in (
            (proposed_grid_range, proposed_idx_range),
            (assigned_grid_range, assigned_idx_range),
        ):
            # The CF formula is evaluated per cell in the range. Use the
            # top-left cell of the range as the reference cell — Excel
            # auto-shifts relative refs.
            top_left = grid_range.split(":")[0]
            top_idx_letter = idx_range.split(":")[0]
            # Strip the row from top_left to get the column letter; the
            # idx-row stays absolute.
            tl_col_letter = "".join(
                ch for ch in top_left if ch.isalpha()
            )
            tl_row_num = "".join(
                ch for ch in top_left if ch.isdigit()
            )
            # Reference into the helper row at the same column (relative col,
            # absolute row).
            idx_ref = (
                f"{tl_col_letter}${month_index_row}"
            )
            cell_ref_inline = f"{tl_col_letter}{tl_row_num}"
            formula = (
                f'AND({cell_ref_inline}="{wp.code}",'
                f"OR({idx_ref}<{wp.active_start},{idx_ref}>{wp.active_end}))"
            )
            ws.conditional_formatting.add(
                grid_range,
                FormulaRule(formula=[formula], fill=CF_RED_FILL),
            )
            # idx_range/top_idx_letter computed for symmetry; not directly
            # used in the formula since we compose idx_ref by hand above.
            del top_idx_letter

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
    op_col_letter = get_column_letter(summary_col_start + 3)
    op_range = (
        f"{op_col_letter}{person_first_row}:{op_col_letter}{person_last_row}"
    )
    ws.conditional_formatting.add(
        op_range,
        FormulaRule(
            formula=[f"{op_col_letter}{person_first_row}=TRUE"],
            fill=CF_RED_FILL,
        ),
    )

    # Rule 4 (per-WP summary CF): Undershoot € > 0 → red.
    us_range = (
        f"E{wp_summary_first_row}:E{wp_summary_last_row}"
    )
    ws.conditional_formatting.add(
        us_range,
        FormulaRule(
            formula=[f"E{wp_summary_first_row}>0"],
            fill=CF_RED_FILL,
        ),
    )

    # Rule 5 (per-WP summary CF): Any out-of-window? > 0 → red.
    oow_range = (
        f"H{wp_summary_first_row}:H{wp_summary_last_row}"
    )
    ws.conditional_formatting.add(
        oow_range,
        FormulaRule(
            formula=[f"H{wp_summary_first_row}>0"],
            fill=CF_RED_FILL,
        ),
    )

    # ── Freeze panes & column widths ──
    # Freeze rows above the data (banner + traffic light + headers + section)
    # and the Person column. Freeze pane = top-left of the unfrozen area.
    ws.freeze_panes = ws.cell(row=person_first_row, column=MONTH_COL_START).coordinate
    ws.column_dimensions["A"].width = 25
    # Rate helper col B: narrow + collapsible (outline_level=1).
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["B"].outline_level = 1
    # Gap columns: narrow.
    gap1_col_letter = get_column_letter(proposed_col_end + 1)
    gap2_col_letter = get_column_letter(assigned_col_end + 1)
    for gap_letter in (gap1_col_letter, gap2_col_letter):
        ws.column_dimensions[gap_letter].width = 2
    # Summary cols a touch wider for headers.
    for i in range(8):
        col_l = get_column_letter(summary_col_start + i)
        ws.column_dimensions[col_l].width = 12

    return {
        "sheet": SHEET_TIMELOG_PLAN,
        "person_first_row": person_first_row,
        "person_last_row": person_last_row,
        "person_rows": person_rows,
        "proposed_col_start": proposed_col_start,
        "proposed_col_end": proposed_col_end,
        "assigned_col_start": assigned_col_start,
        "assigned_col_end": assigned_col_end,
        "summary_col_start": summary_col_start,
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


def _write_assigned_headers(
    ws: Worksheet,
    year_row: int,
    month_row: int,
    project_params: dict,
    assigned_col_start: int,
    assigned_col_end: int,
) -> None:
    """Write a mirror of the year/month header band over the Assigned grid.

    The build_year_month_headers helper writes only the Proposed-side band; we
    duplicate the month labels (and a single combined year cell) over the
    Assigned grid so first/last-log INDEX/MATCH formulas can reference the
    right-hand month_header_row directly.
    """
    from dateutil.relativedelta import relativedelta

    project_months = project_params["months"]
    project_start = project_params["start"]

    # Year row: single merged "ASSIGNED" label (cosmetic — Proposed's own
    # year row already shows the calendar years).
    ws.merge_cells(
        start_row=year_row, start_column=assigned_col_start,
        end_row=year_row, end_column=assigned_col_end,
    )
    style_cell(
        ws, year_row, assigned_col_start,
        value="(Assigned grid)", font=HEADER_FONT, alignment=CENTER,
    )

    # Month row mirror.
    for mi in range(project_months):
        c = assigned_col_start + mi
        cal_date = project_start + relativedelta(months=mi)
        mlabel = f"M{mi + 1}-{cal_date.strftime('%b%y')}"
        style_cell(
            ws, month_row, c, value=mlabel,
            font=Font(name="Arial", size=8), alignment=CENTER,
        )


def _grid_range_ref(
    col_start: int, col_end: int,
    row_start: int, row_end: int,
) -> str:
    """Return ``"C7:K20"`` for a rectangular range."""
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
    proposed_col: int,
    wp_code: str | None,
    in_any_window: bool,
) -> None:
    """Write a single Assigned-grid cell.

    Defaults to ``=<Proposed cell>``. Cells outside any person-WP window are
    pre-greyed (matching Proposed) and unlocked so the user can override.
    """
    proposed_letter = get_column_letter(proposed_col)
    formula = f"={proposed_letter}{row}"
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
    # BLUE_INPUT triggers auto-unlock in style_cell; cell already unlocked.
    # Keep explicit assignment to make the intent unmistakable.
    cell.protection = Protection(locked=False)


def _write_person_summary(
    ws: Worksheet,
    row: int,
    *,
    assigned_col_start: int,
    assigned_col_end: int,
    summary_col_start: int,
    project_months_ref: str,
    month_header_row: int,
    cfg_row: int | None,
) -> None:
    """Fill the 8 per-person summary cells on a single person row."""
    assigned_row_range = month_range_ref(
        row, assigned_col_start, assigned_col_end,
    )
    # For SUMPRODUCT switch counter, we compare two same-shape ranges offset
    # by one column.
    assigned_left_range = month_range_ref(
        row, assigned_col_start, assigned_col_end - 1,
    )
    assigned_right_range = month_range_ref(
        row, assigned_col_start + 1, assigned_col_end,
    )
    # Month-headers range on the assigned side (matches assigned_row_range
    # column-for-column so INDEX returns the correct month label).
    month_headers_range = month_range_ref(
        month_header_row, assigned_col_start, assigned_col_end,
    )

    # Col 1: Assigned count = COUNTIF(row, "WP*")
    style_cell(
        ws, row, summary_col_start, font=BLACK_FORMULA, fmt="0",
    ).value = f'=COUNTIF({assigned_row_range},"WP*")'

    # Col 2: Pool = ROUND(max_fte * project_months, 0)
    # max_fte lives at Config!B{cfg_row}; if cfg_row is missing, fall back to 0.
    if cfg_row is not None:
        max_fte_ref = cell_ref(SHEET_CONFIG, cfg_row, 2)
        style_cell(
            ws, row, summary_col_start + 1, font=GREEN_XREF, fmt="0",
        ).value = f"=ROUND({max_fte_ref}*{project_months_ref},0)"
    else:
        style_cell(
            ws, row, summary_col_start + 1, value=0,
            font=BLACK_FORMULA, fmt="0",
        )

    # Col 3: Remaining = Pool - Assigned
    style_cell(
        ws, row, summary_col_start + 2, font=BLACK_FORMULA, fmt="0",
    ).value = (
        f"={col_letter(summary_col_start + 1)}{row}"
        f"-{col_letter(summary_col_start)}{row}"
    )

    # Col 4: Over-pool? = Assigned > Pool
    style_cell(
        ws, row, summary_col_start + 3, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f"={col_letter(summary_col_start)}{row}"
        f">{col_letter(summary_col_start + 1)}{row}"
    )

    # Col 5: Switches = SUMPRODUCT(1*(left <> right))
    # Counts ALL adjacent diffs including empty<->WP per design Q15a.
    style_cell(
        ws, row, summary_col_start + 4, font=BLACK_FORMULA, fmt="0",
    ).value = (
        f"=SUMPRODUCT(1*({assigned_left_range}<>{assigned_right_range}))"
    )

    # Col 6: First log = IFERROR(INDEX(headers, MATCH(TRUE, row<>"", 0)), "")
    style_cell(
        ws, row, summary_col_start + 5, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f"=IFERROR(INDEX({month_headers_range},"
        f'MATCH(TRUE,{assigned_row_range}<>"",0)),"")'
    )

    # Col 7: Last log = IFERROR(LOOKUP(2, 1/(row<>""), headers), "")
    style_cell(
        ws, row, summary_col_start + 6, font=BLACK_FORMULA, alignment=CENTER,
    ).value = (
        f"=IFERROR(LOOKUP(2,1/({assigned_row_range}<>\"\"),"
        f'{month_headers_range}),"")'
    )

    # Col 8: Largest gap — placeholder, see module docstring.
    cell = style_cell(
        ws, row, summary_col_start + 7, value="",
        font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, alignment=CENTER,
    )
    cell.comment = sized_comment(
        "Largest gap calc deferred — see solver design Q follow-ups "
        "(needs hidden streak-length helper row + MAX). Placeholder cell.",
        author="Claude",
    )
