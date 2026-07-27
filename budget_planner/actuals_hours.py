"""Actuals-Hours sheet builder — per-person monthly hours timelog.

Mirrors the Capacity sheet geometry 1:1 (same WPs, same externals, same
months, same header rows) but holds **hours** rather than FTE. Hour input
cells are blue inputs the user fills with **posted-cumulative** totals
(hours posted through that month) per the 2026-05-09 design; inactive WP
months are greyed and locked, matching the Capacity behavior the auditors
expect. The shared ``person_section.filter_zero_fte_wps`` drops WPs whose
Config assignment has ``total_fte == 0`` from BOTH sheets so the row-order
mirror invariant holds — see ``docs/spec.md`` § Row-order contracts.

Surfaces: posted-cumulative hour inputs, per-row validation Sum, per-person
TOTAL row.

Role: monthly cells are an input surface for hour timelog (primary hours
surface); per-cell convention is "posted-cumulative-hours-through-month".
Trust boundary (``hours_actuals_through_month`` named range at
``'Actuals-Hours'!$D$1``) and two-cadence rules are canonical in
``docs/spec.md`` § Two-cadence actuals.
"""

from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import Outline

from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_ACTUALS_HOURS,
    SHEET_CONFIG,
    build_year_month_headers,
    month_range_ref,
    month_col_end as _month_col_end,
    val_columns,
    wp_active_months,
)
from budget_planner.person_section import (
    plan_person_layout,
    PersonGridContext,
    apply_person_row_outline,
    filter_zero_fte_wps,
)
from budget_planner.styles import (
    BLUE_INPUT,
    BLACK_FORMULA,
    GREEN_XREF,
    HEADER_FONT,
    SECTION_FILL,
    GREY_FILL,
    LIGHT_GREEN_FILL,
    style_cell,
    col_letter,
    sized_comment,
)

HOUR_FMT = "#,##0"


def _cum_hours_formula(row: int, month_col_start: int, months: int) -> str:
    """Latest filled posted-cumulative-hours value on this row at-or-before
    ``hours_actuals_through_month`` (`'Actuals-Hours'!$D$1`).

    Written on the Capacity-consumed ``cum_hours_col`` as a single helper
    so Capacity can reference one cell per row instead of rebuilding a
    cross-sheet INDEX scan.

    Under posted-cumulative cell semantics (design 2026-05-09 §1) each
    cell already holds the running total through its month, so we need
    the latest filled cell within the window — not a SUM, which would
    over-count cumulative snapshots. SUMPRODUCT-MAX-INDEX latest-filled
    scan mirrors the EUR-side ``posted_idx``/``posted_value`` pattern
    in ``actuals.py``.

    When no postings fall within the window the inner SUMPRODUCT yields
    0; ``INDEX(rng, 1, 0)`` then returns the entire row as an array
    (Excel) or 0 (LibreOffice). In Excel, scalar-context coercion of
    the array reference fires ``#VALUE!``, which ``IFERROR`` catches →
    returns 0. LibreOffice already returns 0, so IFERROR is harmless.
    Capacity's ``_wp_actual_formula`` divides by ``hours_per_fte_month``
    and surfaces "Actual PM = 0" cleanly under both engines.
    """
    start_cl = col_letter(month_col_start)
    end_cl = col_letter(month_col_start + months - 1)
    rng = f"${start_cl}${row}:${end_cl}${row}"
    return (
        f"=IFERROR(INDEX({rng},1,"
        f"SUMPRODUCT(MAX("
        f"(COLUMN({rng})-COLUMN(${start_cl}$1)+1<=hours_actuals_through_month)"
        f"*(COLUMN({rng})-COLUMN(${start_cl}$1)+1)"
        f"*ISNUMBER({rng})"
        f"))),0)"
    )


def _write_person_hours_grid(
    ws,
    r: int,
    member: dict,
    *,
    ctx: PersonGridContext,
    person_assignments: list[dict],
    cum_hours_col: int,
    preserved_for_person: dict[str, list[float | None]] | None,
) -> tuple[int, dict]:
    """Write one person's hour-grid section (header, WP rows, externals, TOTAL).

    Returns:
        (next_row, person_group_dict)
    """
    # Unpack loop-invariant context — keeps the cell-write body unchanged
    # from the pre-refactor signature.
    work_packages = ctx.work_packages
    external_commitments = ctx.external_commitments
    wp_active_map = ctx.wp_active_map
    project_params = ctx.project_params
    month_col_start = ctx.month_col_start
    month_col_end = ctx.month_col_end
    val_col_sum = ctx.val_col_sum
    config_first_wp_row = ctx.config_first_wp_row

    name = member["name"]

    # O(1) WP lookups — avoid O(N²) linear scans in the per-assignment loops
    # below (filter + Config-row offset).
    wp_by_id = {w["id"]: w for w in work_packages}
    wp_idx = {w["id"]: i for i, w in enumerate(work_packages)}

    # Pre-compute ordered WP ids and externals so plan_person_layout can
    # assign row numbers up front — mirror the Capacity behavior.
    ordered_wp_ids: list[str] = []
    for asgn in person_assignments:
        wp_id = asgn["wp"]
        if wp_id in wp_by_id and wp_id not in ordered_wp_ids:
            ordered_wp_ids.append(wp_id)
    ordered_wp_ids = filter_zero_fte_wps(person_assignments, ordered_wp_ids)
    person_externals = [e for e in external_commitments if e["person"] == name]
    ordered_external_names = [e["project"] for e in person_externals]

    layout_start = r
    if not ordered_wp_ids and not ordered_external_names:
        layout_start = r + 1
    layout = plan_person_layout(
        start_row=layout_start,
        wp_ids=ordered_wp_ids,
        external_names=ordered_external_names,
        include_remaining=False,
    )

    # ── Person header row ──
    header_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, font=HEADER_FONT, fill=SECTION_FILL)
    # Literal name (not =Config!A<roster_row>) so display under data_only
    # reads is stable. The GREEN_XREF convention still applies (cross-sheet
    # origin) but the header uses white-on-SECTION_FILL styling for the
    # person banner.
    ws.cell(row=r, column=1).value = name
    ws.cell(row=r, column=1).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    style_cell(ws, r, 2, fill=SECTION_FILL)
    r += 1

    data_row_start = r
    wp_rows: dict[str, int] = {}
    external_rows: dict[str, int] = {}

    # ── programme WP rows ──
    first_wp_row_for_person = True
    for wp_id in ordered_wp_ids:
        config_row = config_first_wp_row + wp_idx[wp_id]
        assert r == layout.wp_rows[wp_id]

        style_cell(ws, r, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
        style_cell(ws, r, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        wp_active = wp_active_map.get(wp_id, frozenset())
        preserved_months = preserved_for_person.get(wp_id) if preserved_for_person else None

        comment_added = False
        for mi in range(project_params["months"]):
            c = month_col_start + mi
            month_num = mi + 1

            if month_num in wp_active:
                # Start blank — user fills with posted-cumulative totals as
                # the accountant posts each month. Blanks INSIDE the active
                # window are allowed under posted-cumulative semantics.
                # Preserved values from a prior rebuild overwrite blanks
                # when provided.
                val: float | str = ""
                if preserved_months is not None and mi < len(preserved_months):
                    p = preserved_months[mi]
                    if p is not None:
                        val = p
                style_cell(ws, r, c, val, font=BLUE_INPUT, fmt=HOUR_FMT)
                if first_wp_row_for_person and not comment_added:
                    ws.cell(row=r, column=c).comment = sized_comment(
                        "Enter posted-cumulative hours through this month "
                        "for this work package (i.e. total hours posted "
                        "to-date, not the month's delta). Leave blank for "
                        "months where nothing has been posted yet."
                    )
                    comment_added = True
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL, fmt=HOUR_FMT)

        # Validation: Sum column only (no Budget/Variance — hours log is a
        # record, not a budget-vs-actual view; design §1).
        sum_range = month_range_ref(r, month_col_start, month_col_end)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fmt=HOUR_FMT).value = f"=SUM({sum_range})"
        style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA, fmt=HOUR_FMT).value = _cum_hours_formula(
            r,
            month_col_start,
            project_params["months"],
        )

        wp_rows[wp_id] = r
        first_wp_row_for_person = False
        r += 1

    # ── External project rows ──
    for entry in person_externals:
        assert r == layout.external_rows[entry["project"]]

        style_cell(ws, r, 1, entry["project"], font=BLUE_INPUT)
        style_cell(ws, r, 2, font=BLUE_INPUT)

        ext_start = entry["start"]
        ext_end = entry["end"]
        preserved_months = preserved_for_person.get(entry["project"]) if preserved_for_person else None

        for mi in range(project_params["months"]):
            c = month_col_start + mi
            month_num = mi + 1

            if ext_start <= month_num <= ext_end:
                val: float | str = ""
                if preserved_months is not None and mi < len(preserved_months):
                    p = preserved_months[mi]
                    if p is not None:
                        val = p
                style_cell(ws, r, c, val, font=BLUE_INPUT, fmt=HOUR_FMT)
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL, fmt=HOUR_FMT)

        sum_range = month_range_ref(r, month_col_start, month_col_end)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fmt=HOUR_FMT).value = f"=SUM({sum_range})"
        style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA, fmt=HOUR_FMT).value = _cum_hours_formula(
            r,
            month_col_start,
            project_params["months"],
        )

        external_rows[entry["project"]] = r
        r += 1

    data_row_end = r - 1

    # Person with no data rows: keep placeholder row so TOTAL sits where
    # plan_person_layout said it would.
    if data_row_end < data_row_start:
        data_row_end = data_row_start
        r += 1

    # ── TOTAL row (SUM of all data rows, hours per month) ──
    assert r == layout.total_row
    total_row = layout.total_row
    style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    ws.cell(row=r, column=2).comment = sized_comment(
        "Per-month SUM of all hours-cells across programme WP rows and external "
        "project rows for this person. Under posted-cumulative semantics "
        "this column-wise SUM is a coarse aggregate — the per-row cum "
        "helper col on the right is what Capacity consumes."
    )
    for mi in range(project_params["months"]):
        c = month_col_start + mi
        cl = col_letter(c)
        style_cell(
            ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=HOUR_FMT
        ).value = f"=SUM({cl}{data_row_start}:{cl}{data_row_end})"
    scl = col_letter(val_col_sum)
    style_cell(
        ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=HOUR_FMT
    ).value = f"=SUM({scl}{data_row_start}:{scl}{data_row_end})"
    style_cell(
        ws, r, cum_hours_col, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=HOUR_FMT
    ).value = _cum_hours_formula(
        r,
        month_col_start,
        project_params["months"],
    )
    r += 1

    assert r == layout.next_row
    apply_person_row_outline(ws, layout)
    return layout.next_row, {
        "header_row": header_row,
        "wp_rows": wp_rows,
        "external_rows": external_rows,
        "total_row": total_row,
    }


def build_actuals_hours_sheet(
    ws,
    *,
    project_params,
    work_packages,
    team,
    external_commitments,
    assignments,
    config_first_wp_row,
    wb,
    preserved: dict[str, dict[str, list[float | None]]] | None = None,
    hours_through_override: int | None = None,
) -> dict:
    """Build the Actuals-Hours sheet: per-person monthly hour timelog grid.

    Args:
        ws: The Actuals-Hours worksheet to populate.
        project_params: dict with 'title', 'start', 'months', 'rate'.
        work_packages: list of WP dicts with single-range schedule fields
            ('start', 'end', 'type') for active-month greying.
        team: list of team member dicts.
        external_commitments: list of external commitment dicts.
        assignments: list of person-WP assignment dicts (drives which WP
            rows appear per person — same intersection as Capacity).
        config_first_wp_row: First WP row in Config WP table.
        preserved: Optional rebuild-preservation payload
            ``{person_name: {wp_id_or_project: [hours_per_month, ...]}}``.
            When provided, blue-input hour cells are pre-filled from this
            map (None entries are skipped).

    Returns:
        actuals_hours_layout dict: ``{year_header_row, month_header_row,
        cum_hours_col, person_groups: {name: {header_row, wp_rows,
        external_rows, total_row}}}``.
    """
    if assignments is None:
        assignments = []

    months = project_params["months"]
    month_col_start = MONTH_COL_START
    month_col_end = _month_col_end(months)
    _vcols = val_columns(months)
    val_col_sum = _vcols["sum"]
    # Cum-hours helper column: blank spacer at val_col_sum+1, helper at +2.
    # Capacity's to-date Actual PM references this cell divided by
    # hours_per_fte_month (see design doc 2026-04-14-t5).
    cum_hours_col = val_col_sum + 2

    ws.sheet_properties.tabColor = "C00000"  # distinct red — hours timelog
    # Place the row-outline button next to each person header
    # (summaryBelow=False) so collapsing a person hides their rows beneath
    # the header.
    outline_pr = ws.sheet_properties.outlinePr
    if outline_pr is None:
        ws.sheet_properties.outlinePr = Outline(
            summaryRight=True,
            summaryBelow=False,
        )
    else:
        outline_pr.summaryRight = True
        outline_pr.summaryBelow = False

    # ── Header rows (Sum only — no Budget/Variance) ──
    # Row 1 holds the hours through-month control strip (label / override /
    # auto-derived / effective). Year + month header rows follow on rows
    # 2 and 3.
    r = 2
    year_header_row, month_header_row, r = build_year_month_headers(
        ws,
        r,
        project_params,
        month_col_start,
        val_headers=[(val_col_sum, "Sum")],
    )

    ws.cell(row=year_header_row, column=val_col_sum).comment = sized_comment(
        "Row-wide SUM of the per-month posted-cumulative hour cells for "
        "this row. Under posted-cumulative semantics this SUM is a coarse "
        "aggregate; the cum-hours helper col is what Capacity consumes."
    )

    # Cum-hours header: year row blank but fill-matched; month row carries
    # the label. Green so the reader sees this col as a helper aggregate.
    style_cell(ws, year_header_row, cum_hours_col, fill=LIGHT_GREEN_FILL)
    style_cell(
        ws, month_header_row, cum_hours_col, value="Cum hours (to date)", font=HEADER_FONT, fill=LIGHT_GREEN_FILL
    )
    ws.cell(row=month_header_row, column=cum_hours_col).comment = sized_comment(
        "Cumulative hours on this row through the `hours_actuals_through_month` "
        "named range (`'Actuals-Hours'!$D$1`). Capacity's to-date Actual PM reads this cell "
        "divided by hours_per_fte_month."
    )

    first_person_header_row = r

    # ── Per-person assignment lookup (dedup by person+wp) ──
    _dedup: dict[tuple[str, str], dict] = {}
    for entry in assignments:
        _dedup[(entry["person"], entry["wp"])] = entry
    per_person_assignments: dict[str, list[dict]] = {}
    for entry in _dedup.values():
        per_person_assignments.setdefault(entry["person"], []).append(entry)

    wp_active_map: dict[str, frozenset[int]] = {
        wp["id"]: wp_active_months(wp["id"], work_packages) for wp in work_packages
    }

    person_groups: dict[str, dict] = {}

    ctx = PersonGridContext(
        work_packages=work_packages,
        external_commitments=external_commitments,
        wp_active_map=wp_active_map,
        project_params=project_params,
        month_col_start=month_col_start,
        month_col_end=month_col_end,
        val_col_sum=val_col_sum,
        config_first_wp_row=config_first_wp_row,
    )
    for member in team:
        name = member["name"]
        preserved_for_person = preserved.get(name) if preserved else None
        r, group = _write_person_hours_grid(
            ws,
            r,
            member,
            ctx=ctx,
            person_assignments=per_person_assignments.get(name, []),
            cum_hours_col=cum_hours_col,
            preserved_for_person=preserved_for_person,
        )
        person_groups[name] = group

    # ── Row 1: hours through-month control strip ──
    # A: label, B: override (blue; blank by default), C: auto-derived scan
    # of every person's programme-WP rows (excludes externals + per-person TOTAL
    # rows), D: effective = override-if-numeric-else-auto.
    # ``hours_actuals_through_month`` resolves to the effective cell on
    # this sheet rather than Config!B8.
    ms_strip = col_letter(month_col_start)
    me_strip = col_letter(month_col_end)
    grant_wp_ranges: list[tuple[int, int]] = []
    for grp in person_groups.values():
        grp_wp_rows = grp.get("wp_rows") or {}
        if not grp_wp_rows:
            continue
        sorted_rows = sorted(grp_wp_rows.values())
        grant_wp_ranges.append((sorted_rows[0], sorted_rows[-1]))
    if grant_wp_ranges:
        scan_terms = []
        for r0, r1 in grant_wp_ranges:
            rng = f"${ms_strip}${r0}:${me_strip}${r1}"
            scan_terms.append(f"SUMPRODUCT(MAX((COLUMN({rng})-COLUMN(${ms_strip}$1)+1)*ISNUMBER({rng})))")
        auto_formula = f"=MAX({','.join(scan_terms)})"
    else:
        auto_formula = "=0"

    style_cell(ws, 1, 1, value="Hours Actuals Through Month:", font=HEADER_FONT, border=False)
    ws.cell(row=1, column=1).comment = sized_comment(
        "Trust boundary for hours-side cumulative figures. The effective "
        "value in column D is what Capacity and Trends consume via the "
        "named range `hours_actuals_through_month`. Override (B) beats "
        "auto-derived (C) only when it holds a number."
    )
    style_cell(ws, 1, 2, value=hours_through_override, font=BLUE_INPUT, fmt="0", border=False)
    ws.cell(row=1, column=2).comment = sized_comment(
        "Optional override — enter an integer month index (1-based) to "
        "force a tighter hours trust cutoff than the auto-derived value. "
        "Leave blank to defer to the auto scan in column C. Blue = user "
        "input."
    )
    style_cell(ws, 1, 3, font=BLACK_FORMULA, fmt="0", border=False).value = auto_formula
    ws.cell(row=1, column=3).comment = sized_comment(
        "Auto-derived: largest 1-based month index with any numeric hour "
        "value across every person's programme-WP rows (externals and per-person "
        "TOTAL rows are excluded). Returns 0 when no hours have been "
        "posted yet. Do not edit."
    )
    style_cell(ws, 1, 4, font=BLACK_FORMULA, fmt="0", border=False).value = "=IF(ISNUMBER(B1),B1,C1)"
    ws.cell(row=1, column=4).comment = sized_comment(
        "Effective hours trust cutoff consumed by Capacity and Trends via "
        "the named range `hours_actuals_through_month`. Override in B wins "
        "when numeric; otherwise falls through to the auto scan in C. Do "
        "not edit."
    )
    wb.defined_names.add(
        DefinedName(
            "hours_actuals_through_month",
            attr_text=f"'{SHEET_ACTUALS_HOURS}'!$D$1",
        )
    )

    ws.freeze_panes = f"C{first_person_header_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "cum_hours_col": cum_hours_col,
        "person_groups": person_groups,
    }
