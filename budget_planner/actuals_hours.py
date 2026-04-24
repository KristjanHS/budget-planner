"""Actuals-Hours sheet builder — per-person monthly hours timelog.

Mirrors the Capacity sheet geometry 1:1 (same WPs, same externals, same
months, same header rows) but holds **hours** rather than FTE. Hour input
cells are blank blue inputs the user fills in at month-end; inactive WP
months are greyed and locked, matching the Capacity behavior the auditors
expect.

Surfaces: hour inputs, per-row validation Sum, per-person TOTAL and EU
TOTAL hour rows, EU annual ceiling conditional formatting per calendar
year, per-row forecast-EUR band (``-i`` — without-indirect; collapsible
column outline), per-row ``Forecast EUR +i`` rollup column (with-indirect;
default-visible, placed outside the column-outline band).
"""

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import Outline

from budget_planner.layout import (
    MONTH_COL_START, SHEET_ACTUALS_HOURS, SHEET_CONFIG,
    build_year_month_headers, month_range_ref,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
    _calendar_year_groups,
)
from budget_planner.person_section import (
    plan_person_layout, PersonGridContext, apply_person_row_outline,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    WITH_INDIRECT_HEADER_FILL, WITHOUT_INDIRECT_HEADER_FILL,
    CF_AMBER_FILL, CF_RED_FILL, CF_GREY_FILL, style_cell, col_letter,
    sized_comment,
)

HOUR_FMT = "#,##0"
EUR_FMT = "€#,##0"
FORECAST_BAND_LABEL = "FORECAST EUR -i (from hours)"
FORECAST_PLUS_I_HEADER = "Forecast EUR +i"


def _mirror_year_month_headers_to_band(
    ws, year_header_row: int, month_header_row: int,
    month_col_start: int, months: int, project_start,
    forecast_col_start: int,
) -> None:
    """Copy year + month header labels onto the forecast-EUR band.

    Mirrors the hours-side calendar-year labels (merged across each year's
    column span) and the per-month ``M<n>-<MmmYY>`` labels. Fill uses the
    without-indirect (``-i``) orange per the indirect-view convention.
    """
    from openpyxl.styles import Font as _Font
    small = _Font(name="Arial", size=8)
    # Year row — merge across each calendar-year span.
    for group_start, span in _calendar_year_groups(project_start, months):
        yr_start_col = forecast_col_start + group_start
        yr_end_col = yr_start_col + span - 1
        # Pull label from the hours-side cell (already written).
        src_year_col = month_col_start + group_start
        label = ws.cell(row=year_header_row, column=src_year_col).value
        if span > 1:
            ws.merge_cells(start_row=year_header_row, start_column=yr_start_col,
                           end_row=year_header_row, end_column=yr_end_col)
        style_cell(ws, year_header_row, yr_start_col, value=label,
                   font=HEADER_FONT, fill=WITHOUT_INDIRECT_HEADER_FILL)
        for c2 in range(yr_start_col + 1, yr_end_col + 1):
            style_cell(ws, year_header_row, c2,
                       fill=WITHOUT_INDIRECT_HEADER_FILL)
    # Month row — one label per column.
    for mi in range(months):
        src_month = ws.cell(row=month_header_row, column=month_col_start + mi)
        style_cell(
            ws, month_header_row, forecast_col_start + mi,
            value=src_month.value,
            font=small, fill=WITHOUT_INDIRECT_HEADER_FILL,
        )


def _cum_hours_formula(row: int, month_col_start: int, months: int) -> str:
    """Windowed cumulative SUM of this row's hour cells up to
    ``hours_actuals_through_month`` (`'Actuals-Hours'!$D$1`).

    Written on the Capacity-consumed ``cum_hours_col`` as a single helper
    so Capacity can reference one cell per row instead of rebuilding a
    cross-sheet SUM+INDEX (see ``docs/plans/archived/2026-04-14-t5-simplification-design.md``).
    """
    start_cl = col_letter(month_col_start)
    end_cl = col_letter(month_col_start + months - 1)
    rng = f"${start_cl}${row}:${end_cl}${row}"
    return (
        f"=IF(hours_actuals_through_month<=0,0,"
        f"SUM(${start_cl}${row}:"
        f"INDEX({rng},1,hours_actuals_through_month)))"
    )


def _write_person_hours_grid(
    ws, r: int, member: dict, *,
    ctx: PersonGridContext,
    person_assignments: list[dict],
    forecast_col_start: int,
    forecast_col_end: int,
    forecast_plus_i_col: int,
    cum_hours_col: int,
    preserved_for_person: dict[str, list[float | None]] | None,
) -> tuple[int, dict]:
    """Write one person's hour-grid section (header, WP rows, externals, totals).

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
        include_eu_total=True,
        include_hours_todate_summary=True,
    )

    # ── Person header row ──
    header_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, font=HEADER_FONT, fill=SECTION_FILL)
    # Literal name (not =Config!A<roster_row>) so the forecast-EUR VLOOKUP
    # against $A$<header_row> works under data_only reads. The GREEN_XREF
    # convention still applies (cross-sheet origin) but the header uses
    # white-on-SECTION_FILL styling for the person banner.
    ws.cell(row=r, column=1).value = name
    ws.cell(row=r, column=1).font = Font(name="Arial", bold=True, size=10,
                                          color="FFFFFF")
    style_cell(ws, r, 2, fill=SECTION_FILL)
    # Also style forecast-band cells on header row so banding aligns.
    for mi in range(project_params["months"]):
        style_cell(ws, r, forecast_col_start + mi, fill=SECTION_FILL)
    # +i rollup col on header row: same banner fill.
    style_cell(ws, r, forecast_plus_i_col, fill=SECTION_FILL)
    r += 1

    # Person-name header cell used as the lookup key for rate VLOOKUP.
    person_header_cell = f"$A${header_row}"

    data_row_start = r
    wp_rows: dict[str, int] = {}
    external_rows: dict[str, int] = {}

    # ── programme WP rows ──
    first_wp_row_for_person = True
    for wp_id in ordered_wp_ids:
        config_row = config_first_wp_row + wp_idx[wp_id]
        assert r == layout.wp_rows[wp_id]

        style_cell(ws, r, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(ws, r, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        wp_active = wp_active_map.get(wp_id, frozenset())
        preserved_months = (
            preserved_for_person.get(wp_id) if preserved_for_person else None
        )

        comment_added = False
        for mi in range(project_params["months"]):
            c = month_col_start + mi
            fc = forecast_col_start + mi
            month_num = mi + 1

            if month_num in wp_active:
                # Start blank — user fills at month-end. Preserved values
                # from a prior rebuild overwrite blanks when provided.
                val: float | str = ""
                if preserved_months is not None and mi < len(preserved_months):
                    p = preserved_months[mi]
                    if p is not None:
                        val = p
                style_cell(ws, r, c, val, font=BLUE_INPUT, fmt=HOUR_FMT)
                if first_wp_row_for_person and not comment_added:
                    ws.cell(row=r, column=c).comment = sized_comment("Enter hours worked per month for this work package. "
                        "Leave blank until the month closes.")
                    comment_added = True
                # Forecast-EUR leaf — derived from the sibling hour cell.
                # ROUND to whole EUR at the leaf so Plan-Forecast SUMs match
                # auditor-computed totals (hours/143 is fractional; without
                # ROUND, sub-cent drift aggregates to several EUR over a
                # full project).
                hour_ref = f"{col_letter(c)}{r}"
                forecast_formula = (
                    f'=IF({hour_ref}="","",ROUND({hour_ref}/hours_per_fte_month'
                    f"*VLOOKUP({person_header_cell},roster_lookup,4,FALSE),0))"
                )
                style_cell(ws, r, fc, font=BLACK_FORMULA,
                           fmt=EUR_FMT).value = forecast_formula
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt=HOUR_FMT)
                # Inactive months — forecast-EUR cell is empty + grey too.
                style_cell(ws, r, fc, "", font=BLACK_FORMULA, fill=GREY_FILL,
                           fmt=EUR_FMT)

        # Validation: Sum column only (no Budget/Variance — hours log is a
        # record, not a budget-vs-actual view; design §1).
        sum_range = month_range_ref(r, month_col_start, month_col_end)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fmt=HOUR_FMT).value = f"=SUM({sum_range})"
        style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA,
                   fmt=HOUR_FMT).value = _cum_hours_formula(
            r, month_col_start, project_params["months"],
        )
        # Forecast EUR +i rollup: row-sum of the -i forecast band × 1.25.
        # ROUND at the leaf per openpyxl.md EUR rule.
        fc_start_l = col_letter(forecast_col_start)
        fc_end_l = col_letter(forecast_col_end)
        style_cell(ws, r, forecast_plus_i_col, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL, fmt=EUR_FMT).value = (
            f"=ROUND(SUM({fc_start_l}{r}:{fc_end_l}{r})*1.25,0)"
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
        preserved_months = (
            preserved_for_person.get(entry["project"])
            if preserved_for_person else None
        )

        for mi in range(project_params["months"]):
            c = month_col_start + mi
            fc = forecast_col_start + mi
            month_num = mi + 1

            if ext_start <= month_num <= ext_end:
                val: float | str = ""
                if preserved_months is not None and mi < len(preserved_months):
                    p = preserved_months[mi]
                    if p is not None:
                        val = p
                style_cell(ws, r, c, val, font=BLUE_INPUT, fmt=HOUR_FMT)
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt=HOUR_FMT)
            # External rows never contribute to programme EUR forecast — they are
            # non-programme projects (even if EU-flagged for the annual ceiling).
            # Leave the forecast-band cell empty + grey so the band stays
            # visually aligned but doesn't imply programme budget burn.
            style_cell(ws, r, fc, "", font=BLACK_FORMULA, fill=GREY_FILL,
                       fmt=EUR_FMT)

        sum_range = month_range_ref(r, month_col_start, month_col_end)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fmt=HOUR_FMT).value = f"=SUM({sum_range})"
        style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA,
                   fmt=HOUR_FMT).value = _cum_hours_formula(
            r, month_col_start, project_params["months"],
        )
        # Forecast EUR +i: external rows never contribute to programme EUR —
        # empty + grey, matching the -i band treatment.
        style_cell(ws, r, forecast_plus_i_col, "", font=BLACK_FORMULA,
                   fill=GREY_FILL, fmt=EUR_FMT)

        external_rows[entry["project"]] = r
        r += 1

    data_row_end = r - 1

    # Person with no data rows: keep placeholder row so TOTAL sits where
    # plan_person_layout said it would.
    if data_row_end < data_row_start:
        data_row_end = data_row_start
        r += 1

    # ── Hours-to-date summary rows (above TOTAL) ──
    # Two cumulative-through-``hours_actuals_through_month`` aggregates per
    # person: total hours logged and the person-months equivalent. Summaries
    # sit ABOVE TOTAL; TOTAL's per-month SUM still spans only the data rows
    # (``data_row_start..data_row_end``) so it doesn't double-count these.
    cum_hours_cl = col_letter(cum_hours_col)
    assert layout.cum_hours_todate_row is not None
    assert layout.actual_pm_todate_row is not None
    assert r == layout.cum_hours_todate_row
    cum_hours_todate_row = layout.cum_hours_todate_row
    style_cell(ws, r, 2, "Cumulative hours to date",
               font=HEADER_FONT)
    style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA,
               fmt=HOUR_FMT).value = (
        f"=SUM({cum_hours_cl}{data_row_start}:"
        f"{cum_hours_cl}{data_row_end})"
    )
    r += 1

    assert r == layout.actual_pm_todate_row
    actual_pm_todate_row = layout.actual_pm_todate_row
    style_cell(ws, r, 2, "Actual PM to date", font=HEADER_FONT)
    style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA,
               fmt="0.00").value = (
        f"={cum_hours_cl}{cum_hours_todate_row}/hours_per_fte_month"
    )
    r += 1

    # ── TOTAL row (SUM of all data rows, hours per month) ──
    assert r == layout.total_row
    total_row = layout.total_row
    style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    ws.cell(row=r, column=2).comment = sized_comment("Per-month SUM of all hours logged across programme WP rows and external "
        "project rows for this person.")
    for mi in range(project_params["months"]):
        c = month_col_start + mi
        cl = col_letter(c)
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt=HOUR_FMT).value = (
            f"=SUM({cl}{data_row_start}:{cl}{data_row_end})"
        )
        # Forecast-band TOTAL: sum WP-row forecast cells (externals are
        # empty in this band — they don't consume programme budget).
        fc = forecast_col_start + mi
        fcl = col_letter(fc)
        if wp_rows:
            wp_refs = ",".join(f"{fcl}{row}" for row in wp_rows.values())
            forecast_total = f"=SUM({wp_refs})"
        else:
            forecast_total = "=0"
        style_cell(ws, r, fc, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt=EUR_FMT).value = forecast_total
    scl = col_letter(val_col_sum)
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt=HOUR_FMT).value = (
        f"=SUM({scl}{data_row_start}:{scl}{data_row_end})"
    )
    style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA,
               fill=LIGHT_GREEN_FILL,
               fmt=HOUR_FMT).value = _cum_hours_formula(
        r, month_col_start, project_params["months"],
    )
    # Forecast EUR +i on TOTAL: ROUND(SUM(-i band on TOTAL)*1.25, 0).
    fc_start_l = col_letter(forecast_col_start)
    fc_end_l = col_letter(forecast_col_end)
    style_cell(ws, r, forecast_plus_i_col, font=BLACK_FORMULA,
               fill=LIGHT_GREEN_FILL, fmt=EUR_FMT).value = (
        f"=ROUND(SUM({fc_start_l}{r}:{fc_end_l}{r})*1.25,0)"
    )
    r += 1

    # ── EU TOTAL row (SUM of programme WP rows + EU-flagged externals) ──
    assert layout.eu_total_row is not None
    assert r == layout.eu_total_row
    eu_total_row = layout.eu_total_row
    style_cell(ws, r, 1, fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "EU TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)
    ws.cell(row=r, column=2).comment = sized_comment("Per-month SUM of programme WP rows + EU-flagged external project rows "
        "(hours). Compared per calendar year against "
        "eu_annual_ceiling_pm × hours_per_fte_month.")

    eu_row_refs = list(wp_rows.values())
    for proj, ext_row in external_rows.items():
        entry = next((e for e in person_externals
                      if e["project"] == proj), None)
        if entry and entry.get("is_eu", False):
            eu_row_refs.append(ext_row)

    for mi in range(project_params["months"]):
        c = month_col_start + mi
        cl = col_letter(c)
        if eu_row_refs:
            cell_refs = ",".join(f"{cl}{row}" for row in eu_row_refs)
            formula = f"=SUM({cell_refs})"
        else:
            formula = "=0"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
                   fmt=HOUR_FMT).value = formula
        # Forecast-band EU TOTAL: only programme WP rows contribute EUR (EU-flagged
        # externals are non-programme projects — empty in the forecast band).
        fc = forecast_col_start + mi
        fcl = col_letter(fc)
        if wp_rows:
            wp_refs = ",".join(f"{fcl}{row}" for row in wp_rows.values())
            eu_forecast = f"=SUM({wp_refs})"
        else:
            eu_forecast = "=0"
        style_cell(ws, r, fc, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
                   fmt=EUR_FMT).value = eu_forecast

    # Sum column for EU TOTAL
    if eu_row_refs:
        sum_refs = ",".join(f"{col_letter(val_col_sum)}{row}"
                            for row in eu_row_refs)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fill=LIGHT_BLUE_FILL,
                   fmt=HOUR_FMT).value = f"=SUM({sum_refs})"
    else:
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fill=LIGHT_BLUE_FILL, fmt=HOUR_FMT).value = "=0"
    # Cum hours on EU TOTAL: windowed SUM of this row's (per-month EU SUM)
    # cells — mirrors the TOTAL-row pattern.
    style_cell(ws, r, cum_hours_col, font=BLACK_FORMULA,
               fill=LIGHT_BLUE_FILL,
               fmt=HOUR_FMT).value = _cum_hours_formula(
        r, month_col_start, project_params["months"],
    )
    # Forecast EUR +i on EU TOTAL: ROUND(SUM(-i band on EU TOTAL)*1.25, 0).
    style_cell(ws, r, forecast_plus_i_col, font=BLACK_FORMULA,
               fill=LIGHT_BLUE_FILL, fmt=EUR_FMT).value = (
        f"=ROUND(SUM({fc_start_l}{r}:{fc_end_l}{r})*1.25,0)"
    )
    r += 1

    # ── EU ceiling CF per calendar year on EU TOTAL row ──
    # Red: annual_hours > eu_annual_ceiling_pm * hours_per_fte_month
    # Amber: within one hours_per_fte_month unit of the ceiling
    for group_start, span in _calendar_year_groups(
        project_params["start"], project_params["months"]
    ):
        yr_first_col = month_col_start + group_start
        yr_last_col = yr_first_col + span - 1
        yr_range = (f"{col_letter(yr_first_col)}{eu_total_row}:"
                    f"{col_letter(yr_last_col)}{eu_total_row}")
        sum_expr = (f"SUM(${col_letter(yr_first_col)}${eu_total_row}"
                    f":${col_letter(yr_last_col)}${eu_total_row})")
        # Red first — highest priority in openpyxl CF ordering.
        ws.conditional_formatting.add(yr_range, FormulaRule(
            formula=[f"{sum_expr}>eu_annual_ceiling_pm*hours_per_fte_month"],
            fill=CF_RED_FILL,
        ))
        ws.conditional_formatting.add(yr_range, FormulaRule(
            formula=[
                f"{sum_expr}>eu_annual_ceiling_pm*hours_per_fte_month"
                f"-hours_per_fte_month"
            ],
            fill=CF_AMBER_FILL,
        ))

    assert r == layout.next_row
    apply_person_row_outline(ws, layout)
    return layout.next_row, {
        "header_row": header_row,
        "wp_rows": wp_rows,
        "external_rows": external_rows,
        "cum_hours_todate_row": cum_hours_todate_row,
        "actual_pm_todate_row": actual_pm_todate_row,
        "total_row": total_row,
        "eu_total_row": eu_total_row,
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
        person_groups: {name: {header_row, wp_rows, external_rows,
        total_row, eu_total_row}}}``.
    """
    if assignments is None:
        assignments = []

    months = project_params["months"]
    month_col_start = MONTH_COL_START
    month_col_end = _month_col_end(months)
    _vcols = val_columns(months)
    val_col_sum = _vcols["sum"]
    # Forecast-EUR band: one gutter col after Sum, then ``months`` cols.
    forecast_col_start = val_col_sum + 2
    forecast_col_end = forecast_col_start + months - 1
    # Forecast EUR +i rollup column (with-indirect, default-visible). Placed
    # adjacent to the -i band but OUTSIDE its column outline group so it
    # stays visible when the band collapses.
    forecast_plus_i_col = forecast_col_end + 1
    # Cum-hours helper column (one gutter after the +i rollup). Sits outside
    # the forecast outline group so it remains visible when the forecast band
    # collapses; Capacity's to-date Actual PM references this cell divided
    # by hours_per_fte_month (see design doc 2026-04-14-t5).
    cum_hours_col = forecast_plus_i_col + 2

    ws.sheet_properties.tabColor = "C00000"  # distinct red — hours timelog
    # Place the column-outline collapse button to the right of grouped columns;
    # place the row-outline button next to each person header (summaryBelow=False)
    # so collapsing a person hides their rows beneath the header.
    outline_pr = ws.sheet_properties.outlinePr
    if outline_pr is None:
        ws.sheet_properties.outlinePr = Outline(
            summaryRight=True, summaryBelow=False,
        )
    else:
        outline_pr.summaryRight = True
        outline_pr.summaryBelow = False

    # ── Forecast band top label row ──
    # Row 1 holds the hours through-month control strip (label / override /
    # auto-derived / effective) — see docs/plans/2026-04-18-config-through-
    # month-cleanup-design.md. The merged "FORECAST EUR (from hours)"
    # header shifts down to row 2, keeping its position above the year
    # labels on row 3.
    band_label_row = 2
    ws.merge_cells(start_row=band_label_row, start_column=forecast_col_start,
                   end_row=band_label_row, end_column=forecast_col_end)
    style_cell(ws, band_label_row, forecast_col_start,
               value=FORECAST_BAND_LABEL,
               font=HEADER_FONT, fill=WITHOUT_INDIRECT_HEADER_FILL,
               alignment=None)
    # Forecast EUR +i header cell sits on the band-label row, outside the
    # merged -i label. Green (with-indirect) fill per the convention.
    style_cell(ws, band_label_row, forecast_plus_i_col,
               value=FORECAST_PLUS_I_HEADER,
               font=HEADER_FONT, fill=WITH_INDIRECT_HEADER_FILL,
               alignment=None)

    # ── Header rows (Sum only — no Budget/Variance) ──
    r = band_label_row + 1
    year_header_row, month_header_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[(val_col_sum, "Sum")],
    )

    # Mirror year + month headers across the forecast band so auditors see
    # the same month labels there.
    _mirror_year_month_headers_to_band(
        ws, year_header_row, month_header_row,
        month_col_start, months, project_params["start"],
        forecast_col_start,
    )

    # Mark each forecast-band month column as outline group level 1 so the
    # user can collapse the band via Excel's outline button. Visible by
    # default per the "prefer outline groups over hidden" rule. The +i
    # rollup col is left at outline_level=0 so it stays visible when the
    # -i band collapses.
    for mi in range(months):
        cl = col_letter(forecast_col_start + mi)
        ws.column_dimensions[cl].outline_level = 1

    # Forecast EUR +i header fill on the year/month header rows (green,
    # matching the band-label row).
    style_cell(ws, year_header_row, forecast_plus_i_col,
               fill=WITH_INDIRECT_HEADER_FILL)
    style_cell(ws, month_header_row, forecast_plus_i_col,
               fill=WITH_INDIRECT_HEADER_FILL)

    ws.cell(row=year_header_row, column=val_col_sum).comment = sized_comment("Row-wide SUM of hours logged this row. Per-person TOTAL and EU "
        "TOTAL rows sum across these.")

    # Cum-hours header: year row blank but fill-matched; month row carries
    # the label. Green so the reader sees this col as a helper aggregate.
    style_cell(ws, year_header_row, cum_hours_col, fill=LIGHT_GREEN_FILL)
    style_cell(ws, month_header_row, cum_hours_col,
               value="Cum hours (to date)", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
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
        wp["id"]: wp_active_months(wp["id"], work_packages)
        for wp in work_packages
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
            ws, r, member,
            ctx=ctx,
            person_assignments=per_person_assignments.get(name, []),
            forecast_col_start=forecast_col_start,
            forecast_col_end=forecast_col_end,
            forecast_plus_i_col=forecast_plus_i_col,
            cum_hours_col=cum_hours_col,
            preserved_for_person=preserved_for_person,
        )
        person_groups[name] = group

    # ── Reconciliation visual: grey out forecast-EUR leaf cells whose
    # month index is <= `eur_actuals_through_month`. Those months have
    # been closed by finance, so the forecast cell is reference-only —
    # the booked Actual EUR (on the Actuals sheet) is authoritative.
    # Mirror cell on the Actuals sheet gets the opposite treatment
    # (grey when month > threshold).
    if person_groups:
        last_row_used = max(
            grp["eu_total_row"] for grp in person_groups.values()
        )
        fc_start_l = col_letter(forecast_col_start)
        fc_end_l = col_letter(forecast_col_end)
        forecast_band_range = (
            f"{fc_start_l}{first_person_header_row}:"
            f"{fc_end_l}{last_row_used}"
        )
        ws.conditional_formatting.add(forecast_band_range, FormulaRule(
            formula=[
                f"COLUMN({fc_start_l}{first_person_header_row})"
                f"-COLUMN(${fc_start_l}$1)+1<=eur_actuals_through_month"
            ],
            fill=CF_GREY_FILL,
        ))
        # Second grey pass: forecast-EUR cells for months with no hours yet
        # logged (``month > hours_actuals_through_month``) are empty/unreliable
        # — shade so auditors don't mistake a blank for a zero forecast.
        ws.conditional_formatting.add(forecast_band_range, FormulaRule(
            formula=[
                f"COLUMN({fc_start_l}{first_person_header_row})"
                f"-COLUMN(${fc_start_l}$1)+1>hours_actuals_through_month"
            ],
            fill=CF_GREY_FILL,
        ))

    # ── Row 1: hours through-month control strip ──
    # A: label, B: override (blue; blank by default), C: auto-derived scan
    # of every person's programme-WP rows (excludes externals + per-person
    # subtotal/TOTAL/EU TOTAL rows), D: effective = override-if-numeric-
    # else-auto. ``hours_actuals_through_month`` now resolves to the
    # effective cell on this sheet rather than Config!B8.
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
            scan_terms.append(
                f"SUMPRODUCT(MAX((COLUMN({rng})-COLUMN(${ms_strip}$1)+1)"
                f"*ISNUMBER({rng})))"
            )
        auto_formula = f"=MAX({','.join(scan_terms)})"
    else:
        auto_formula = "=0"

    style_cell(ws, 1, 1, value="Hours Actuals Through Month:",
               font=HEADER_FONT, border=False)
    ws.cell(row=1, column=1).comment = sized_comment(
        "Trust boundary for hours-side cumulative figures. The effective "
        "value in column D is what Capacity, Plan-Forecast, and Trends "
        "consume via the named range `hours_actuals_through_month`. "
        "Override (B) beats auto-derived (C) only when it holds a number."
    )
    style_cell(ws, 1, 2, value=hours_through_override, font=BLUE_INPUT,
               fmt="0", border=False)
    ws.cell(row=1, column=2).comment = sized_comment(
        "Optional override — enter an integer month index (1-based) to "
        "force a tighter hours trust cutoff than the auto-derived value. "
        "Leave blank to defer to the auto scan in column C. Blue = user "
        "input."
    )
    style_cell(ws, 1, 3, font=BLACK_FORMULA, fmt="0", border=False).value = (
        auto_formula
    )
    ws.cell(row=1, column=3).comment = sized_comment(
        "Auto-derived: largest 1-based month index with any numeric hour "
        "value across every person's programme-WP rows (externals and subtotal/"
        "TOTAL/EU TOTAL rows are excluded). Returns 0 when no hours have "
        "been logged yet. Do not edit."
    )
    style_cell(ws, 1, 4, font=BLACK_FORMULA, fmt="0", border=False).value = (
        "=IF(ISNUMBER(B1),B1,C1)"
    )
    ws.cell(row=1, column=4).comment = sized_comment(
        "Effective hours trust cutoff consumed by Capacity, Plan-Forecast, "
        "and Trends via the named range `hours_actuals_through_month`. "
        "Override in B wins when numeric; otherwise falls through to the "
        "auto scan in C. Do not edit."
    )
    wb.defined_names.add(DefinedName(
        "hours_actuals_through_month",
        attr_text=f"'{SHEET_ACTUALS_HOURS}'!$D$1",
    ))

    ws.freeze_panes = f"C{first_person_header_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "band_label_row": band_label_row,
        "forecast_col_start": forecast_col_start,
        "forecast_col_end": forecast_col_end,
        "forecast_plus_i_col": forecast_plus_i_col,
        "cum_hours_col": cum_hours_col,
        "person_groups": person_groups,
    }
