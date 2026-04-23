"""Plan-Forecast sheet builder — forecast EUR aggregated from Actuals-Hours.

Mirrors the structure of ``plan.py`` but each monthly cell sums the
forecast-EUR leaf cells on the Actuals-Hours sheet for the persons
assigned to that WP.

Formula shape per WP × active-month cell::

    =SUM('Actuals-Hours'!<leafA>, 'Actuals-Hours'!<leafB>, ...)

- Leaves come from ``actuals_hours_layout["person_groups"]``: for each
  person assigned to the WP, locate that person's WP row, then pick
  the column at ``forecast_col_start + month_idx``.
- When no persons are assigned to the WP, the cell formula is ``=0``
  (explicit — avoids an empty ``SUM()``).
- Inactive months (WP not scheduled) are greyed out and empty.

Indirect-view layout (T4 of 2026-04-22 restructure):
- Monthly forecast grid wraps in a **column outline group** (hidden by
  default) — still without-indirect at source per D7. Section header
  suffixed "-i".
- Default-visible rollup columns are with-indirect (+i): Sum +i,
  Budget +i, Variance +i, Cum EUR +i, % of Budget +i. Green header fill.
- Inside the collapsible band sit the -i sibling rollups of the same
  five columns. Orange header fill. They feed accounting reconciliation
  when the band is expanded.
"""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG, SHEET_ACTUALS_HOURS, SHEET_PLAN_FORECAST,
    build_year_month_headers, month_col_end as _month_col_end, val_columns,
    wp_active_months,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT, SECTION_FONT,
    GREY_FILL, LIGHT_GREEN_FILL, SECTION_FILL,
    WITH_INDIRECT_HEADER_FILL, WITHOUT_INDIRECT_HEADER_FILL,
    CF_GREY_FILL, CF_RED_FILL, CENTER,
    style_cell, col_letter, sized_comment,
)


# Excel quotes sheet names containing non-alphanumeric chars (e.g. hyphens).
_AH_QUOTED = f"'{SHEET_ACTUALS_HOURS}'"
EUR_FMT = "€#,##0"


def _forecast_leaf_refs(wp_id, month_idx, *, assignments, actuals_hours_layout):
    """Return list of ``'Actuals-Hours'!<ref>`` strings for a WP × month.

    One ref per person assigned to ``wp_id``. The column is the same for
    all refs (forecast_col_start + month_idx); the row differs per person
    (their WP row on Actuals-Hours). Persons assigned to the WP but
    without a matching WP row on Actuals-Hours (shouldn't happen in
    practice — they share the same assignment source) are skipped.
    """
    forecast_col_start = actuals_hours_layout["forecast_col_start"]
    person_groups = actuals_hours_layout["person_groups"]
    col = forecast_col_start + month_idx
    cl = col_letter(col)

    # Deterministic person order: walk assignments for this WP; dedup.
    seen: set[str] = set()
    refs: list[str] = []
    for a in assignments:
        if a["wp"] != wp_id:
            continue
        name = a["person"]
        if name in seen:
            continue
        seen.add(name)
        group = person_groups.get(name)
        if group is None:
            continue
        wp_row = group["wp_rows"].get(wp_id)
        if wp_row is None:
            continue
        refs.append(f"{_AH_QUOTED}!{cl}{wp_row}")
    return refs


def build_plan_forecast_sheet(
    ws,
    *,
    project_params,
    work_packages,
    assignments,
    config_first_wp_row,
    actuals_hours_layout,
):
    """Build the Plan-Forecast sheet.

    Args:
        ws: Target worksheet.
        project_params: dict with ``title``, ``start``, ``months``, ``rate``.
        work_packages: list of WP dicts with single-range 'start'/'end'/'type'
            fields — drives active-month greying.
        assignments: list of person-WP assignment dicts; determines which
            persons' forecast-EUR leaves flow into each WP cell.
        config_first_wp_row: Row number of the first WP on Config sheet
            (for ID/Name/Budget cross-refs).
        actuals_hours_layout: dict returned by ``build_actuals_hours_sheet``.
            Must contain ``forecast_col_start`` and ``person_groups``.

    Returns:
        Layout dict with WP row positions + total row.
    """
    if assignments is None:
        assignments = []

    months = project_params["months"]
    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    m_end = _month_col_end(months)

    _vcols = val_columns(months)
    # -i rollups (live inside collapsible band — hidden by default).
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    val_col_cum = _vcols["cum"]
    val_col_pct = _vcols["pct_ceil"]
    # +i mirrors (default-visible outside outline band).
    val_col_sum_pi = _vcols["sum_plus_i"]
    val_col_budget_pi = _vcols["budget_plus_i"]
    val_col_diff_pi = _vcols["diff_plus_i"]
    val_col_cum_pi = _vcols["cum_plus_i"]
    val_col_pct_pi = _vcols["pct_ceil_plus_i"]
    # Columns that belong inside the collapsible -i outline band: the
    # monthly forecast grid (cols 3..m_end) plus the -i rollups. The +i
    # mirrors stay outside the band.
    _minus_i_band_cols = (
        *range(month_col_start, m_end + 1),
        val_col_sum, val_col_budget, val_col_diff,
        val_col_cum, val_col_pct,
    )

    ws.sheet_properties.tabColor = "00B050"  # green — forecast view

    cfg_wp_first = config_first_wp_row
    cfg_wp_last = cfg_wp_first + num_wps - 1

    # ── Row 1: compact header (Title / Start / Budget) ──
    # Title block
    style_cell(ws, 1, 1, "Title:", font=HEADER_FONT, border=False)
    style_cell(
        ws, 1, 2,
        font=Font(name="Arial", bold=True, size=14, color="008000"),
        border=False,
    ).value = f"={SHEET_CONFIG}!B2"

    # Start block
    style_cell(ws, 1, 4, "Start:", font=HEADER_FONT, border=False)
    style_cell(
        ws, 1, 5, font=GREEN_XREF, border=False, fmt="MMM YYYY",
    ).value = "=project_start"

    # Budget block — sum of WP Budget -i cells (same source as Plan!G1).
    # Cross-ref to the without-indirect column keeps numeric parity with
    # the legacy Ceiling header value that deferred consumers still read.
    style_cell(ws, 1, 6, "Budget:", font=HEADER_FONT, border=False)
    style_cell(
        ws, 1, 7, font=GREEN_XREF, border=False, fmt=EUR_FMT,
    ).value = f"=SUM({SHEET_CONFIG}!E{cfg_wp_first}:E{cfg_wp_last})"

    # -- FORECAST EUR block (starts at row 3) --
    r = 3
    # Two merged section-header regions: "-i" spans the collapsible band
    # (cols 1..val_col_pct); "+i" spans the default-visible mirror block
    # (val_col_sum_pi..val_col_pct_pi). Mirrors the Actuals pattern.
    def _section_header(first_col, last_col, label):
        ws.merge_cells(start_row=r, start_column=first_col,
                       end_row=r, end_column=last_col)
        style_cell(ws, r, first_col, label, font=SECTION_FONT,
                   fill=SECTION_FILL, alignment=CENTER)
        for c2 in range(first_col + 1, last_col + 1):
            style_cell(ws, r, c2, fill=SECTION_FILL)

    _section_header(1, val_col_pct, "FORECAST EUR (from hours) -i")
    _section_header(val_col_sum_pi, val_col_pct_pi,
                    "FORECAST EUR (from hours) +i")
    r += 1

    year_row, month_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            # -i rollups (collapsible band)
            (val_col_sum, "Sum -i"),
            (val_col_budget, "Budget -i"),
            (val_col_diff, "Variance -i"),
            (val_col_cum, "Cum EUR -i"),
            (val_col_pct, "% of Budget -i"),
            # +i mirrors (default-visible)
            (val_col_sum_pi, "Sum +i"),
            (val_col_budget_pi, "Budget +i"),
            (val_col_diff_pi, "Variance +i"),
            (val_col_cum_pi, "Cum EUR +i"),
            (val_col_pct_pi, "% of Budget +i"),
        ],
    )
    first_wp_row = r

    # Override rollup header fills so +i (green) / -i (orange) are visible
    # (default fill from build_year_month_headers is LIGHT_BLUE_FILL).
    _minus_i_header_cols = (
        val_col_sum, val_col_budget, val_col_diff,
        val_col_cum, val_col_pct,
    )
    _plus_i_header_cols = (
        val_col_sum_pi, val_col_budget_pi, val_col_diff_pi,
        val_col_cum_pi, val_col_pct_pi,
    )
    for c in _minus_i_header_cols:
        ws.cell(year_row, c).fill = WITHOUT_INDIRECT_HEADER_FILL
        ws.cell(month_row, c).fill = WITHOUT_INDIRECT_HEADER_FILL
    for c in _plus_i_header_cols:
        ws.cell(year_row, c).fill = WITH_INDIRECT_HEADER_FILL
        ws.cell(month_row, c).fill = WITH_INDIRECT_HEADER_FILL

    # Month-section label (-i suffix): anchor on the first calendar-year
    # cell at month_col_start so readers can tell the monthly grid is the
    # without-indirect band. Subsequent year cells stay unsuffixed.
    _year_anchor = ws.cell(year_row, month_col_start)
    if _year_anchor.value:
        _year_anchor.value = f"{_year_anchor.value} -i"
    else:
        _year_anchor.value = "-i"

    # ── WP rows ──
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = cfg_wp_first + i
        wp_id = wp["id"]

        # A: WP ID cross-ref  B: WP Name cross-ref
        style_cell(ws, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(ws, row, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        active = wp_active_months(wp_id, work_packages)
        for m in range(months):
            c = month_col_start + m
            month_num = m + 1
            if month_num in active:
                refs = _forecast_leaf_refs(
                    wp_id, m,
                    assignments=assignments,
                    actuals_hours_layout=actuals_hours_layout,
                )
                if refs:
                    formula = f"=SUM({', '.join(refs)})"
                else:
                    formula = "=0"
                style_cell(
                    ws, row, c, font=GREEN_XREF, fmt=EUR_FMT,
                ).value = formula
            else:
                style_cell(
                    ws, row, c, "", font=GREEN_XREF,
                    fill=GREY_FILL, fmt=EUR_FMT,
                )

        # ── -i rollup block (inside collapsible band) ──
        sum_range = (
            f"{col_letter(month_col_start)}{row}:{col_letter(m_end)}{row}"
        )
        # Sum -i: row-sum of monthly -i grid
        style_cell(
            ws, row, val_col_sum, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = f"=SUM({sum_range})"
        # Budget -i: cross-ref to Config WP Budget -i column (col 5).
        # Numerically identical to the legacy "Ceiling" column value.
        style_cell(
            ws, row, val_col_budget, font=GREEN_XREF, fmt=EUR_FMT,
        ).value = f"={SHEET_CONFIG}!E{config_row}"
        # Variance -i = Budget -i - Sum -i
        style_cell(
            ws, row, val_col_diff, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = (
            f"={col_letter(val_col_budget)}{row}"
            f"-{col_letter(val_col_sum)}{row}"
        )
        # Cum EUR -i — SUMPRODUCT masked by hours_actuals_through_month
        # (forecast-EUR is derived from hours timelog, so the hours-side
        # through-month scalar is the correct cadence anchor).
        ms = col_letter(month_col_start)
        me = col_letter(m_end)
        style_cell(
            ws, row, val_col_cum, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = (
            f"=SUMPRODUCT((COLUMN({ms}{row}:{me}{row})"
            f"-COLUMN(${ms}$1)+1<=hours_actuals_through_month)"
            f"*({ms}{row}:{me}{row}))"
        )
        # % of Budget -i = Cum -i / Budget -i
        cum_l = col_letter(val_col_cum)
        budget_l = col_letter(val_col_budget)
        style_cell(
            ws, row, val_col_pct, font=BLACK_FORMULA, fmt="0.0%",
        ).value = f"={cum_l}{row}/{budget_l}{row}"

        # ── +i mirror block (default-visible) ──
        # Sum +i = ROUND(Sum -i * 1.25, 0) per openpyxl.md EUR rule —
        # whole-unit format compounds fractional leaves into auditor
        # mismatches without ROUND at each leaf.
        sum_l = col_letter(val_col_sum)
        style_cell(
            ws, row, val_col_sum_pi, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = f"=ROUND({sum_l}{row}*1.25,0)"
        # Budget +i = INDEX(wp_budget_with_indirect, i+1). Targets the
        # Config Budget +i column directly — already with-indirect, so
        # NO ×1.25 applied here.
        style_cell(
            ws, row, val_col_budget_pi, font=GREEN_XREF, fmt=EUR_FMT,
        ).value = f"=INDEX(wp_budget_with_indirect,{i + 1})"
        # Variance +i = Budget +i - Sum +i
        sum_pi_l = col_letter(val_col_sum_pi)
        budget_pi_l = col_letter(val_col_budget_pi)
        style_cell(
            ws, row, val_col_diff_pi, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = f"={budget_pi_l}{row}-{sum_pi_l}{row}"
        # Cum EUR +i = ROUND(Cum -i * 1.25, 0)
        style_cell(
            ws, row, val_col_cum_pi, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = f"=ROUND({cum_l}{row}*1.25,0)"
        # % of Budget +i = Cum +i / Budget +i
        cum_pi_l = col_letter(val_col_cum_pi)
        style_cell(
            ws, row, val_col_pct_pi, font=BLACK_FORMULA, fmt="0.0%",
        ).value = f"={cum_pi_l}{row}/{budget_pi_l}{row}"

    last_wp_row = first_wp_row + num_wps - 1
    total_row = last_wp_row + 1

    # ── TOTAL row ──
    style_cell(ws, total_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(
        ws, total_row, 2, "MONTHLY TOTAL",
        font=HEADER_FONT, fill=LIGHT_GREEN_FILL,
    )
    for m in range(months):
        c = month_col_start + m
        cl = col_letter(c)
        style_cell(
            ws, total_row, c, font=BLACK_FORMULA,
            fill=LIGHT_GREEN_FILL, fmt=EUR_FMT,
        ).value = f"=SUM({cl}{first_wp_row}:{cl}{last_wp_row})"

    # ── TOTAL row: -i rollup totals ──
    for vcol in (val_col_sum, val_col_budget, val_col_diff, val_col_cum):
        cl = col_letter(vcol)
        style_cell(
            ws, total_row, vcol, font=BLACK_FORMULA,
            fill=LIGHT_GREEN_FILL, fmt=EUR_FMT,
        ).value = f"=SUM({cl}{first_wp_row}:{cl}{last_wp_row})"
    # % of Budget -i total = Cum -i total / Budget -i total
    style_cell(
        ws, total_row, val_col_pct, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="0.0%",
    ).value = (
        f"={col_letter(val_col_cum)}{total_row}"
        f"/{col_letter(val_col_budget)}{total_row}"
    )

    # ── TOTAL row: +i mirror totals ──
    for vcol in (val_col_sum_pi, val_col_budget_pi, val_col_diff_pi,
                 val_col_cum_pi):
        cl = col_letter(vcol)
        style_cell(
            ws, total_row, vcol, font=BLACK_FORMULA,
            fill=LIGHT_GREEN_FILL, fmt=EUR_FMT,
        ).value = f"=SUM({cl}{first_wp_row}:{cl}{last_wp_row})"
    # % of Budget +i total = Cum +i total / Budget +i total
    style_cell(
        ws, total_row, val_col_pct_pi, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="0.0%",
    ).value = (
        f"={col_letter(val_col_cum_pi)}{total_row}"
        f"/{col_letter(val_col_budget_pi)}{total_row}"
    )

    # Red CF on Variance -i when < 0 (Sum has exceeded Budget)
    diff_col_l = col_letter(val_col_diff)
    cf_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cf_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))
    # Red CF on Variance +i — same rule against the +i mirror column.
    diff_pi_col_l = col_letter(val_col_diff_pi)
    cf_pi_range = f"{diff_pi_col_l}{first_wp_row}:{diff_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cf_pi_range, FormulaRule(
        formula=[f"{diff_pi_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))

    # Cosmetic grey CF: forecast-EUR cells for months past
    # ``hours_actuals_through_month`` — no hours logged yet, so the forecast
    # aggregates zero contributions and should read as empty/unreliable
    # rather than a firm projection. Spans the WP block + TOTAL row.
    ms = col_letter(month_col_start)
    me = col_letter(m_end)
    forecast_body_range = f"{ms}{first_wp_row}:{me}{total_row}"
    ws.conditional_formatting.add(forecast_body_range, FormulaRule(
        formula=[
            f"COLUMN({ms}{first_wp_row})-COLUMN(${ms}$1)+1"
            f">hours_actuals_through_month"
        ],
        fill=CF_GREY_FILL,
    ))

    # Comment on the Budget header — explain provenance.
    ws.cell(row=1, column=7).comment = sized_comment(
        "Sum of WP Budget -i from Config. Same value as Plan!G1 and "
        "Actuals!G1 — verify all three agree."
    )

    # ── Outline group: collapsible -i band (T4 indirect-view
    # restructure — 2026-04-22). Monthly forecast grid + -i rollups all
    # collapse together; +i mirrors stay outside (default-visible).
    # `hidden=True` collapses the band by default per D3. Summary arrow
    # points right so the toggle sits next to the +i mirrors.
    for col in _minus_i_band_cols:
        dim = ws.column_dimensions[col_letter(col)]
        dim.outline_level = 1
        dim.hidden = True
    if ws.sheet_properties.outlinePr is None:
        from openpyxl.worksheet.properties import Outline
        ws.sheet_properties.outlinePr = Outline(summaryRight=True)
    else:
        ws.sheet_properties.outlinePr.summaryRight = True

    # Freeze panes: lock WP ID/Name cols + header rows
    ws.freeze_panes = f"C{first_wp_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "config_first_wp_row": cfg_wp_first,
        "config_last_wp_row": cfg_wp_last,
        "forecast_header_row": 3,
        "forecast_first_wp_row": first_wp_row,
        "forecast_last_wp_row": last_wp_row,
        "forecast_total_row": total_row,
    }


# Re-export for convenience
__all__ = ["build_plan_forecast_sheet", "SHEET_PLAN_FORECAST"]
