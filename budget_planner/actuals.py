"""Actuals sheet builder — EUR actuals per WP per month from accountant.

Role: monthly cells are an input surface for EUR spend (accountant
figures). Cross-sheet trust boundary (``eur_actuals_through_month`` named
range at ``Actuals!$D$1``) and the through-month / two-cadence rules are
canonical in ``docs/spec.md`` § Two-cadence actuals.

## Conditional formatting

- Negative variance: **red**.
- % of Budget (``-i`` and ``+i`` mirror rule — scale-invariant): **amber**
  > 85%, **red** > 100%.
- Breach In: **amber** < 12 months, **red** < 6 months.
- **Reconciliation grey (T8):** Actual-EUR monthly cells past
  ``eur_actuals_through_month`` are greyed (not yet finance-closed).
  Mirror rule on Actuals-Hours forecast band greys the *opposite* side
  (months ≤ ``eur_actuals_through_month``) so users see at a glance which
  side is authoritative per month.
- Parallel forecast-monitoring block: same amber/red thresholds as
  actual-EUR block on Cum Forecast Variance / % of Budget / Breach In
  (each mirrored on the ``+i`` companion columns).
"""

from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.layout import (
    MONTH_COL_START, SHEET_ACTUALS, SHEET_CONFIG, SHEET_PLAN_FORECAST,
    COL_WP_BUDGET_MINUS_I,
    ACTUALS_AMBER_PCT_CEIL, ACTUALS_AMBER_BREACH_MONTHS, ACTUALS_RED_BREACH_MONTHS,
    build_year_month_headers, cell_ref, month_range_ref,
    forecast_val_columns,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT, SECTION_FONT,
    GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    SECTION_FILL,
    WITH_INDIRECT_HEADER_FILL, WITHOUT_INDIRECT_HEADER_FILL,
    CF_RED_FILL, CF_AMBER_FILL, CF_GREY_FILL,
    CENTER,
    sized_comment, style_cell, col_letter,
)


# Excel quotes sheet names containing non-alphanumeric chars (e.g. hyphens)
_PF_QUOTED = f"'{SHEET_PLAN_FORECAST}'"


def build_actuals_sheet(wb, layout, *, project_params=None, work_packages=None,
                        preserved=None,
                        plan_forecast_layout=None,
                        eur_through_override: int | None = None):
    """Build the Actuals sheet for EUR actuals per WP per month.

    Monthly EUR cells are always left empty (blue input) for manual entry
    or demo seeding. When ``preserved`` is provided (rebuild), those values
    are restored.

    Args:
        wb: The parent Workbook.
        layout: Layout dict from build_plan_sheet(), needs config_first_wp_row.
        plan_forecast_layout: Layout dict from build_plan_forecast_sheet(),
            required to source the parallel forecast-monitoring columns.

    Returns:
        Layout dict with keys: header_row, year_row, month_row,
        first_wp_row, last_wp_row, total_row.
    """
    if plan_forecast_layout is None:
        raise ValueError(
            "build_actuals_sheet requires plan_forecast_layout — "
            "build Plan-Forecast before Actuals."
        )
    if project_params is None or work_packages is None:
        from budget_planner import config
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES

    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    # -i rollups (live inside collapsible band — hidden by default).
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    val_col_cum = _vcols["cum"]
    val_col_pct_ceil = _vcols["pct_ceil"]
    val_col_breach = _vcols["breach"]
    # +i mirrors (default-visible outside outline band).
    val_col_sum_pi = _vcols["sum_plus_i"]
    val_col_budget_pi = _vcols["budget_plus_i"]
    val_col_diff_pi = _vcols["diff_plus_i"]
    val_col_cum_pi = _vcols["cum_plus_i"]
    val_col_pct_ceil_pi = _vcols["pct_ceil_plus_i"]
    val_col_breach_pi = _vcols["breach_plus_i"]

    _fcols = forecast_val_columns(project_params["months"])
    # Forecast -i siblings (inside collapsible band).
    col_cum_forecast = _fcols["cum_forecast"]
    col_var_forecast = _fcols["var_forecast"]
    col_pct_ceil_forecast = _fcols["pct_ceil_forecast"]
    col_breach_forecast = _fcols["breach_forecast"]
    # Forecast +i mirrors (default-visible).
    col_cum_forecast_pi = _fcols["cum_forecast_plus_i"]
    col_var_forecast_pi = _fcols["var_forecast_plus_i"]
    col_pct_ceil_forecast_pi = _fcols["pct_ceil_forecast_plus_i"]
    col_breach_forecast_pi = _fcols["breach_forecast_plus_i"]
    # Columns that belong inside the collapsible -i outline band: the
    # monthly grid (cols 3..month_col_end), the -i rollups, and the
    # forecast -i siblings. +i mirrors stay outside.
    _minus_i_band_cols = (
        *range(month_col_start, month_col_end + 1),
        val_col_sum, val_col_budget, val_col_diff,
        val_col_cum, val_col_pct_ceil, val_col_breach,
        col_cum_forecast, col_var_forecast,
        col_pct_ceil_forecast, col_breach_forecast,
    )

    # Plan-Forecast row positions (WP order matches Actuals WP order)
    pf_first_wp_row = plan_forecast_layout["forecast_first_wp_row"]
    pf_total_row = plan_forecast_layout["forecast_total_row"]

    config_first_wp_row = layout["config_first_wp_row"]

    ws = wb.create_sheet(SHEET_ACTUALS)
    ws.sheet_properties.tabColor = "0000FF"

    # Row 1 holds the EUR through-month control strip (label / override /
    # auto-derived / effective). The strip formulas depend on
    # ``first_wp_row``/``last_wp_row``, so we write the placeholder cells
    # here to claim the row geometry and back-patch formulas + comments
    # after the WP block has been written. Trust-boundary invariant lives
    # in ``docs/spec.md`` § Two-cadence actuals.
    r = 2

    # -- "As of" header row (shifted to row 2 by the through-month strip) --
    # Two label/date pairs: one above the Actual EUR block, one above the
    # Forecast Monitoring block. Through-month cells hold integer month
    # indices; EDATE(project_start, N-1) turns them into a calendar date
    # so the header reads "As of: Apr 2026" not "1900-01".
    _ASOF_DATE_FMT = "MMM YYYY"
    style_cell(ws, r, 2, value="As of:", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, border=False)
    style_cell(ws, r, 3,
               value="=EDATE(project_start, eur_actuals_through_month - 1)",
               font=GREEN_XREF, fill=LIGHT_BLUE_FILL,
               fmt=_ASOF_DATE_FMT, border=False)
    # "As of" label duplicated at both the -i and +i forecast-block anchors.
    # The -i position preserves the legacy assertion in
    # ``test_actuals_breach_todate.py:113`` (cell at col_cum_forecast); the
    # +i position is where the user actually reads it now that the +i
    # mirror is the default-visible forecast surface.
    style_cell(ws, r, col_cum_forecast, value="As of:", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, border=False)
    style_cell(ws, r, col_cum_forecast + 1,
               value="=EDATE(project_start, hours_actuals_through_month - 1)",
               font=GREEN_XREF, fill=LIGHT_BLUE_FILL,
               fmt=_ASOF_DATE_FMT, border=False)
    style_cell(ws, r, col_cum_forecast_pi, value="As of:", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, border=False)
    style_cell(ws, r, col_cum_forecast_pi + 1,
               value="=EDATE(project_start, hours_actuals_through_month - 1)",
               font=GREEN_XREF, fill=LIGHT_BLUE_FILL,
               fmt=_ASOF_DATE_FMT, border=False)
    r += 1

    # -- Section headers --
    # ACTUAL EUR splits into two merged regions: "-i" over the
    # collapsible rollup block (cols 1..breach) and "+i" over the
    # default-visible mirror block (sum_pi..breach_pi). FORECAST
    # MONITORING mirrors the same split over its -i / +i sub-blocks.
    def _section_header(first_col, last_col, label):
        ws.merge_cells(start_row=r, start_column=first_col,
                       end_row=r, end_column=last_col)
        style_cell(ws, r, first_col, label, font=SECTION_FONT,
                   fill=SECTION_FILL, alignment=CENTER)
        for c2 in range(first_col + 1, last_col + 1):
            style_cell(ws, r, c2, fill=SECTION_FILL)

    _section_header(1, val_col_breach, "ACTUAL EUR -i")
    _section_header(val_col_sum_pi, val_col_breach_pi, "ACTUAL EUR +i")
    _section_header(col_cum_forecast, col_breach_forecast,
                    "FORECAST MONITORING -i")
    _section_header(col_cum_forecast_pi, col_breach_forecast_pi,
                    "FORECAST MONITORING +i")
    r += 1

    year_row, month_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            # -i rollup block (collapsible band)
            (val_col_sum, "Sum -i"),
            (val_col_budget, "Budget -i"),
            (val_col_diff, "Variance -i"),
            (val_col_cum, "Cum EUR -i"),
            (val_col_pct_ceil, "% of Budget -i"),
            (val_col_breach, "Breach-In -i"),
            # +i mirror block (default-visible)
            (val_col_sum_pi, "Sum +i"),
            (val_col_budget_pi, "Budget +i"),
            (val_col_diff_pi, "Variance +i"),
            (val_col_cum_pi, "Cum EUR +i"),
            (val_col_pct_ceil_pi, "% of Budget +i"),
            (val_col_breach_pi, "Breach-In +i"),
            # Forecast -i siblings (inside band)
            (col_cum_forecast, "Cum EUR -i"),
            (col_var_forecast, "Variance -i"),
            (col_pct_ceil_forecast, "% of Budget -i"),
            (col_breach_forecast, "Breach-In -i"),
            # Forecast +i mirrors (default-visible)
            (col_cum_forecast_pi, "Cum EUR +i"),
            (col_var_forecast_pi, "Variance +i"),
            (col_pct_ceil_forecast_pi, "% of Budget +i"),
            (col_breach_forecast_pi, "Breach-In +i"),
        ],
    )
    first_wp_row = r

    # Override header fills for the -i / +i blocks (default fill from
    # build_year_month_headers is LIGHT_BLUE_FILL — replace with the
    # T1 indirect-view fills so +i (green) and -i (orange) blocks are
    # visually distinguishable at the month-header row).
    _minus_i_header_cols = (
        val_col_sum, val_col_budget, val_col_diff,
        val_col_cum, val_col_pct_ceil, val_col_breach,
        col_cum_forecast, col_var_forecast,
        col_pct_ceil_forecast, col_breach_forecast,
    )
    _plus_i_header_cols = (
        val_col_sum_pi, val_col_budget_pi, val_col_diff_pi,
        val_col_cum_pi, val_col_pct_ceil_pi, val_col_breach_pi,
        col_cum_forecast_pi, col_var_forecast_pi,
        col_pct_ceil_forecast_pi, col_breach_forecast_pi,
    )
    for c in _minus_i_header_cols:
        ws.cell(year_row, c).fill = WITHOUT_INDIRECT_HEADER_FILL
        ws.cell(month_row, c).fill = WITHOUT_INDIRECT_HEADER_FILL
    for c in _plus_i_header_cols:
        ws.cell(year_row, c).fill = WITH_INDIRECT_HEADER_FILL
        ws.cell(month_row, c).fill = WITH_INDIRECT_HEADER_FILL

    # Month-section label (-i suffix): the year-row cells above the
    # monthly grid carry the calendar-year labels; append "-i" to the
    # first year-row cell so the section reads as without-indirect per
    # D8. Uses the first calendar-year cell (month_col_start) as the
    # anchor — subsequent year cells are unsuffixed.
    _year_anchor = ws.cell(year_row, month_col_start)
    if _year_anchor.value:
        _year_anchor.value = f"{_year_anchor.value} -i"
    else:
        _year_anchor.value = "-i"

    # -- WP rows (blue input EUR cells) --
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = config_first_wp_row + i
        # Cross-sheet refs to Config WP table
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        # Get preserved monthly values if available
        preserved_values = None
        if preserved and wp["id"] in preserved:
            preserved_values = preserved[wp["id"]]

        active = wp_active_months(wp["id"], work_packages)
        for m_idx in range(project_params["months"]):
            c = month_col_start + m_idx
            month_num = m_idx + 1
            if month_num in active:
                if preserved_values is not None and m_idx < len(preserved_values) and preserved_values[m_idx] is not None:
                    eur_value: int | float = preserved_values[m_idx]
                    style_cell(ws, row, c, eur_value, font=BLUE_INPUT, fmt="\u20ac#,##0")
                else:
                    style_cell(ws, row, c, font=BLUE_INPUT, fmt="\u20ac#,##0")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="\u20ac#,##0")

        # \u2500\u2500 -i rollup block (inside collapsible band) \u2500\u2500
        # Sum -i: row-sum of monthly -i grid
        sum_range = month_range_ref(row, month_col_start, month_col_end)
        style_cell(ws, row, val_col_sum, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
        # Budget -i (from Config WP Budget -i column \u2014 the black-formula
        # cell that derives = Budget +i / 1.25, so the numeric value here
        # matches the legacy Ceiling value)
        style_cell(ws, row, val_col_budget, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = (
            f"={cell_ref(SHEET_CONFIG, config_row, COL_WP_BUDGET_MINUS_I, absolute=False)}"
        )
        # Variance -i = Budget -i - Sum -i
        style_cell(ws, row, val_col_diff, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(val_col_budget)}{row}"
            f"-{col_letter(val_col_sum)}{row}"
        )

        # Cum EUR -i (up to current_month)
        ms = col_letter(month_col_start)
        me = col_letter(month_col_end)
        style_cell(ws, row, val_col_cum, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"=SUMPRODUCT((COLUMN({ms}{row}:{me}{row})"
            f"-COLUMN(${ms}$1)+1<=current_month)"
            f"*({ms}{row}:{me}{row}))"
        )
        # % of Budget -i = Cum -i / Budget -i
        cum_l = col_letter(val_col_cum)
        budget_l = col_letter(val_col_budget)
        style_cell(ws, row, val_col_pct_ceil, font=BLACK_FORMULA,
                   fmt="0.0%").value = f"={cum_l}{row}/{budget_l}{row}"
        # Breach-In -i = months until Budget -i exhausted at cumulative-
        # to-date rate. Denominator inlines MAX(1,
        # eur_actuals_through_month) so the slope reflects spend per
        # closed finance-month, not the raw project-month index.
        style_cell(ws, row, val_col_breach, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({cum_l}{row}=0,"\u2014",'
            f"({budget_l}{row}-{cum_l}{row})"
            f"/({cum_l}{row}/MAX(1,eur_actuals_through_month)))"
        )

        # \u2500\u2500 +i mirror block (default-visible) \u2500\u2500
        # Sum +i = ROUND(Sum -i * 1.25, 0). ROUND per openpyxl.md EUR rule
        # (whole-unit format compounds fractional leaves into auditor
        # mismatches).
        sum_l = col_letter(val_col_sum)
        style_cell(ws, row, val_col_sum_pi, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"=ROUND({sum_l}{row}*1.25,0)"
        )
        # Budget +i = INDEX(wp_budget_with_indirect, i+1). This targets
        # the Config Budget +i column directly and is already
        # with-indirect, so NO \u00d71.25 applied.
        style_cell(ws, row, val_col_budget_pi, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = (
            f"=INDEX(wp_budget_with_indirect,{i + 1})"
        )
        # Variance +i = Budget +i - Sum +i
        sum_pi_l = col_letter(val_col_sum_pi)
        budget_pi_l = col_letter(val_col_budget_pi)
        style_cell(ws, row, val_col_diff_pi, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={budget_pi_l}{row}-{sum_pi_l}{row}"
        )
        # Cum EUR +i = ROUND(Cum -i * 1.25, 0)
        style_cell(ws, row, val_col_cum_pi, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"=ROUND({cum_l}{row}*1.25,0)"
        )
        # % of Budget +i = Cum +i / Budget +i
        cum_pi_l = col_letter(val_col_cum_pi)
        style_cell(ws, row, val_col_pct_ceil_pi, font=BLACK_FORMULA,
                   fmt="0.0%").value = (
            f"={cum_pi_l}{row}/{budget_pi_l}{row}"
        )
        # Breach-In +i \u2014 same shape as -i, against the +i cells.
        style_cell(ws, row, val_col_breach_pi, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({cum_pi_l}{row}=0,"\u2014",'
            f"({budget_pi_l}{row}-{cum_pi_l}{row})"
            f"/({cum_pi_l}{row}/MAX(1,eur_actuals_through_month)))"
        )

        # ── Forecast monitoring block (parallel columns on Plan-Forecast) ──
        # Source row on Plan-Forecast for this WP (same WP order)
        pf_row = pf_first_wp_row + i
        pf_ms = col_letter(month_col_start)
        pf_me = col_letter(month_col_end)
        pf_range = (
            f"{_PF_QUOTED}!{pf_ms}{pf_row}:{pf_me}{pf_row}"
        )
        # -- Forecast -i siblings (inside collapsible band) --
        # Cum Forecast EUR -i: SUMPRODUCT masked by current_month (values
        # on Plan-Forecast are without-indirect per D7).
        style_cell(
            ws, row, col_cum_forecast, font=GREEN_XREF, fmt="\u20ac#,##0",
        ).value = (
            f"=SUMPRODUCT((COLUMN({pf_range})"
            f"-COLUMN({_PF_QUOTED}!${pf_ms}$1)+1<=current_month)"
            f"*({pf_range}))"
        )
        cum_fore_l = col_letter(col_cum_forecast)
        # Variance -i = Budget -i - Cum Forecast -i
        style_cell(
            ws, row, col_var_forecast, font=BLACK_FORMULA, fmt="\u20ac#,##0",
        ).value = (
            f"={budget_l}{row}-{cum_fore_l}{row}"
        )
        # % of Budget -i = Cum Forecast -i / Budget -i
        style_cell(
            ws, row, col_pct_ceil_forecast, font=BLACK_FORMULA, fmt="0.0%",
        ).value = (
            f"={cum_fore_l}{row}/{budget_l}{row}"
        )
        # Breach-In -i — denominator inlines MAX(1,
        # hours_actuals_through_month) so the forecast-spend slope is
        # anchored to the hours-timelog cadence that drives forecast EUR.
        style_cell(
            ws, row, col_breach_forecast, font=BLACK_FORMULA, fmt="0.0",
        ).value = (
            f'=IF({cum_fore_l}{row}=0,"\u2014",'
            f"({budget_l}{row}-{cum_fore_l}{row})"
            f"/({cum_fore_l}{row}/MAX(1,hours_actuals_through_month)))"
        )

        # -- Forecast +i mirrors (default-visible) --
        # Cum Forecast EUR +i = ROUND(Cum Forecast -i * 1.25, 0). Source
        # values on Plan-Forecast are -i per D7, so x1.25 applies at
        # the display rollup (D3/D7: indirect at consumer, not source).
        style_cell(
            ws, row, col_cum_forecast_pi, font=BLACK_FORMULA, fmt="\u20ac#,##0",
        ).value = (
            f"=ROUND({cum_fore_l}{row}*1.25,0)"
        )
        cum_fore_pi_l = col_letter(col_cum_forecast_pi)
        # Variance +i = Budget +i - Cum Forecast +i
        style_cell(
            ws, row, col_var_forecast_pi, font=BLACK_FORMULA, fmt="\u20ac#,##0",
        ).value = (
            f"={budget_pi_l}{row}-{cum_fore_pi_l}{row}"
        )
        # % of Budget +i = Cum Forecast +i / Budget +i
        style_cell(
            ws, row, col_pct_ceil_forecast_pi, font=BLACK_FORMULA, fmt="0.0%",
        ).value = (
            f"={cum_fore_pi_l}{row}/{budget_pi_l}{row}"
        )
        # Breach-In +i — same shape as -i, against the +i cells.
        style_cell(
            ws, row, col_breach_forecast_pi, font=BLACK_FORMULA, fmt="0.0",
        ).value = (
            f'=IF({cum_fore_pi_l}{row}=0,"\u2014",'
            f"({budget_pi_l}{row}-{cum_fore_pi_l}{row})"
            f"/({cum_fore_pi_l}{row}/MAX(1,hours_actuals_through_month)))"
        )

    last_wp_row = first_wp_row + num_wps - 1
    r = last_wp_row + 1

    # -- TOTAL row --
    total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="\u20ac#,##0").value = f"=SUM({col_range})"

    # ── TOTAL row: -i rollups (inside collapsible band) ──
    # Sum -i = SUM over WP -i Sum column
    sum_range = (f"{col_letter(val_col_sum)}{first_wp_row}:"
                 f"{col_letter(val_col_sum)}{last_wp_row}")
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
    # Budget -i = partner_ceiling (without-indirect).
    style_cell(ws, r, val_col_budget, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = "=partner_ceiling"
    # Variance -i = Budget -i - Sum -i
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"={col_letter(val_col_budget)}{r}"
        f"-{col_letter(val_col_sum)}{r}"
    )
    # Cum -i (over monthly grid)
    ms = col_letter(month_col_start)
    me = col_letter(month_col_end)
    style_cell(ws, r, val_col_cum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"=SUMPRODUCT((COLUMN({ms}{r}:{me}{r})"
        f"-COLUMN(${ms}$1)+1<=current_month)"
        f"*({ms}{r}:{me}{r}))"
    )
    # % of Budget -i = Cum -i / partner_ceiling
    cum_l = col_letter(val_col_cum)
    style_cell(ws, r, val_col_pct_ceil, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.0%").value = f"={cum_l}{r}/partner_ceiling"
    # Breach-In -i — inlines MAX(1, eur_actuals_through_month), matching
    # the WP rows.
    style_cell(ws, r, val_col_breach, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({cum_l}{r}=0,"\u2014",'
        f"(partner_ceiling-{cum_l}{r})"
        f"/({cum_l}{r}/MAX(1,eur_actuals_through_month)))"
    )

    # ── TOTAL row: +i mirrors (default-visible) ──
    # Sum +i = SUM over WP +i Sum column (each WP row +i Sum is already
    # ROUND(Sum_-i * 1.25, 0), so the column-sum stays consistent).
    sum_pi_range = (f"{col_letter(val_col_sum_pi)}{first_wp_row}:"
                    f"{col_letter(val_col_sum_pi)}{last_wp_row}")
    style_cell(ws, r, val_col_sum_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = f"=SUM({sum_pi_range})"
    # Budget +i = partner_ceiling * 1.25. (The partner_ceiling named range
    # is -i; consumers apply ×1.25 — no separate +i named range.)
    style_cell(ws, r, val_col_budget_pi, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = "=ROUND(partner_ceiling*1.25,0)"
    # Variance +i = Budget +i - Sum +i
    style_cell(ws, r, val_col_diff_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"={col_letter(val_col_budget_pi)}{r}"
        f"-{col_letter(val_col_sum_pi)}{r}"
    )
    # Cum +i = ROUND(Cum -i * 1.25, 0)
    style_cell(ws, r, val_col_cum_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"=ROUND({cum_l}{r}*1.25,0)"
    )
    # % of Budget +i = Cum +i / (partner_ceiling * 1.25)
    cum_pi_l = col_letter(val_col_cum_pi)
    style_cell(ws, r, val_col_pct_ceil_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.0%").value = f"={cum_pi_l}{r}/(partner_ceiling*1.25)"
    # Breach-In +i — same shape as -i, against +i cells.
    style_cell(ws, r, val_col_breach_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({cum_pi_l}{r}=0,"\u2014",'
        f"((partner_ceiling*1.25)-{cum_pi_l}{r})"
        f"/({cum_pi_l}{r}/MAX(1,eur_actuals_through_month)))"
    )

    # ── TOTAL row: forecast -i siblings (inside collapsible band) ──
    pf_ms = col_letter(month_col_start)
    pf_me = col_letter(month_col_end)
    pf_total_range = (
        f"{_PF_QUOTED}!{pf_ms}{pf_total_row}:{pf_me}{pf_total_row}"
    )
    style_cell(
        ws, r, col_cum_forecast, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
        fmt="\u20ac#,##0",
    ).value = (
        f"=SUMPRODUCT((COLUMN({pf_total_range})"
        f"-COLUMN({_PF_QUOTED}!${pf_ms}$1)+1<=current_month)"
        f"*({pf_total_range}))"
    )
    cum_fore_total_l = col_letter(col_cum_forecast)
    style_cell(
        ws, r, col_var_forecast, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0",
    ).value = (
        f"=partner_ceiling-{cum_fore_total_l}{r}"
    )
    style_cell(
        ws, r, col_pct_ceil_forecast, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="0.0%",
    ).value = (
        f"={cum_fore_total_l}{r}/partner_ceiling"
    )
    style_cell(
        ws, r, col_breach_forecast, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="0.0",
    ).value = (
        f'=IF({cum_fore_total_l}{r}=0,"\u2014",'
        f"(partner_ceiling-{cum_fore_total_l}{r})"
        f"/({cum_fore_total_l}{r}/MAX(1,hours_actuals_through_month)))"
    )

    # ── TOTAL row: forecast +i mirrors (default-visible) ──
    style_cell(
        ws, r, col_cum_forecast_pi, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0",
    ).value = f"=ROUND({cum_fore_total_l}{r}*1.25,0)"
    cum_fore_total_pi_l = col_letter(col_cum_forecast_pi)
    style_cell(
        ws, r, col_var_forecast_pi, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0",
    ).value = (
        f"=ROUND(partner_ceiling*1.25,0)-{cum_fore_total_pi_l}{r}"
    )
    style_cell(
        ws, r, col_pct_ceil_forecast_pi, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="0.0%",
    ).value = (
        f"={cum_fore_total_pi_l}{r}/(partner_ceiling*1.25)"
    )
    style_cell(
        ws, r, col_breach_forecast_pi, font=BLACK_FORMULA,
        fill=LIGHT_GREEN_FILL, fmt="0.0",
    ).value = (
        f'=IF({cum_fore_total_pi_l}{r}=0,"\u2014",'
        f"((partner_ceiling*1.25)-{cum_fore_total_pi_l}{r})"
        f"/({cum_fore_total_pi_l}{r}/MAX(1,hours_actuals_through_month)))"
    )

    # -- Conditional formatting: red on negative variance --
    diff_col_l = col_letter(val_col_diff)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))

    # -- Conditional formatting: % Ceiling column (amber >ACTUALS_AMBER_PCT_CEIL, red >100%) --
    # Red first = highest priority (lower number wins in Excel CF)
    pct_col_l = col_letter(val_col_pct_ceil)
    pct_range = f"{pct_col_l}{first_wp_row}:{pct_col_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"], fill=CF_AMBER_FILL,
    ))

    # -- Conditional formatting: Breach In column (amber <ACTUALS_AMBER_BREACH_MONTHS, red <ACTUALS_RED_BREACH_MONTHS) --
    # Red first = highest priority
    breach_col_l = col_letter(val_col_breach)
    breach_range = f"{breach_col_l}{first_wp_row}:{breach_col_l}{last_wp_row}"
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"], fill=CF_AMBER_FILL,
    ))

    # ── Conditional formatting: Forecast -i block (mirror Actual-EUR -i CF exactly) ──
    # Forecast Variance -i: red if < 0 (over budget)
    var_fore_l = col_letter(col_var_forecast)
    var_fore_range = f"{var_fore_l}{first_wp_row}:{var_fore_l}{last_wp_row}"
    ws.conditional_formatting.add(var_fore_range, FormulaRule(
        formula=[f"{var_fore_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))
    # Forecast % of Budget -i: red >100%, amber >=ACTUALS_AMBER_PCT_CEIL
    pct_fore_l = col_letter(col_pct_ceil_forecast)
    pct_fore_range = f"{pct_fore_l}{first_wp_row}:{pct_fore_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_fore_range, FormulaRule(
        formula=[f"{pct_fore_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_fore_range, FormulaRule(
        formula=[f"{pct_fore_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"], fill=CF_AMBER_FILL,
    ))
    # Forecast Breach-In -i: red <ACTUALS_RED_BREACH_MONTHS, amber <ACTUALS_AMBER_BREACH_MONTHS
    breach_fore_l = col_letter(col_breach_forecast)
    breach_fore_range = (
        f"{breach_fore_l}{first_wp_row}:{breach_fore_l}{last_wp_row}"
    )
    ws.conditional_formatting.add(breach_fore_range, FormulaRule(
        formula=[f"{breach_fore_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"],
        fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(breach_fore_range, FormulaRule(
        formula=[f"{breach_fore_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"],
        fill=CF_AMBER_FILL,
    ))

    # ── Conditional formatting: +i mirror block (T3 — intentional CF
    # rule-count jump; one-time bump permitted per testing.md). Every
    # -i CF rule above has a +i sibling targeting the mirror columns. ──
    # Variance +i: red if < 0
    diff_pi_col_l = col_letter(val_col_diff_pi)
    diff_pi_range = f"{diff_pi_col_l}{first_wp_row}:{diff_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(diff_pi_range, FormulaRule(
        formula=[f"{diff_pi_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))
    # % of Budget +i: red >100%, amber >=ACTUALS_AMBER_PCT_CEIL
    pct_pi_col_l = col_letter(val_col_pct_ceil_pi)
    pct_pi_range = f"{pct_pi_col_l}{first_wp_row}:{pct_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_pi_range, FormulaRule(
        formula=[f"{pct_pi_col_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_pi_range, FormulaRule(
        formula=[f"{pct_pi_col_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"],
        fill=CF_AMBER_FILL,
    ))
    # Breach-In +i: red/amber bands
    breach_pi_col_l = col_letter(val_col_breach_pi)
    breach_pi_range = f"{breach_pi_col_l}{first_wp_row}:{breach_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(breach_pi_range, FormulaRule(
        formula=[f"{breach_pi_col_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"],
        fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(breach_pi_range, FormulaRule(
        formula=[f"{breach_pi_col_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"],
        fill=CF_AMBER_FILL,
    ))
    # Forecast Variance +i: red if < 0
    var_fore_pi_l = col_letter(col_var_forecast_pi)
    var_fore_pi_range = f"{var_fore_pi_l}{first_wp_row}:{var_fore_pi_l}{last_wp_row}"
    ws.conditional_formatting.add(var_fore_pi_range, FormulaRule(
        formula=[f"{var_fore_pi_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))
    # Forecast % of Budget +i: red >100%, amber >=ACTUALS_AMBER_PCT_CEIL
    pct_fore_pi_l = col_letter(col_pct_ceil_forecast_pi)
    pct_fore_pi_range = f"{pct_fore_pi_l}{first_wp_row}:{pct_fore_pi_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_fore_pi_range, FormulaRule(
        formula=[f"{pct_fore_pi_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_fore_pi_range, FormulaRule(
        formula=[f"{pct_fore_pi_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"],
        fill=CF_AMBER_FILL,
    ))
    # Forecast Breach-In +i: red/amber bands
    breach_fore_pi_l = col_letter(col_breach_forecast_pi)
    breach_fore_pi_range = (
        f"{breach_fore_pi_l}{first_wp_row}:{breach_fore_pi_l}{last_wp_row}"
    )
    ws.conditional_formatting.add(breach_fore_pi_range, FormulaRule(
        formula=[f"{breach_fore_pi_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"],
        fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(breach_fore_pi_range, FormulaRule(
        formula=[f"{breach_fore_pi_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"],
        fill=CF_AMBER_FILL,
    ))

    # ── Reconciliation visual: grey out monthly Actual EUR cells whose
    # month index is past the user-controlled `eur_actuals_through_month`
    # threshold. Those months haven't been closed by finance yet, so any
    # value (including blanks) should read as stale/incoming rather than
    # authoritative. Mirror cell on Actuals-Hours forecast band gets the
    # opposite treatment (grey when month <= threshold).
    ms = col_letter(month_col_start)
    me = col_letter(month_col_end)
    actual_eur_range = f"{ms}{first_wp_row}:{me}{total_row}"
    ws.conditional_formatting.add(actual_eur_range, FormulaRule(
        formula=[
            f"COLUMN({ms}{first_wp_row})-COLUMN(${ms}$1)+1"
            f">eur_actuals_through_month"
        ],
        fill=CF_GREY_FILL,
    ))

    # ── Outline group: collapsible -i band (T3 indirect-view
    # restructure — 2026-04-22). Bands together: monthly -i grid
    # (cols C..me), -i rollups (Sum/Budget/Variance/Cum/%/Breach),
    # forecast -i siblings. +i mirrors stay outside (default-visible).
    # `hidden=True` collapses the band by default per D3: "Actual EUR
    # -i rollup sits inside the collapsible band — hidden by default
    # under the collapsed group". Summary arrow points right so the
    # group toggle sits next to the +i mirrors.
    for col in _minus_i_band_cols:
        dim = ws.column_dimensions[col_letter(col)]
        dim.outline_level = 1
        dim.hidden = True
    if ws.sheet_properties.outlinePr is None:
        from openpyxl.worksheet.properties import Outline
        ws.sheet_properties.outlinePr = Outline(summaryRight=True)
    else:
        ws.sheet_properties.outlinePr.summaryRight = True

    # ── Row 1: EUR through-month control strip ──
    # A: label, B: override (blue; blank by default), C: auto-derived scan
    # of the WP data block (excludes TOTAL), D: effective = override-if-
    # numeric-else-auto. ``eur_actuals_through_month`` now resolves to the
    # effective cell on this sheet rather than Config!B7.
    style_cell(ws, 1, 1, value="EUR Actuals Through Month:",
               font=HEADER_FONT, border=False)
    ws.cell(row=1, column=1).comment = sized_comment(
        "Trust boundary for EUR-side cumulative figures. The effective "
        "value in column D is what downstream sheets consume via the "
        "named range `eur_actuals_through_month`. Override (B) beats auto-"
        "derived (C) only when it holds a number."
    )
    style_cell(ws, 1, 2, value=eur_through_override, font=BLUE_INPUT, fmt="0",
               border=False)
    ws.cell(row=1, column=2).comment = sized_comment(
        "Optional override — enter an integer month index (1-based) to "
        "force a tighter EUR trust cutoff than the auto-derived value. "
        "Leave blank to defer to the auto scan in column C. Blue = user "
        "input."
    )
    ms_strip = col_letter(month_col_start)
    me_strip = col_letter(month_col_end)
    wp_scan_range = (
        f"${ms_strip}${first_wp_row}:${me_strip}${last_wp_row}"
    )
    style_cell(ws, 1, 3, font=BLACK_FORMULA, fmt="0", border=False).value = (
        f"=SUMPRODUCT(MAX((COLUMN({wp_scan_range})"
        f"-COLUMN(${ms_strip}$1)+1)"
        f"*ISNUMBER({wp_scan_range})))"
    )
    ws.cell(row=1, column=3).comment = sized_comment(
        "Auto-derived: largest 1-based month index with any numeric "
        "value in the WP data block (TOTAL row excluded). Returns 0 "
        "when no actuals have been logged yet. Do not edit."
    )
    style_cell(ws, 1, 4, font=BLACK_FORMULA, fmt="0", border=False).value = (
        "=IF(ISNUMBER(B1),B1,C1)"
    )
    ws.cell(row=1, column=4).comment = sized_comment(
        "Effective EUR trust cutoff consumed by Actuals, Summary, and "
        "Trends via the named range `eur_actuals_through_month`. Override "
        "in B wins when numeric; otherwise falls through to the auto scan "
        "in C. Do not edit."
    )
    wb.defined_names.add(DefinedName(
        "eur_actuals_through_month",
        attr_text=f"{SHEET_ACTUALS}!$D$1",
    ))

    # Freeze panes
    ws.freeze_panes = f"C{first_wp_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "asof_row": 1,       # "As of" header row (T7 addition; section header shifted to row 2)
        "year_row": year_row,
        "month_row": month_row,
        "first_wp_row": first_wp_row,
        "last_wp_row": last_wp_row,
        "total_row": total_row,
        "cum_col": val_col_cum,
        "pct_ceiling_col": val_col_pct_ceil,
        "breach_col": val_col_breach,
        "cum_forecast_col": col_cum_forecast,
        "var_forecast_col": col_var_forecast,
        "pct_ceiling_forecast_col": col_pct_ceil_forecast,
        "breach_forecast_col": col_breach_forecast,
    }
