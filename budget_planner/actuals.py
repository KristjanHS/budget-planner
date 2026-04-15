"""Actuals sheet builder — EUR actuals per WP per month from accountant."""

from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_ACTUALS, SHEET_CONFIG, SHEET_PLAN_FORECAST,
    COL_WP_CEILING,
    ACTUALS_AMBER_PCT_CEIL, ACTUALS_AMBER_BREACH_MONTHS, ACTUALS_RED_BREACH_MONTHS,
    build_year_month_headers, cell_ref, month_range_ref,
    forecast_val_columns,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CF_RED_FILL, CF_AMBER_FILL, CF_GREY_FILL,
    style_cell, col_letter, write_section_header,
)


# Excel quotes sheet names containing non-alphanumeric chars (e.g. hyphens)
_PF_QUOTED = f"'{SHEET_PLAN_FORECAST}'"


def build_actuals_sheet(wb, layout, *, project_params=None, work_packages=None,
                        wp_schedule=None, preserved=None,
                        plan_forecast_layout=None):
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
    if project_params is None or work_packages is None or wp_schedule is None:
        from budget_planner import config
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES
        if wp_schedule is None:
            wp_schedule = config.WP_SCHEDULE

    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    val_col_cum = _vcols["cum"]
    val_col_pct_ceil = _vcols["pct_ceil"]
    val_col_breach = _vcols["breach"]

    _fcols = forecast_val_columns(project_params["months"])
    col_cum_forecast = _fcols["cum_forecast"]
    col_var_forecast = _fcols["var_forecast"]
    col_pct_ceil_forecast = _fcols["pct_ceil_forecast"]
    col_breach_forecast = _fcols["breach_forecast"]
    forecast_cols = (
        col_cum_forecast,
        col_var_forecast,
        col_pct_ceil_forecast,
        col_breach_forecast,
    )

    # Plan-Forecast row positions (WP order matches Actuals WP order)
    pf_first_wp_row = plan_forecast_layout["forecast_first_wp_row"]
    pf_total_row = plan_forecast_layout["forecast_total_row"]

    config_first_wp_row = layout["config_first_wp_row"]

    ws = wb.create_sheet(SHEET_ACTUALS)
    ws.sheet_properties.tabColor = "0000FF"

    r = 1

    # -- "As of" header row (row 1, above the section headers) --
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
    style_cell(ws, r, col_cum_forecast, value="As of:", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, border=False)
    style_cell(ws, r, col_cum_forecast + 1,
               value="=EDATE(project_start, hours_actuals_through_month - 1)",
               font=GREEN_XREF, fill=LIGHT_BLUE_FILL,
               fmt=_ASOF_DATE_FMT, border=False)
    r += 1

    # -- Section header --
    write_section_header(ws, r, "ACTUAL EUR", merge_end=val_col_breach)
    # Forecast-block section header on the same row, visually distinct
    ws.merge_cells(start_row=r, start_column=col_cum_forecast,
                   end_row=r, end_column=col_breach_forecast)
    style_cell(
        ws, r, col_cum_forecast, "FORECAST MONITORING",
        font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
    )
    for c in range(col_cum_forecast + 1, col_breach_forecast + 1):
        style_cell(ws, r, c, fill=LIGHT_BLUE_FILL)
    r += 1

    year_row, month_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            (val_col_sum, "Sum"),
            (val_col_budget, "Ceiling"),
            (val_col_diff, "Variance"),
            (val_col_cum, "Cum. EUR"),
            (val_col_pct_ceil, "% Ceiling"),
            (val_col_breach, "Breach In"),
            (col_cum_forecast, "Cum. Forecast EUR"),
            (col_var_forecast, "Forecast Variance"),
            (col_pct_ceil_forecast, "Forecast % Ceiling"),
            (col_breach_forecast, "Forecast Breach In"),
        ],
    )
    first_wp_row = r

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

        active = wp_active_months(wp["id"], wp_schedule)
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

        # Validation: Sum
        sum_range = month_range_ref(row, month_col_start, month_col_end)
        style_cell(ws, row, val_col_sum, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
        # Validation: Ceiling (from Config Personnel Ceiling column)
        style_cell(ws, row, val_col_budget, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = (
            f"={cell_ref(SHEET_CONFIG, config_row, COL_WP_CEILING, absolute=False)}"
        )
        # Validation: Variance = Ceiling - Sum
        style_cell(ws, row, val_col_diff, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(val_col_budget)}{row}"
            f"-{col_letter(val_col_sum)}{row}"
        )

        # Cumulative EUR (up to current_month)
        ms = col_letter(month_col_start)
        me = col_letter(month_col_end)
        style_cell(ws, row, val_col_cum, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"=SUMPRODUCT((COLUMN({ms}{row}:{me}{row})"
            f"-COLUMN(${ms}$1)+1<=current_month)"
            f"*({ms}{row}:{me}{row}))"
        )
        # % Ceiling = Cum. EUR / Ceiling
        cum_l = col_letter(val_col_cum)
        ceil_l = col_letter(val_col_budget)
        style_cell(ws, row, val_col_pct_ceil, font=BLACK_FORMULA,
                   fmt="0.0%").value = f"={cum_l}{row}/{ceil_l}{row}"
        # Breach In = months until ceiling exhausted at cumulative-to-date
        # rate. Denominator uses months_elapsed_eur so the slope reflects
        # spend per closed finance-month, not the raw project-month index.
        style_cell(ws, row, val_col_breach, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({cum_l}{row}=0,"\u2014",'
            f"({ceil_l}{row}-{cum_l}{row})"
            f"/({cum_l}{row}/months_elapsed_eur))"
        )

        # ── Forecast monitoring block (parallel columns on Plan-Forecast) ──
        # Source row on Plan-Forecast for this WP (same WP order)
        pf_row = pf_first_wp_row + i
        pf_ms = col_letter(month_col_start)
        pf_me = col_letter(month_col_end)
        pf_range = (
            f"{_PF_QUOTED}!{pf_ms}{pf_row}:{pf_me}{pf_row}"
        )
        # Cumulative Forecast EUR: SUMPRODUCT masked by current_month
        style_cell(
            ws, row, col_cum_forecast, font=GREEN_XREF, fmt="\u20ac#,##0",
        ).value = (
            f"=SUMPRODUCT((COLUMN({pf_range})"
            f"-COLUMN({_PF_QUOTED}!${pf_ms}$1)+1<=current_month)"
            f"*({pf_range}))"
        )
        cum_fore_l = col_letter(col_cum_forecast)
        # Forecast Variance = Ceiling - Cum Forecast
        style_cell(
            ws, row, col_var_forecast, font=BLACK_FORMULA, fmt="\u20ac#,##0",
        ).value = (
            f"={ceil_l}{row}-{cum_fore_l}{row}"
        )
        # Forecast % Ceiling = Cum Forecast / Ceiling
        style_cell(
            ws, row, col_pct_ceil_forecast, font=BLACK_FORMULA, fmt="0.0%",
        ).value = (
            f"={cum_fore_l}{row}/{ceil_l}{row}"
        )
        # Forecast Breach In — denominator uses months_elapsed_hours so the
        # forecast-spend slope is anchored to the hours-timelog cadence that
        # drives the forecast EUR values.
        style_cell(
            ws, row, col_breach_forecast, font=BLACK_FORMULA, fmt="0.0",
        ).value = (
            f'=IF({cum_fore_l}{row}=0,"\u2014",'
            f"({ceil_l}{row}-{cum_fore_l}{row})"
            f"/({cum_fore_l}{row}/months_elapsed_hours))"
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
    # TOTAL Sum
    sum_range = (f"{col_letter(val_col_sum)}{first_wp_row}:"
                 f"{col_letter(val_col_sum)}{last_wp_row}")
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
    # TOTAL Ceiling = partner_ceiling named range
    style_cell(ws, r, val_col_budget, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = "=partner_ceiling"
    # TOTAL Variance = Ceiling - Sum
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"={col_letter(val_col_budget)}{r}"
        f"-{col_letter(val_col_sum)}{r}"
    )
    # TOTAL Cumulative EUR
    ms = col_letter(month_col_start)
    me = col_letter(month_col_end)
    style_cell(ws, r, val_col_cum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"=SUMPRODUCT((COLUMN({ms}{r}:{me}{r})"
        f"-COLUMN(${ms}$1)+1<=current_month)"
        f"*({ms}{r}:{me}{r}))"
    )
    # TOTAL % Ceiling = Cum. EUR / partner_ceiling
    cum_l = col_letter(val_col_cum)
    style_cell(ws, r, val_col_pct_ceil, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.0%").value = f"={cum_l}{r}/partner_ceiling"
    # TOTAL Breach In — uses months_elapsed_eur, matching the WP rows.
    style_cell(ws, r, val_col_breach, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({cum_l}{r}=0,"\u2014",'
        f"(partner_ceiling-{cum_l}{r})"
        f"/({cum_l}{r}/months_elapsed_eur))"
    )

    # -- TOTAL row forecast block (mirrors Actual-EUR TOTAL, targets Plan-Forecast TOTAL row) --
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
        f"/({cum_fore_total_l}{r}/months_elapsed_hours))"
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

    # ── Conditional formatting: Forecast block (mirror Actual-EUR CF exactly) ──
    # Forecast Variance: red if < 0 (over ceiling)
    var_fore_l = col_letter(col_var_forecast)
    var_fore_range = f"{var_fore_l}{first_wp_row}:{var_fore_l}{last_wp_row}"
    ws.conditional_formatting.add(var_fore_range, FormulaRule(
        formula=[f"{var_fore_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))
    # Forecast % Ceiling: red >100%, amber >=ACTUALS_AMBER_PCT_CEIL
    pct_fore_l = col_letter(col_pct_ceil_forecast)
    pct_fore_range = f"{pct_fore_l}{first_wp_row}:{pct_fore_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_fore_range, FormulaRule(
        formula=[f"{pct_fore_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_fore_range, FormulaRule(
        formula=[f"{pct_fore_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"], fill=CF_AMBER_FILL,
    ))
    # Forecast Breach In: red <ACTUALS_RED_BREACH_MONTHS, amber <ACTUALS_AMBER_BREACH_MONTHS
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

    # ── Outline group: forecast-block columns collapsible as one band ──
    for col in forecast_cols:
        ws.column_dimensions[col_letter(col)].outline_level = 1
    # Group summary appears to the right of the grouped columns
    if ws.sheet_properties.outlinePr is None:
        from openpyxl.worksheet.properties import Outline
        ws.sheet_properties.outlinePr = Outline(summaryRight=True)
    else:
        ws.sheet_properties.outlinePr.summaryRight = True

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
