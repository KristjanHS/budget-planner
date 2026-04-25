"""Summary sheet builder — annual budget overview with per-year and project-total columns.

Calendar-year column groups (5 sub-columns/year: Plan FTE, Plan EUR,
Actual EUR, Forecast EUR, Variance EUR) + PROJECT TOTAL. Budget Overview
applies the **25% indirect uplift** to Total Claimable. Reporting Period
Summary emits **Form C preparation rows** grouped by Config reporting
periods — compliance-load-bearing, not free to restructure.

Role: read-only.

## Conditional formatting

- Remaining EUR column: **red** when ``value < 0``.
- Variance EUR (per year + project total): **red** when negative.
- Forecast EUR: **red** when ``Forecast < 0.8 * Plan EUR`` AND year is
  fully covered (``hours_actuals_through_month >= year_end_pm``) — gated
  so partial-coverage years don't trigger false-alarm underspend alerts;
  those are instead shaded grey (see next point).
- Per-year partial-coverage grey: Act € shaded by
  ``eur_actuals_through_month < year_end_pm``; Fcst € by
  ``hours_actuals_through_month < year_end_pm``; Var € by the stricter
  MIN. PROJECT TOTAL columns have no grey CF (span all years).
- Budget Overview Remaining: **amber** when < 10% of ceiling, **red**
  when negative.
"""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.layout import (
    MONTH_COL_START, SHEET_SUMMARY, SHEET_PLAN, SHEET_CONFIG, SHEET_ACTUALS,
    SHEET_PLAN_FORECAST,
    val_columns, _calendar_year_groups,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT, TITLE_FONT,
    SECTION_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CF_AMBER_FILL, CF_GREY_FILL, CF_RED_FILL, CENTER, style_cell, col_letter,
    write_section_header,
)


def build_summary_sheet(wb, layout, actuals_layout, *, project_params,
                        work_packages,
                        reporting_periods, plan_forecast_layout):
    """Create the Summary sheet with annual budget view and budget overview.

    Every data parameter is required — the orchestrator threads live data
    through, and there is no lazy-import fallback to ``config.py`` seed
    data (a caller that forgets to thread state would otherwise silently
    render seed data into the Summary sheet).

    Parameters
    ----------
    wb : openpyxl.Workbook
    layout : dict
        From build_plan_sheet(): config_first_wp_row, planned_first_wp_row.
    actuals_layout : dict
        From build_actuals_sheet(): first_wp_row, total_row.
    project_params : dict
        Keys: title, start, months, rate.
    work_packages : list[dict]
        List of WP dicts.
    reporting_periods : list[dict]
        Reporting period entries.

    Returns
    -------
    dict
        Summary layout with keys ``first_wp_row``, ``last_wp_row``, ``total_row``.
    """

    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    val_col_sum = val_columns(project_params["months"])["sum"]
    config_first_wp_row = layout["config_first_wp_row"]
    planned_first_wp_row = layout["planned_first_wp_row"]
    actuals_first_wp_row = actuals_layout["first_wp_row"]
    actuals_total_row = actuals_layout["total_row"]

    summary = wb.create_sheet(SHEET_SUMMARY)
    summary.sheet_properties.tabColor = "008000"

    r = 1
    summary.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(summary, r, 1, "ANNUAL BUDGET SUMMARY", font=TITLE_FONT, border=False)
    r += 2

    # Calculate calendar year groups
    project_start = project_params["start"]
    cal_groups = list(_calendar_year_groups(project_start, project_params["months"]))
    num_years = len(cal_groups)

    # Year group headers
    YEAR_GROUP_START = 3
    # Plan FTE, Plan EUR, Actual EUR, Forecast EUR, Variance EUR — same 5-column
    # shape for each calendar-year group AND for the PROJECT TOTAL group.
    COLS_PER_YEAR = 5

    # Project total group
    TOTAL_GROUP_START = YEAR_GROUP_START + num_years * COLS_PER_YEAR

    # Year headers row
    style_cell(summary, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(summary, r, 2, "", fill=LIGHT_BLUE_FILL)
    for y, (group_start, _span) in enumerate(cal_groups):
        start_col = YEAR_GROUP_START + y * COLS_PER_YEAR
        end_col = start_col + COLS_PER_YEAR - 1
        cal_date = project_start + relativedelta(months=group_start)
        summary.merge_cells(start_row=r, start_column=start_col, end_row=r, end_column=end_col)
        style_cell(summary, r, start_col, str(cal_date.year), font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(start_col + 1, end_col + 1):
            style_cell(summary, r, c2, fill=LIGHT_BLUE_FILL)

    summary.merge_cells(start_row=r, start_column=TOTAL_GROUP_START, end_row=r, end_column=TOTAL_GROUP_START + COLS_PER_YEAR - 1)
    style_cell(summary, r, TOTAL_GROUP_START, font=Font(name="Arial", bold=True, size=10, color="FFFFFF"), fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(TOTAL_GROUP_START + 1, TOTAL_GROUP_START + COLS_PER_YEAR):
        style_cell(summary, r, c2, fill=SECTION_FILL)
    r += 1

    # Sub-headers row
    style_cell(summary, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(summary, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    year_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Fcst \u20ac", "Var \u20ac"]
    for y in range(num_years):
        for si, sh in enumerate(year_sub):
            c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
            style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Fcst \u20ac", "Var \u20ac"]
    for si, sh in enumerate(total_sub):
        c = TOTAL_GROUP_START + si
        style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_GREEN_FILL, alignment=CENTER)
    r += 1

    SUMMARY_FIRST_WP_ROW = r

    # WP data rows
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = config_first_wp_row + i
        style_cell(summary, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
        style_cell(summary, row, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        for y, (group_start, span) in enumerate(cal_groups):
            first_col_l = col_letter(month_col_start + group_start)
            last_col_l = col_letter(month_col_start + group_start + span - 1)

            sc = YEAR_GROUP_START + y * COLS_PER_YEAR
            plan_wp_row = planned_first_wp_row + i
            actuals_wp_row = actuals_first_wp_row + i

            # Plan FTE (cross-sheet ref to Plan Planned block)
            style_cell(summary, row, sc, font=GREEN_XREF, fmt="0.00").value = (
                f"=SUM({SHEET_PLAN}!{first_col_l}{plan_wp_row}:{last_col_l}{plan_wp_row})"
            )
            # Plan EUR = Plan FTE * per-WP rate from Config col D
            style_cell(summary, row, sc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={col_letter(sc)}{row}*{SHEET_CONFIG}!$D${config_row}"
            )
            # Actual EUR (cross-sheet ref to Actuals sheet)
            style_cell(summary, row, sc + 2, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_ACTUALS}!{first_col_l}{actuals_wp_row}:{last_col_l}{actuals_wp_row})"
            )
            # Forecast EUR (cross-sheet ref to Plan-Forecast sheet).
            # Sheet name contains a hyphen → Excel requires single-quotes.
            # No current_month gate: forecast spans the full calendar year.
            pf_wp_row = plan_forecast_layout["forecast_first_wp_row"] + i
            style_cell(summary, row, sc + 3, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
                f"=SUM('{SHEET_PLAN_FORECAST}'!{first_col_l}{pf_wp_row}:{last_col_l}{pf_wp_row})"
            )
            # Variance EUR = Plan EUR - Actual EUR
            style_cell(summary, row, sc + 4, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
            )

        # Project totals — sum each year-group column across years.
        # `year_sum` closes over `row`; redefined per WP iteration by design.
        # Don't hoist it out of the loop — cell refs must track the current row.
        tc = TOTAL_GROUP_START
        def year_sum(offset: int) -> str:
            cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR + offset) for y in range(num_years)]
            return "+".join(f"{cl}{row}" for cl in cols)

        style_cell(summary, row, tc, font=BLACK_FORMULA, fmt="0.00").value = f"={year_sum(0)}"
        style_cell(summary, row, tc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={year_sum(1)}"
        style_cell(summary, row, tc + 2, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={year_sum(2)}"
        style_cell(summary, row, tc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={year_sum(3)}"
        # Total variance EUR = Plan EUR total - Actual EUR total
        style_cell(summary, row, tc + 4, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
            f"={col_letter(tc + 1)}{row}-{col_letter(tc + 2)}{row}"
        )

    SUMMARY_LAST_WP_ROW = SUMMARY_FIRST_WP_ROW + num_wps - 1
    totals_row = SUMMARY_LAST_WP_ROW + 1

    # Totals row
    style_cell(summary, totals_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(summary, totals_row, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for y in range(num_years):
        for si in range(COLS_PER_YEAR):
            c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
            fmt = "0.00" if si == 0 else "\u20ac#,##0"
            col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
            style_cell(summary, totals_row, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

    for si in range(COLS_PER_YEAR):
        c = TOTAL_GROUP_START + si
        fmt = "0.00" if si == 0 else "\u20ac#,##0"
        col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
        style_cell(summary, totals_row, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

    # Conditional formatting: red on negative Variance EUR (year columns)
    for y in range(num_years):
        var_col = YEAR_GROUP_START + y * COLS_PER_YEAR + 4
        var_col_l = col_letter(var_col)
        var_range = f"{var_col_l}{SUMMARY_FIRST_WP_ROW}:{var_col_l}{SUMMARY_LAST_WP_ROW}"
        summary.conditional_formatting.add(var_range, FormulaRule(
            formula=[f"{var_col_l}{SUMMARY_FIRST_WP_ROW}<0"], fill=CF_RED_FILL,
        ))

    # Conditional formatting: red on negative Variance EUR (project total)
    total_var_col = TOTAL_GROUP_START + 4
    total_var_col_l = col_letter(total_var_col)
    total_var_range = f"{total_var_col_l}{SUMMARY_FIRST_WP_ROW}:{total_var_col_l}{SUMMARY_LAST_WP_ROW}"
    summary.conditional_formatting.add(total_var_range, FormulaRule(
        formula=[f"{total_var_col_l}{SUMMARY_FIRST_WP_ROW}<0"], fill=CF_RED_FILL,
    ))

    # Conditional formatting: red on Forecast EUR < 0.8 * Plan EUR (year columns).
    # Mirrors the variance-style "material underspend" alert.
    # Gated on hours_actuals_through_month >= end-of-year so partial-coverage
    # years don't spuriously flag underspend; the grey partial-year rule below
    # takes care of those.
    for y, (group_start, span) in enumerate(cal_groups):
        plan_eur_col = YEAR_GROUP_START + y * COLS_PER_YEAR + 1
        fcst_col = YEAR_GROUP_START + y * COLS_PER_YEAR + 3
        plan_eur_col_l = col_letter(plan_eur_col)
        fcst_col_l = col_letter(fcst_col)
        fcst_range = f"{fcst_col_l}{SUMMARY_FIRST_WP_ROW}:{fcst_col_l}{SUMMARY_LAST_WP_ROW}"
        year_end_pm = group_start + span
        summary.conditional_formatting.add(fcst_range, FormulaRule(
            formula=[
                f"AND(hours_actuals_through_month>={year_end_pm},"
                f"{fcst_col_l}{SUMMARY_FIRST_WP_ROW}<0.8*{plan_eur_col_l}{SUMMARY_FIRST_WP_ROW})"
            ],
            fill=CF_RED_FILL,
        ))

    # Conditional formatting: grey on per-year partial-coverage columns.
    # Act €: year not fully covered by eur_actuals_through_month.
    # Fcst €: year not fully covered by hours_actuals_through_month.
    # Var €: stricter — either coverage short of the year end.
    # Skips PROJECT TOTAL (summarises across years; no single coverage gate).
    for y, (group_start, span) in enumerate(cal_groups):
        year_end_pm = group_start + span
        sc = YEAR_GROUP_START + y * COLS_PER_YEAR
        act_col_l = col_letter(sc + 2)
        fcst_col_l = col_letter(sc + 3)
        var_col_l = col_letter(sc + 4)

        act_range = f"{act_col_l}{SUMMARY_FIRST_WP_ROW}:{act_col_l}{SUMMARY_LAST_WP_ROW}"
        summary.conditional_formatting.add(act_range, FormulaRule(
            formula=[f"eur_actuals_through_month<{year_end_pm}"],
            fill=CF_GREY_FILL,
        ))

        fcst_range_grey = f"{fcst_col_l}{SUMMARY_FIRST_WP_ROW}:{fcst_col_l}{SUMMARY_LAST_WP_ROW}"
        summary.conditional_formatting.add(fcst_range_grey, FormulaRule(
            formula=[f"hours_actuals_through_month<{year_end_pm}"],
            fill=CF_GREY_FILL,
        ))

        var_range_grey = f"{var_col_l}{SUMMARY_FIRST_WP_ROW}:{var_col_l}{SUMMARY_LAST_WP_ROW}"
        summary.conditional_formatting.add(var_range_grey, FormulaRule(
            formula=[
                f"MIN(eur_actuals_through_month,hours_actuals_through_month)"
                f"<{year_end_pm}"
            ],
            fill=CF_GREY_FILL,
        ))

    # Conditional formatting: red on Forecast EUR < 0.8 * Plan EUR (project total)
    total_plan_eur_col = TOTAL_GROUP_START + 1
    total_fcst_col = TOTAL_GROUP_START + 3
    total_plan_eur_col_l = col_letter(total_plan_eur_col)
    total_fcst_col_l = col_letter(total_fcst_col)
    total_fcst_range = (
        f"{total_fcst_col_l}{SUMMARY_FIRST_WP_ROW}:"
        f"{total_fcst_col_l}{SUMMARY_LAST_WP_ROW}"
    )
    summary.conditional_formatting.add(total_fcst_range, FormulaRule(
        formula=[
            f"{total_fcst_col_l}{SUMMARY_FIRST_WP_ROW}"
            f"<0.8*{total_plan_eur_col_l}{SUMMARY_FIRST_WP_ROW}"
        ],
        fill=CF_RED_FILL,
    ))

    # ── BUDGET OVERVIEW ──
    r = totals_row + 2
    write_section_header(summary, r, "BUDGET OVERVIEW", merge_end=2)
    r += 2

    # Personnel Actual — total EUR from Actuals sheet
    actual_row = r
    style_cell(summary, r, 1, "Personnel Actual", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, font=GREEN_XREF, fmt="\u20ac#,##0",
               border=False).value = (
        f"={SHEET_ACTUALS}!{col_letter(val_col_sum)}{actuals_total_row}"
    )
    r += 1

    # Total Claimable (+25%)
    style_cell(summary, r, 1, "Total Claimable (+25%)", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, font=BLACK_FORMULA, fmt="\u20ac#,##0",
               border=False).value = f"=B{actual_row}*1.25"
    r += 1

    # Personnel Ceiling
    ceiling_row = r
    style_cell(summary, r, 1, "Personnel Ceiling", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, font=GREEN_XREF, fmt="\u20ac#,##0",
               border=False).value = "=partner_ceiling"
    r += 1

    # Remaining = Ceiling - Personnel Actual
    remaining_row = r
    style_cell(summary, r, 1, "Remaining", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, font=BLACK_FORMULA, fmt="\u20ac#,##0",
               border=False).value = f"=B{ceiling_row}-B{actual_row}"
    r += 1

    # % Consumed
    style_cell(summary, r, 1, "% Consumed", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, font=BLACK_FORMULA, fmt="0.0%",
               border=False).value = (
        f"=IF(B{ceiling_row}=0,0,B{actual_row}/B{ceiling_row})"
    )

    # Conditional formatting on Remaining — red first = highest priority
    # Red if negative
    summary.conditional_formatting.add(
        f"B{remaining_row}", FormulaRule(
            formula=[f"B{remaining_row}<0"], fill=CF_RED_FILL,
        ))
    # Amber if < 10% of ceiling
    summary.conditional_formatting.add(
        f"B{remaining_row}", FormulaRule(
            formula=[f"B{remaining_row}<B{ceiling_row}*0.1"], fill=CF_AMBER_FILL,
        ))

    # ── REPORTING PERIOD SUMMARY ──
    r += 2
    write_section_header(summary, r, "REPORTING PERIOD SUMMARY", merge_end=2)
    r += 2

    num_rps = len(reporting_periods)
    RP_GROUP_START = 3
    COLS_PER_RP = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR

    RP_TOTAL_GROUP_START = RP_GROUP_START + num_rps * COLS_PER_RP

    # RP headers row
    style_cell(summary, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(summary, r, 2, "", fill=LIGHT_BLUE_FILL)
    for rp_idx, rp in enumerate(reporting_periods):
        start_col = RP_GROUP_START + rp_idx * COLS_PER_RP
        end_col = start_col + COLS_PER_RP - 1
        label = f"{rp['id']} (M{rp['start']}\u2013M{rp['end']})"
        summary.merge_cells(start_row=r, start_column=start_col,
                            end_row=r, end_column=end_col)
        style_cell(summary, r, start_col, label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(start_col + 1, end_col + 1):
            style_cell(summary, r, c2, fill=LIGHT_BLUE_FILL)

    summary.merge_cells(start_row=r, start_column=RP_TOTAL_GROUP_START,
                        end_row=r, end_column=RP_TOTAL_GROUP_START + COLS_PER_RP - 1)
    style_cell(summary, r, RP_TOTAL_GROUP_START,
               font=Font(name="Arial", bold=True, size=10, color="FFFFFF"),
               fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(RP_TOTAL_GROUP_START + 1, RP_TOTAL_GROUP_START + COLS_PER_RP):
        style_cell(summary, r, c2, fill=SECTION_FILL)
    r += 1

    # RP sub-headers row
    style_cell(summary, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(summary, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    rp_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for rp_idx in range(num_rps):
        for si, sh in enumerate(rp_sub):
            c = RP_GROUP_START + rp_idx * COLS_PER_RP + si
            style_cell(summary, r, c, sh, font=HEADER_FONT,
                       fill=LIGHT_BLUE_FILL, alignment=CENTER)
    rp_total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for si, sh in enumerate(rp_total_sub):
        c = RP_TOTAL_GROUP_START + si
        style_cell(summary, r, c, sh, font=HEADER_FONT,
                   fill=LIGHT_GREEN_FILL, alignment=CENTER)
    r += 1

    RP_FIRST_WP_ROW = r

    # RP WP data rows
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = config_first_wp_row + i
        style_cell(summary, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(summary, row, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        for rp_idx, rp in enumerate(reporting_periods):
            first_col_l = col_letter(month_col_start + rp["start"] - 1)
            last_col_l = col_letter(month_col_start + rp["end"] - 1)

            sc = RP_GROUP_START + rp_idx * COLS_PER_RP
            plan_wp_row = planned_first_wp_row + i
            actuals_wp_row = actuals_first_wp_row + i

            # Plan FTE
            style_cell(summary, row, sc, font=GREEN_XREF, fmt="0.00").value = (
                f"=SUM({SHEET_PLAN}!{first_col_l}{plan_wp_row}:"
                f"{last_col_l}{plan_wp_row})"
            )
            # Plan EUR = Plan FTE * per-WP rate from Config col D
            style_cell(summary, row, sc + 1, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0").value = (
                f"={col_letter(sc)}{row}*{SHEET_CONFIG}!$D${config_row}"
            )
            # Actual EUR
            style_cell(summary, row, sc + 2, font=GREEN_XREF,
                       fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_ACTUALS}!{first_col_l}{actuals_wp_row}:"
                f"{last_col_l}{actuals_wp_row})"
            )
            # Variance EUR = Plan EUR - Actual EUR
            style_cell(summary, row, sc + 3, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0").value = (
                f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
            )

        # RP Project totals (sum across RP groups)
        tc = RP_TOTAL_GROUP_START
        rp_plan_cols = [col_letter(RP_GROUP_START + rp_idx * COLS_PER_RP)
                        for rp_idx in range(num_rps)]
        style_cell(summary, row, tc, font=BLACK_FORMULA, fmt="0.00").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_plan_cols)
        )
        rp_plan_eur_cols = [col_letter(RP_GROUP_START + rp_idx * COLS_PER_RP + 1)
                            for rp_idx in range(num_rps)]
        style_cell(summary, row, tc + 1, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_plan_eur_cols)
        )
        rp_act_eur_cols = [col_letter(RP_GROUP_START + rp_idx * COLS_PER_RP + 2)
                           for rp_idx in range(num_rps)]
        style_cell(summary, row, tc + 2, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_act_eur_cols)
        )
        style_cell(summary, row, tc + 3, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(tc + 1)}{row}-{col_letter(tc + 2)}{row}"
        )

    RP_LAST_WP_ROW = RP_FIRST_WP_ROW + num_wps - 1
    rp_totals_row = RP_LAST_WP_ROW + 1

    # RP TOTAL row
    style_cell(summary, rp_totals_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(summary, rp_totals_row, 2, "TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    for rp_idx in range(num_rps):
        for si in range(COLS_PER_RP):
            c = RP_GROUP_START + rp_idx * COLS_PER_RP + si
            fmt = "0.00" if si == 0 else "\u20ac#,##0"
            col_range = (f"{col_letter(c)}{RP_FIRST_WP_ROW}:"
                         f"{col_letter(c)}{RP_LAST_WP_ROW}")
            style_cell(summary, rp_totals_row, c, font=BLACK_FORMULA,
                       fill=LIGHT_GREEN_FILL, fmt=fmt).value = (
                f"=SUM({col_range})"
            )
    for si in range(COLS_PER_RP):
        c = RP_TOTAL_GROUP_START + si
        fmt = "0.00" if si == 0 else "\u20ac#,##0"
        col_range = (f"{col_letter(c)}{RP_FIRST_WP_ROW}:"
                     f"{col_letter(c)}{RP_LAST_WP_ROW}")
        style_cell(summary, rp_totals_row, c, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL, fmt=fmt).value = (
            f"=SUM({col_range})"
        )

    # CF: red on negative RP Variance EUR
    for rp_idx in range(num_rps):
        var_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 3
        var_col_l = col_letter(var_col)
        var_range = (f"{var_col_l}{RP_FIRST_WP_ROW}:"
                     f"{var_col_l}{RP_LAST_WP_ROW}")
        summary.conditional_formatting.add(var_range, FormulaRule(
            formula=[f"{var_col_l}{RP_FIRST_WP_ROW}<0"], fill=CF_RED_FILL,
        ))
    # CF: red on negative PROJECT TOTAL Variance EUR
    rp_total_var_col = RP_TOTAL_GROUP_START + 3
    rp_total_var_col_l = col_letter(rp_total_var_col)
    rp_total_var_range = (f"{rp_total_var_col_l}{RP_FIRST_WP_ROW}:"
                          f"{rp_total_var_col_l}{RP_LAST_WP_ROW}")
    summary.conditional_formatting.add(rp_total_var_range, FormulaRule(
        formula=[f"{rp_total_var_col_l}{RP_FIRST_WP_ROW}<0"], fill=CF_RED_FILL,
    ))

    # ── FORM C PREPARATION ──
    r = rp_totals_row + 2

    # Row labels and formulas use only the Act EUR sub-column (offset +2) per RP
    # A. Personnel EUR — reference TOTAL Act EUR for each RP
    personnel_fc_row = r
    style_cell(summary, r, 1, "A. Personnel EUR", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{rp_totals_row}"
        )
    # Project total
    total_act_col = RP_TOTAL_GROUP_START + 2
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{rp_totals_row}"
    )
    r += 1

    # E. Indirect (25%) — 25% of Personnel EUR
    indirect_row = r
    style_cell(summary, r, 1, "E. Indirect (25%)", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{personnel_fc_row}*0.25"
        )
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{personnel_fc_row}*0.25"
    )
    r += 1

    # Total Eligible = Personnel + Indirect
    eligible_row = r
    style_cell(summary, r, 1, "Total Eligible", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{personnel_fc_row}"
            f"+{col_letter(act_col)}{indirect_row}"
        )
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{personnel_fc_row}"
        f"+{col_letter(total_act_col)}{indirect_row}"
    )
    r += 1

    # Cumulative Eligible — running sum across RPs
    cum_row = r
    style_cell(summary, r, 1, "Cum. Eligible", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        if rp_idx == 0:
            style_cell(summary, r, act_col, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0", border=False).value = (
                f"={col_letter(act_col)}{eligible_row}"
            )
        else:
            prev_act_col = RP_GROUP_START + (rp_idx - 1) * COLS_PER_RP + 2
            style_cell(summary, r, act_col, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0", border=False).value = (
                f"={col_letter(prev_act_col)}{cum_row}"
                f"+{col_letter(act_col)}{eligible_row}"
            )
    # Project total cumulative = last RP's cumulative value
    last_rp_cum_col = RP_GROUP_START + (num_rps - 1) * COLS_PER_RP + 2
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(last_rp_cum_col)}{cum_row}"
    )
    r += 1

    # CFS Required? — flag when cumulative > EUR 430k
    CFS_THRESHOLD = 430_000  # EUR — audit certificate required above this
    cfs_row = r
    style_cell(summary, r, 1, "CFS Required?", font=HEADER_FONT, border=False)
    style_cell(summary, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        style_cell(summary, r, act_col, font=BLACK_FORMULA,
                   border=False).value = (
            f'=IF({col_letter(act_col)}{cum_row}>{CFS_THRESHOLD},"YES","No")'
        )
    style_cell(summary, r, total_act_col, font=BLACK_FORMULA,
               border=False).value = (
        f'=IF({col_letter(total_act_col)}{cum_row}>{CFS_THRESHOLD},"YES","No")'
    )

    # CF on CFS Required: amber fill on "YES"
    for rp_idx in range(num_rps):
        act_col = RP_GROUP_START + rp_idx * COLS_PER_RP + 2
        cell_ref = f"{col_letter(act_col)}{cfs_row}"
        summary.conditional_formatting.add(cell_ref, FormulaRule(
            formula=[f'{cell_ref}="YES"'], fill=CF_AMBER_FILL,
        ))
    total_cfs_ref = f"{col_letter(total_act_col)}{cfs_row}"
    summary.conditional_formatting.add(total_cfs_ref, FormulaRule(
        formula=[f'{total_cfs_ref}="YES"'], fill=CF_AMBER_FILL,
    ))

    summary.freeze_panes = f"C{SUMMARY_FIRST_WP_ROW}"
    summary.column_dimensions["B"].width = 25

    return {
        "first_wp_row": SUMMARY_FIRST_WP_ROW,
        "last_wp_row": SUMMARY_LAST_WP_ROW,
        "total_row": totals_row,
    }
