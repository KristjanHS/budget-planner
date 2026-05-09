"""Summary sheet builder — annual budget overview with per-year and project-total columns.

Calendar-year column groups (4 sub-columns/year: Plan FTE, Plan EUR,
Actual EUR, Variance EUR) + PROJECT TOTAL. A single project-end
**Forecast EUR** column sits between the year band and PROJECT TOTAL,
projecting from the Actuals posted-cumulative cells (posted_idx +
posted_value + slope) out to project_months. Budget Overview applies
the **25% indirect uplift** to Total Claimable. Reporting Period
Summary emits **Form C preparation rows** grouped by Config reporting
periods — compliance-load-bearing, not free to restructure.

Role: read-only.

## Conditional formatting

- Remaining EUR column: **red** when ``value < 0``.
- Variance EUR (per year + project total): **red** when negative.
- Forecast EUR (single project-end column): **red** when
  ``Forecast < 0.8 * Plan EUR`` (project-total Plan EUR), with an
  ``<>""`` guard so the empty-sentinel ("" when no postings yet)
  doesn't compare-fire the rule.
- Per-year partial-coverage grey: Act € shaded by
  ``eur_actuals_through_month < year_end_pm``; Var € by the stricter
  MIN of EUR/hours coverage. PROJECT TOTAL columns have no grey CF
  (span all years).
- Budget Overview Remaining: **amber** when < 10% of ceiling, **red**
  when negative.
"""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.layout import (
    MONTH_COL_START, SHEET_SUMMARY, SHEET_PLAN, SHEET_CONFIG, SHEET_ACTUALS,
    val_columns, _calendar_year_groups,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT, TITLE_FONT,
    SECTION_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CF_AMBER_FILL, CF_GREY_FILL, CF_RED_FILL, CENTER, style_cell, col_letter,
    write_section_header,
)


# Plan FTE, Plan EUR, Actual EUR, Variance EUR — same 4-column shape for each
# calendar-year group AND for the PROJECT TOTAL group. The project-end
# Forecast EUR sits in its own single column between the year band and
# PROJECT TOTAL.
_COLS_PER_YEAR = 4
_YEAR_GROUP_START = 3
_COLS_PER_RP = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR
_RP_GROUP_START = 3


def build_summary_sheet(wb, layout, actuals_layout, *, project_params,
                        work_packages,
                        reporting_periods):
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
        From build_actuals_sheet(): first_wp_row, total_row, plus the
        posted-cumulative redesign keys ``posted_idx_col`` /
        ``posted_value_col`` / ``slope_col`` consumed by the project-end
        Forecast EUR column.
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
    num_wps = len(work_packages)
    val_col_sum = val_columns(project_params["months"])["sum"]
    config_first_wp_row = layout["config_first_wp_row"]
    planned_first_wp_row = layout["planned_first_wp_row"]
    actuals_first_wp_row = actuals_layout["first_wp_row"]
    actuals_total_row = actuals_layout["total_row"]

    summary = wb.create_sheet(SHEET_SUMMARY)
    summary.sheet_properties.tabColor = "008000"

    cal_groups = list(_calendar_year_groups(
        project_params["start"], project_params["months"]))
    num_years = len(cal_groups)
    forecast_col = _YEAR_GROUP_START + num_years * _COLS_PER_YEAR
    total_group_start = forecast_col + 1

    r = _write_title(summary)
    r = _write_year_headers(summary, r, cal_groups, project_params["start"],
                            num_years, forecast_col, total_group_start)
    r = _write_year_subheaders(summary, r, num_years, forecast_col,
                               total_group_start)
    summary_first_wp_row = r
    _write_wp_data_rows(summary, r, work_packages, cal_groups, num_years,
                        forecast_col, total_group_start, config_first_wp_row,
                        planned_first_wp_row, actuals_first_wp_row,
                        actuals_layout)
    summary_last_wp_row = summary_first_wp_row + num_wps - 1
    totals_row = summary_last_wp_row + 1
    _write_totals_row(summary, totals_row, num_years, forecast_col,
                      total_group_start, summary_first_wp_row,
                      summary_last_wp_row)
    _apply_year_variance_cf(summary, num_years, summary_first_wp_row,
                            summary_last_wp_row)
    _apply_total_variance_cf(summary, total_group_start, summary_first_wp_row,
                             summary_last_wp_row)
    _apply_partial_coverage_cf(summary, cal_groups, summary_first_wp_row,
                               summary_last_wp_row)
    _apply_total_forecast_cf(summary, forecast_col, total_group_start,
                             summary_first_wp_row, summary_last_wp_row)

    r = _write_budget_overview(summary, totals_row, val_col_sum,
                               actuals_total_row)

    r = _write_rp_headers(summary, r, reporting_periods)
    num_rps = len(reporting_periods)
    rp_total_group_start = _RP_GROUP_START + num_rps * _COLS_PER_RP
    r = _write_rp_subheaders(summary, r, num_rps, rp_total_group_start)
    rp_first_wp_row = r
    _write_rp_wp_data_rows(summary, r, work_packages, reporting_periods,
                           rp_total_group_start, config_first_wp_row,
                           planned_first_wp_row, actuals_first_wp_row)
    rp_last_wp_row = rp_first_wp_row + num_wps - 1
    rp_totals_row = rp_last_wp_row + 1
    _write_rp_totals_row(summary, rp_totals_row, num_rps, rp_total_group_start,
                         rp_first_wp_row, rp_last_wp_row)
    _apply_rp_variance_cf(summary, num_rps, rp_total_group_start,
                          rp_first_wp_row, rp_last_wp_row)

    _write_form_c(summary, rp_totals_row + 2, rp_totals_row, num_rps,
                  rp_total_group_start)

    summary.freeze_panes = f"C{summary_first_wp_row}"
    summary.column_dimensions["B"].width = 25

    return {
        "first_wp_row": summary_first_wp_row,
        "last_wp_row": summary_last_wp_row,
        "total_row": totals_row,
    }


# ---------------------------------------------------------------------------
# Helpers — annual budget section
# ---------------------------------------------------------------------------


def _write_title(ws) -> int:
    """Write the top title banner. Returns the next row to write at."""
    r = 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, "ANNUAL BUDGET SUMMARY", font=TITLE_FONT, border=False)
    return r + 2


def _write_year_headers(ws, r: int, cal_groups, project_start, num_years: int,
                        forecast_col: int, total_group_start: int) -> int:
    """Write the calendar-year banner row + Forecast EUR banner + PROJECT TOTAL banner."""
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "", fill=LIGHT_BLUE_FILL)
    for y, (group_start, _span) in enumerate(cal_groups):
        start_col = _YEAR_GROUP_START + y * _COLS_PER_YEAR
        end_col = start_col + _COLS_PER_YEAR - 1
        cal_date = project_start + relativedelta(months=group_start)
        ws.merge_cells(start_row=r, start_column=start_col, end_row=r, end_column=end_col)
        style_cell(ws, r, start_col, str(cal_date.year), font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(start_col + 1, end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)

    # Single project-end Forecast EUR column banner — uses LIGHT_GREEN_FILL
    # to flag it as a project-wide rollup (not a per-year metric like the
    # blue year band).
    style_cell(ws, r, forecast_col, "FORECAST", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL, alignment=CENTER)

    ws.merge_cells(start_row=r, start_column=total_group_start, end_row=r, end_column=total_group_start + _COLS_PER_YEAR - 1)
    style_cell(ws, r, total_group_start, font=Font(name="Arial", bold=True, size=10, color="FFFFFF"), fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(total_group_start + 1, total_group_start + _COLS_PER_YEAR):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    return r + 1


def _write_year_subheaders(ws, r: int, num_years: int, forecast_col: int,
                           total_group_start: int) -> int:
    """Write the per-year sub-headers (Plan FTE / Plan € / Act € / Var €) +
    project-end Forecast €."""
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    year_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for y in range(num_years):
        for si, sh in enumerate(year_sub):
            c = _YEAR_GROUP_START + y * _COLS_PER_YEAR + si
            style_cell(ws, r, c, sh, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Project-end Forecast EUR sub-header (single column).
    style_cell(ws, r, forecast_col, "Forecast \u20ac", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL, alignment=CENTER)

    total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for si, sh in enumerate(total_sub):
        c = total_group_start + si
        style_cell(ws, r, c, sh, font=HEADER_FONT, fill=LIGHT_GREEN_FILL, alignment=CENTER)
    return r + 1


def _write_wp_data_rows(ws, first_row: int, work_packages, cal_groups,
                        num_years: int, forecast_col: int,
                        total_group_start: int,
                        config_first_wp_row: int, planned_first_wp_row: int,
                        actuals_first_wp_row: int,
                        actuals_layout: dict) -> None:
    """Write per-WP rows: cross-sheet refs + per-year columns + project-end
    Forecast EUR + project totals."""
    month_col_start = MONTH_COL_START
    posted_idx_l = col_letter(actuals_layout["posted_idx_col"])
    posted_value_l = col_letter(actuals_layout["posted_value_col"])
    slope_l = col_letter(actuals_layout["slope_col"])
    for i, _wp in enumerate(work_packages):
        row = first_row + i
        config_row = config_first_wp_row + i
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        for y, (group_start, span) in enumerate(cal_groups):
            first_col_l = col_letter(month_col_start + group_start)
            last_col_l = col_letter(month_col_start + group_start + span - 1)

            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            plan_wp_row = planned_first_wp_row + i
            actuals_wp_row = actuals_first_wp_row + i

            # Plan FTE (cross-sheet ref to Plan Planned block)
            style_cell(ws, row, sc, font=GREEN_XREF, fmt="0.00").value = (
                f"=SUM({SHEET_PLAN}!{first_col_l}{plan_wp_row}:{last_col_l}{plan_wp_row})"
            )
            # Plan EUR = Plan FTE * per-WP rate from Config col D
            style_cell(ws, row, sc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={col_letter(sc)}{row}*{SHEET_CONFIG}!$D${config_row}"
            )
            # Actual EUR (cross-sheet ref to Actuals sheet)
            style_cell(ws, row, sc + 2, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_ACTUALS}!{first_col_l}{actuals_wp_row}:{last_col_l}{actuals_wp_row})"
            )
            # Variance EUR = Plan EUR - Actual EUR (offset +3, was +4 before
            # the per-year Forecast EUR slot retired in Phase 2.3).
            style_cell(ws, row, sc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
            )

        # Project-end Forecast EUR \u2014 cross-sheet projection from Actuals
        # posted-cumulative cells. When posted_idx is 0 (no posting yet) the
        # formula returns the empty-string sentinel; otherwise it
        # extrapolates posted_value out to project_months at the per-WP
        # slope. MAX(0, ...) clamps a negative slope projection.
        # No ROUND() \u2014 leaf forecast doesn't compound through SUMPRODUCT.
        act_row = actuals_first_wp_row + i
        forecast_formula = (
            f'=IF({SHEET_ACTUALS}!{posted_idx_l}{act_row}=0,"",'
            f'MAX(0,{SHEET_ACTUALS}!{posted_value_l}{act_row}'
            f'+{SHEET_ACTUALS}!{slope_l}{act_row}'
            f'*(project_months-{SHEET_ACTUALS}!{posted_idx_l}{act_row})))'
        )
        style_cell(ws, row, forecast_col, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = forecast_formula

        # Project totals — sum each year-group column across years.
        # `year_sum` closes over `row`; redefined per WP iteration by design.
        # Don't hoist it out of the loop — cell refs must track the current row.
        tc = total_group_start
        def year_sum(offset: int) -> str:
            cols = [col_letter(_YEAR_GROUP_START + y * _COLS_PER_YEAR + offset) for y in range(num_years)]
            return "+".join(f"{cl}{row}" for cl in cols)

        style_cell(ws, row, tc, font=BLACK_FORMULA, fmt="0.00").value = f"={year_sum(0)}"
        style_cell(ws, row, tc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={year_sum(1)}"
        style_cell(ws, row, tc + 2, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={year_sum(2)}"
        # Total variance EUR = Plan EUR total - Actual EUR total (offset +3,
        # mirrors the year band after Phase 2.3's Forecast slot retirement).
        style_cell(ws, row, tc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
            f"={col_letter(tc + 1)}{row}-{col_letter(tc + 2)}{row}"
        )


def _write_totals_row(ws, totals_row: int, num_years: int, forecast_col: int,
                      total_group_start: int, first_wp_row: int,
                      last_wp_row: int) -> None:
    """Write the per-column SUM totals row across all year-groups + project-end
    Forecast EUR + PROJECT TOTAL."""
    style_cell(ws, totals_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, totals_row, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for y in range(num_years):
        for si in range(_COLS_PER_YEAR):
            c = _YEAR_GROUP_START + y * _COLS_PER_YEAR + si
            fmt = "0.00" if si == 0 else "\u20ac#,##0"
            col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
            style_cell(ws, totals_row, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

    # Project-end Forecast EUR TOTAL (single column, separate from PROJECT TOTAL).
    fc_range = f"{col_letter(forecast_col)}{first_wp_row}:{col_letter(forecast_col)}{last_wp_row}"
    style_cell(ws, totals_row, forecast_col, font=BLACK_FORMULA,
               fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0").value = f"=SUM({fc_range})"

    for si in range(_COLS_PER_YEAR):
        c = total_group_start + si
        fmt = "0.00" if si == 0 else "\u20ac#,##0"
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, totals_row, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"


def _apply_year_variance_cf(ws, num_years: int, first_wp_row: int,
                            last_wp_row: int) -> None:
    """CF: red on negative Variance EUR (per-year columns).

    Variance is now at offset +3 (was +4 before Phase 2.3 retired the
    per-year Forecast EUR slot).
    """
    for y in range(num_years):
        var_col = _YEAR_GROUP_START + y * _COLS_PER_YEAR + 3
        var_col_l = col_letter(var_col)
        var_range = f"{var_col_l}{first_wp_row}:{var_col_l}{last_wp_row}"
        ws.conditional_formatting.add(var_range, FormulaRule(
            formula=[f"{var_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
        ))


def _apply_total_variance_cf(ws, total_group_start: int, first_wp_row: int,
                             last_wp_row: int) -> None:
    """CF: red on negative Variance EUR (PROJECT TOTAL column).

    Variance is now at offset +3 (was +4 before Phase 2.3).
    """
    total_var_col = total_group_start + 3
    total_var_col_l = col_letter(total_var_col)
    total_var_range = f"{total_var_col_l}{first_wp_row}:{total_var_col_l}{last_wp_row}"
    ws.conditional_formatting.add(total_var_range, FormulaRule(
        formula=[f"{total_var_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))


def _apply_partial_coverage_cf(ws, cal_groups, first_wp_row: int,
                               last_wp_row: int) -> None:
    """CF: grey on per-year partial-coverage Act/Var € columns.

    Act €: year not fully covered by eur_actuals_through_month.
    Var €: stricter — either EUR or hours coverage short of the year end.
    Skips PROJECT TOTAL (summarises across years; no single coverage gate).
    Phase 2.3 retired the per-year Fcst € grey rule (no per-year Forecast
    column anymore — replaced by a single project-end Forecast EUR column).
    """
    for y, (group_start, span) in enumerate(cal_groups):
        year_end_pm = group_start + span
        sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
        act_col_l = col_letter(sc + 2)
        # Variance is at offset +3 (was +4 before Phase 2.3 retired the
        # per-year Forecast slot). The per-year Fcst € grey rule retired
        # along with the column.
        var_col_l = col_letter(sc + 3)

        act_range = f"{act_col_l}{first_wp_row}:{act_col_l}{last_wp_row}"
        ws.conditional_formatting.add(act_range, FormulaRule(
            formula=[f"eur_actuals_through_month<{year_end_pm}"],
            fill=CF_GREY_FILL,
        ))

        var_range_grey = f"{var_col_l}{first_wp_row}:{var_col_l}{last_wp_row}"
        ws.conditional_formatting.add(var_range_grey, FormulaRule(
            formula=[
                f"MIN(eur_actuals_through_month,hours_actuals_through_month)"
                f"<{year_end_pm}"
            ],
            fill=CF_GREY_FILL,
        ))


def _apply_total_forecast_cf(ws, forecast_col: int, total_group_start: int,
                             first_wp_row: int, last_wp_row: int) -> None:
    """CF: red on project-end Forecast EUR < 0.8 * Plan EUR (project total).

    Compares per-row forecast (single column at ``forecast_col``) to
    per-row Plan EUR project total at ``total_group_start + 1``. The
    ``<>""`` guard skips the empty-string sentinel emitted by the
    forecast formula when posted_idx is 0 (no postings yet) -- without
    it, the comparison ``""<0.8*X`` fires the red rule for every WP.
    """
    fc = col_letter(forecast_col)
    plan_total_l = col_letter(total_group_start + 1)
    fc_range = f"{fc}{first_wp_row}:{fc}{last_wp_row}"
    ws.conditional_formatting.add(fc_range, FormulaRule(
        formula=[
            f'AND({fc}{first_wp_row}<>"",'
            f"{fc}{first_wp_row}<0.8*{plan_total_l}{first_wp_row})"
        ],
        fill=CF_RED_FILL,
    ))


# ---------------------------------------------------------------------------
# Helpers — Budget Overview section
# ---------------------------------------------------------------------------


def _write_budget_overview(ws, totals_row: int, val_col_sum: int,
                           actuals_total_row: int) -> int:
    """Write the BUDGET OVERVIEW block + its CF rules. Returns next-r cursor."""
    r = totals_row + 2
    write_section_header(ws, r, "BUDGET OVERVIEW", merge_end=2)
    r += 2

    # Personnel Actual — total EUR from Actuals sheet
    actual_row = r
    style_cell(ws, r, 1, "Personnel Actual", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, font=GREEN_XREF, fmt="\u20ac#,##0",
               border=False).value = (
        f"={SHEET_ACTUALS}!{col_letter(val_col_sum)}{actuals_total_row}"
    )
    r += 1

    # Total Claimable (+25%)
    style_cell(ws, r, 1, "Total Claimable (+25%)", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, font=BLACK_FORMULA, fmt="\u20ac#,##0",
               border=False).value = f"=B{actual_row}*1.25"
    r += 1

    # Personnel Ceiling
    ceiling_row = r
    style_cell(ws, r, 1, "Personnel Ceiling", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, font=GREEN_XREF, fmt="\u20ac#,##0",
               border=False).value = "=partner_ceiling"
    r += 1

    # Remaining = Ceiling - Personnel Actual
    remaining_row = r
    style_cell(ws, r, 1, "Remaining", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, font=BLACK_FORMULA, fmt="\u20ac#,##0",
               border=False).value = f"=B{ceiling_row}-B{actual_row}"
    r += 1

    # % Consumed
    style_cell(ws, r, 1, "% Consumed", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, font=BLACK_FORMULA, fmt="0.0%",
               border=False).value = (
        f"=IF(B{ceiling_row}=0,0,B{actual_row}/B{ceiling_row})"
    )

    # Conditional formatting on Remaining — red first = highest priority
    # Red if negative
    ws.conditional_formatting.add(
        f"B{remaining_row}", FormulaRule(
            formula=[f"B{remaining_row}<0"], fill=CF_RED_FILL,
        ))
    # Amber if < 10% of ceiling
    ws.conditional_formatting.add(
        f"B{remaining_row}", FormulaRule(
            formula=[f"B{remaining_row}<B{ceiling_row}*0.1"], fill=CF_AMBER_FILL,
        ))

    return r + 2


# ---------------------------------------------------------------------------
# Helpers — Reporting Period Summary section
# ---------------------------------------------------------------------------


def _write_rp_headers(ws, r: int, reporting_periods) -> int:
    """Write the REPORTING PERIOD SUMMARY section header + per-RP banner row."""
    write_section_header(ws, r, "REPORTING PERIOD SUMMARY", merge_end=2)
    r += 2

    num_rps = len(reporting_periods)
    rp_total_group_start = _RP_GROUP_START + num_rps * _COLS_PER_RP

    # RP headers row
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "", fill=LIGHT_BLUE_FILL)
    for rp_idx, rp in enumerate(reporting_periods):
        start_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP
        end_col = start_col + _COLS_PER_RP - 1
        label = f"{rp['id']} (M{rp['start']}\u2013M{rp['end']})"
        ws.merge_cells(start_row=r, start_column=start_col,
                       end_row=r, end_column=end_col)
        style_cell(ws, r, start_col, label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(start_col + 1, end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)

    ws.merge_cells(start_row=r, start_column=rp_total_group_start,
                   end_row=r, end_column=rp_total_group_start + _COLS_PER_RP - 1)
    style_cell(ws, r, rp_total_group_start,
               font=Font(name="Arial", bold=True, size=10, color="FFFFFF"),
               fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(rp_total_group_start + 1, rp_total_group_start + _COLS_PER_RP):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    return r + 1


def _write_rp_subheaders(ws, r: int, num_rps: int,
                         rp_total_group_start: int) -> int:
    """Write the RP sub-headers (Plan FTE / Plan € / Act € / Var €)."""
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    rp_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for rp_idx in range(num_rps):
        for si, sh in enumerate(rp_sub):
            c = _RP_GROUP_START + rp_idx * _COLS_PER_RP + si
            style_cell(ws, r, c, sh, font=HEADER_FONT,
                       fill=LIGHT_BLUE_FILL, alignment=CENTER)
    rp_total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for si, sh in enumerate(rp_total_sub):
        c = rp_total_group_start + si
        style_cell(ws, r, c, sh, font=HEADER_FONT,
                   fill=LIGHT_GREEN_FILL, alignment=CENTER)
    return r + 1


def _write_rp_wp_data_rows(ws, first_row: int, work_packages,
                           reporting_periods, rp_total_group_start: int,
                           config_first_wp_row: int, planned_first_wp_row: int,
                           actuals_first_wp_row: int) -> None:
    """Write per-WP rows in the RP section: cross-sheet refs + per-RP cols + totals."""
    month_col_start = MONTH_COL_START
    num_rps = len(reporting_periods)
    for i, _wp in enumerate(work_packages):
        row = first_row + i
        config_row = config_first_wp_row + i
        style_cell(ws, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(ws, row, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        for rp_idx, rp in enumerate(reporting_periods):
            first_col_l = col_letter(month_col_start + rp["start"] - 1)
            last_col_l = col_letter(month_col_start + rp["end"] - 1)

            sc = _RP_GROUP_START + rp_idx * _COLS_PER_RP
            plan_wp_row = planned_first_wp_row + i
            actuals_wp_row = actuals_first_wp_row + i

            # Plan FTE
            style_cell(ws, row, sc, font=GREEN_XREF, fmt="0.00").value = (
                f"=SUM({SHEET_PLAN}!{first_col_l}{plan_wp_row}:"
                f"{last_col_l}{plan_wp_row})"
            )
            # Plan EUR = Plan FTE * per-WP rate from Config col D
            style_cell(ws, row, sc + 1, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0").value = (
                f"={col_letter(sc)}{row}*{SHEET_CONFIG}!$D${config_row}"
            )
            # Actual EUR
            style_cell(ws, row, sc + 2, font=GREEN_XREF,
                       fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_ACTUALS}!{first_col_l}{actuals_wp_row}:"
                f"{last_col_l}{actuals_wp_row})"
            )
            # Variance EUR = Plan EUR - Actual EUR
            style_cell(ws, row, sc + 3, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0").value = (
                f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
            )

        # RP Project totals (sum across RP groups)
        tc = rp_total_group_start
        rp_plan_cols = [col_letter(_RP_GROUP_START + rp_idx * _COLS_PER_RP)
                        for rp_idx in range(num_rps)]
        style_cell(ws, row, tc, font=BLACK_FORMULA, fmt="0.00").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_plan_cols)
        )
        rp_plan_eur_cols = [col_letter(_RP_GROUP_START + rp_idx * _COLS_PER_RP + 1)
                            for rp_idx in range(num_rps)]
        style_cell(ws, row, tc + 1, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_plan_eur_cols)
        )
        rp_act_eur_cols = [col_letter(_RP_GROUP_START + rp_idx * _COLS_PER_RP + 2)
                           for rp_idx in range(num_rps)]
        style_cell(ws, row, tc + 2, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            "=" + "+".join(f"{cl}{row}" for cl in rp_act_eur_cols)
        )
        style_cell(ws, row, tc + 3, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(tc + 1)}{row}-{col_letter(tc + 2)}{row}"
        )


def _write_rp_totals_row(ws, rp_totals_row: int, num_rps: int,
                         rp_total_group_start: int, first_wp_row: int,
                         last_wp_row: int) -> None:
    """Write the per-column SUM totals row for the RP section."""
    style_cell(ws, rp_totals_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, rp_totals_row, 2, "TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    for rp_idx in range(num_rps):
        for si in range(_COLS_PER_RP):
            c = _RP_GROUP_START + rp_idx * _COLS_PER_RP + si
            fmt = "0.00" if si == 0 else "\u20ac#,##0"
            col_range = (f"{col_letter(c)}{first_wp_row}:"
                         f"{col_letter(c)}{last_wp_row}")
            style_cell(ws, rp_totals_row, c, font=BLACK_FORMULA,
                       fill=LIGHT_GREEN_FILL, fmt=fmt).value = (
                f"=SUM({col_range})"
            )
    for si in range(_COLS_PER_RP):
        c = rp_total_group_start + si
        fmt = "0.00" if si == 0 else "\u20ac#,##0"
        col_range = (f"{col_letter(c)}{first_wp_row}:"
                     f"{col_letter(c)}{last_wp_row}")
        style_cell(ws, rp_totals_row, c, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL, fmt=fmt).value = (
            f"=SUM({col_range})"
        )


def _apply_rp_variance_cf(ws, num_rps: int, rp_total_group_start: int,
                          first_wp_row: int, last_wp_row: int) -> None:
    """CF: red on negative RP Variance EUR (per-RP + PROJECT TOTAL)."""
    # CF: red on negative RP Variance EUR
    for rp_idx in range(num_rps):
        var_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 3
        var_col_l = col_letter(var_col)
        var_range = (f"{var_col_l}{first_wp_row}:"
                     f"{var_col_l}{last_wp_row}")
        ws.conditional_formatting.add(var_range, FormulaRule(
            formula=[f"{var_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
        ))
    # CF: red on negative PROJECT TOTAL Variance EUR
    rp_total_var_col = rp_total_group_start + 3
    rp_total_var_col_l = col_letter(rp_total_var_col)
    rp_total_var_range = (f"{rp_total_var_col_l}{first_wp_row}:"
                          f"{rp_total_var_col_l}{last_wp_row}")
    ws.conditional_formatting.add(rp_total_var_range, FormulaRule(
        formula=[f"{rp_total_var_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))


# ---------------------------------------------------------------------------
# Helpers — Form C preparation section
# ---------------------------------------------------------------------------


def _write_form_c(ws, r: int, rp_totals_row: int, num_rps: int,
                  rp_total_group_start: int) -> None:
    """Write the FORM C PREPARATION block: A. Personnel, E. Indirect, Total Eligible,
    Cum. Eligible, CFS Required? — plus CFS amber CF.

    ``r`` is the first writable row (caller passes ``rp_totals_row + 2`` for the
    separator). ``rp_totals_row`` supplies the cross-references for A. Personnel.
    """
    # Row labels and formulas use only the Act EUR sub-column (offset +2) per RP

    # A. Personnel EUR — reference TOTAL Act EUR for each RP
    personnel_fc_row = r
    style_cell(ws, r, 1, "A. Personnel EUR", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 2
        style_cell(ws, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{rp_totals_row}"
        )
    # Project total
    total_act_col = rp_total_group_start + 2
    style_cell(ws, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{rp_totals_row}"
    )
    r += 1

    # E. Indirect (25%) — 25% of Personnel EUR
    indirect_row = r
    style_cell(ws, r, 1, "E. Indirect (25%)", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 2
        style_cell(ws, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{personnel_fc_row}*0.25"
        )
    style_cell(ws, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{personnel_fc_row}*0.25"
    )
    r += 1

    # Total Eligible = Personnel + Indirect
    eligible_row = r
    style_cell(ws, r, 1, "Total Eligible", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 2
        style_cell(ws, r, act_col, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0", border=False).value = (
            f"={col_letter(act_col)}{personnel_fc_row}"
            f"+{col_letter(act_col)}{indirect_row}"
        )
    style_cell(ws, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(total_act_col)}{personnel_fc_row}"
        f"+{col_letter(total_act_col)}{indirect_row}"
    )
    r += 1

    # Cumulative Eligible — running sum across RPs
    cum_row = r
    style_cell(ws, r, 1, "Cum. Eligible", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 2
        if rp_idx == 0:
            style_cell(ws, r, act_col, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0", border=False).value = (
                f"={col_letter(act_col)}{eligible_row}"
            )
        else:
            prev_act_col = _RP_GROUP_START + (rp_idx - 1) * _COLS_PER_RP + 2
            style_cell(ws, r, act_col, font=BLACK_FORMULA,
                       fmt="\u20ac#,##0", border=False).value = (
                f"={col_letter(prev_act_col)}{cum_row}"
                f"+{col_letter(act_col)}{eligible_row}"
            )
    # Project total cumulative = last RP's cumulative value
    last_rp_cum_col = _RP_GROUP_START + (num_rps - 1) * _COLS_PER_RP + 2
    style_cell(ws, r, total_act_col, font=BLACK_FORMULA,
               fmt="\u20ac#,##0", border=False).value = (
        f"={col_letter(last_rp_cum_col)}{cum_row}"
    )
    r += 1

    # CFS Required? — flag when cumulative > EUR 430k
    CFS_THRESHOLD = 430_000  # EUR — audit certificate required above this
    cfs_row = r
    style_cell(ws, r, 1, "CFS Required?", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "", border=False)
    for rp_idx in range(num_rps):
        act_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 2
        style_cell(ws, r, act_col, font=BLACK_FORMULA,
                   border=False).value = (
            f'=IF({col_letter(act_col)}{cum_row}>{CFS_THRESHOLD},"YES","No")'
        )
    style_cell(ws, r, total_act_col, font=BLACK_FORMULA,
               border=False).value = (
        f'=IF({col_letter(total_act_col)}{cum_row}>{CFS_THRESHOLD},"YES","No")'
    )

    # CF on CFS Required: amber fill on "YES"
    for rp_idx in range(num_rps):
        act_col = _RP_GROUP_START + rp_idx * _COLS_PER_RP + 2
        cell_ref = f"{col_letter(act_col)}{cfs_row}"
        ws.conditional_formatting.add(cell_ref, FormulaRule(
            formula=[f'{cell_ref}="YES"'], fill=CF_AMBER_FILL,
        ))
    total_cfs_ref = f"{col_letter(total_act_col)}{cfs_row}"
    ws.conditional_formatting.add(total_cfs_ref, FormulaRule(
        formula=[f'{total_cfs_ref}="YES"'], fill=CF_AMBER_FILL,
    ))
