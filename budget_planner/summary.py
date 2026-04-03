"""Summary sheet builder — annual budget overview with per-year and project-total columns."""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.layout import (
    MONTH_COL_START, SHEET_SUMMARY, SHEET_PLAN, SHEET_CONFIG, SHEET_ACTUALS,
    val_columns, _calendar_year_groups,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT, TITLE_FONT,
    SECTION_FONT, SECTION_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    AMBER_FILL, RED_FILL, CENTER, style_cell, col_letter,
)


def build_summary_sheet(wb, layout, actuals_layout, *, project_params=None, work_packages=None):
    """Create the Summary sheet with annual budget view and budget overview.

    Parameters
    ----------
    wb : openpyxl.Workbook
    layout : dict
        From build_plan_sheet(): config_first_wp_row, planned_first_wp_row.
    actuals_layout : dict
        From build_actuals_sheet(): first_wp_row, total_row.
    project_params : dict, optional
        Keys: title, start, months, rate. Lazy-imports from config if None.
    work_packages : list[dict], optional
        List of WP dicts. Lazy-imports from config if None.
    """
    if project_params is None or work_packages is None:
        from budget_planner import config
        if project_params is None:
            project_params = {
                "title": config.PROJECT_TITLE,
                "start": config.PROJECT_START,
                "months": config.PROJECT_MONTHS,
                "rate": config.DEFAULT_RATE,
            }
        if work_packages is None:
            work_packages = config.WORK_PACKAGES

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
    COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Actual EUR, Variance EUR

    # Project total group
    TOTAL_GROUP_START = YEAR_GROUP_START + num_years * COLS_PER_YEAR
    TOTAL_COLS = 4  # Plan FTE, Plan EUR, Actual EUR, Variance EUR

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

    summary.merge_cells(start_row=r, start_column=TOTAL_GROUP_START, end_row=r, end_column=TOTAL_GROUP_START + TOTAL_COLS - 1)
    style_cell(summary, r, TOTAL_GROUP_START, font=Font(name="Arial", bold=True, size=10, color="FFFFFF"), fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(TOTAL_GROUP_START + 1, TOTAL_GROUP_START + TOTAL_COLS):
        style_cell(summary, r, c2, fill=SECTION_FILL)
    r += 1

    # Sub-headers row
    style_cell(summary, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(summary, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    year_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
    for y in range(num_years):
        for si, sh in enumerate(year_sub):
            c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
            style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    total_sub = ["Plan FTE", "Plan \u20ac", "Act \u20ac", "Var \u20ac"]
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
            # Plan EUR = Plan FTE * fte_rate
            style_cell(summary, row, sc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={col_letter(sc)}{row}*fte_rate"
            )
            # Actual EUR (cross-sheet ref to Actuals sheet)
            style_cell(summary, row, sc + 2, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_ACTUALS}!{first_col_l}{actuals_wp_row}:{last_col_l}{actuals_wp_row})"
            )
            # Variance EUR = Plan EUR - Actual EUR
            style_cell(summary, row, sc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={col_letter(sc + 1)}{row}-{col_letter(sc + 2)}{row}"
            )

        # Project totals
        tc = TOTAL_GROUP_START
        # Sum planned FTE across years
        plan_year_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR) for y in range(num_years)]
        plan_sum = "+".join([f"{cl}{row}" for cl in plan_year_cols])
        style_cell(summary, row, tc, font=BLACK_FORMULA, fmt="0.00").value = f"={plan_sum}"
        # Sum planned EUR across years
        plan_eur_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR + 1) for y in range(num_years)]
        plan_eur_sum = "+".join([f"{cl}{row}" for cl in plan_eur_cols])
        style_cell(summary, row, tc + 1, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={plan_eur_sum}"
        # Sum actual EUR across years
        act_eur_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR + 2) for y in range(num_years)]
        act_eur_sum = "+".join([f"{cl}{row}" for cl in act_eur_cols])
        style_cell(summary, row, tc + 2, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={act_eur_sum}"
        # Total variance EUR
        style_cell(summary, row, tc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
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

    for si in range(TOTAL_COLS):
        c = TOTAL_GROUP_START + si
        fmt = "0.00" if si == 0 else "\u20ac#,##0"
        col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
        style_cell(summary, totals_row, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

    # Conditional formatting: red on negative Variance EUR (year columns)
    for y in range(num_years):
        var_col = YEAR_GROUP_START + y * COLS_PER_YEAR + 3
        var_col_l = col_letter(var_col)
        var_range = f"{var_col_l}{SUMMARY_FIRST_WP_ROW}:{var_col_l}{SUMMARY_LAST_WP_ROW}"
        summary.conditional_formatting.add(var_range, FormulaRule(
            formula=[f"{var_col_l}{SUMMARY_FIRST_WP_ROW}<0"], fill=RED_FILL,
        ))

    # Conditional formatting: red on negative Variance EUR (project total)
    total_var_col = TOTAL_GROUP_START + 3
    total_var_col_l = col_letter(total_var_col)
    total_var_range = f"{total_var_col_l}{SUMMARY_FIRST_WP_ROW}:{total_var_col_l}{SUMMARY_LAST_WP_ROW}"
    summary.conditional_formatting.add(total_var_range, FormulaRule(
        formula=[f"{total_var_col_l}{SUMMARY_FIRST_WP_ROW}<0"], fill=RED_FILL,
    ))

    # ── BUDGET OVERVIEW ──
    r = totals_row + 2
    summary.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(summary, r, 1, "BUDGET OVERVIEW", font=SECTION_FONT,
               fill=SECTION_FILL, alignment=CENTER)
    style_cell(summary, r, 2, fill=SECTION_FILL)
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

    # Conditional formatting on Remaining (amber before red per openpyxl convention)
    # Amber if < 10% of ceiling
    summary.conditional_formatting.add(
        f"B{remaining_row}", FormulaRule(
            formula=[f"B{remaining_row}<B{ceiling_row}*0.1"], fill=AMBER_FILL,
        ))
    # Red if negative
    summary.conditional_formatting.add(
        f"B{remaining_row}", FormulaRule(
            formula=[f"B{remaining_row}<0"], fill=RED_FILL,
        ))

    summary.freeze_panes = f"C{SUMMARY_FIRST_WP_ROW}"
