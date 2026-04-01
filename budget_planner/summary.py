"""Summary sheet builder — annual budget overview with per-year and project-total columns."""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT, TITLE_FONT,
    SECTION_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    RED_FILL, CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_START, PROJECT_MONTHS, WORK_PACKAGES, NUM_WPS,
    MONTH_COL_START,
)


def build_summary_sheet(wb, layout):
    """Create the Summary sheet with annual and project-total budget view.

    Parameters
    ----------
    wb : openpyxl.Workbook
        The workbook to add the Summary sheet to.
    layout : dict
        Must contain keys returned by build_plan_sheet():
        - config_first_wp_row
        - planned_first_wp_row
        - actual_first_wp_row
    """
    config_first_wp_row = layout["config_first_wp_row"]
    planned_first_wp_row = layout["planned_first_wp_row"]
    actual_first_wp_row = layout["actual_first_wp_row"]

    summary = wb.create_sheet("Summary")
    summary.sheet_properties.tabColor = "008000"

    r = 1
    summary.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(summary, r, 1, "ANNUAL BUDGET SUMMARY", font=TITLE_FONT, border=False)
    r += 2

    # Calculate years
    num_years = (PROJECT_MONTHS + 11) // 12

    # Year group headers
    YEAR_GROUP_START = 3
    COLS_PER_YEAR = 5  # Planned FTE, Actual FTE, Variance FTE, Planned EUR, Actual EUR

    # Project total group
    TOTAL_GROUP_START = YEAR_GROUP_START + num_years * COLS_PER_YEAR
    TOTAL_COLS = 7  # Total Planned FTE, Total Actual FTE, Total Variance, Total Planned EUR, Total Actual EUR, Remaining EUR, % Consumed

    # Year headers row
    style_cell(summary, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(summary, r, 2, "", fill=LIGHT_BLUE_FILL)
    for y in range(num_years):
        start_col = YEAR_GROUP_START + y * COLS_PER_YEAR
        end_col = start_col + COLS_PER_YEAR - 1
        cal_year = PROJECT_START.year + y
        summary.merge_cells(start_row=r, start_column=start_col, end_row=r, end_column=end_col)
        style_cell(summary, r, start_col, f"Year {y+1} ({cal_year})", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(start_col + 1, end_col + 1):
            style_cell(summary, r, c2, fill=LIGHT_BLUE_FILL)

    summary.merge_cells(start_row=r, start_column=TOTAL_GROUP_START, end_row=r, end_column=TOTAL_GROUP_START + TOTAL_COLS - 1)
    style_cell(summary, r, TOTAL_GROUP_START, "PROJECT TOTAL", font=HEADER_FONT, fill=SECTION_FILL, alignment=CENTER)
    style_cell(summary, r, TOTAL_GROUP_START, font=Font(name="Arial", bold=True, size=10, color="FFFFFF"), fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
    for c2 in range(TOTAL_GROUP_START + 1, TOTAL_GROUP_START + TOTAL_COLS):
        style_cell(summary, r, c2, fill=SECTION_FILL)
    r += 1

    # Sub-headers row
    style_cell(summary, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(summary, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    year_sub = ["Plan FTE", "Act FTE", "Var FTE", "Plan \u20ac", "Act \u20ac"]
    for y in range(num_years):
        for si, sh in enumerate(year_sub):
            c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
            style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    total_sub = ["Plan FTE", "Act FTE", "Var FTE", "Plan \u20ac", "Act \u20ac", "Rem \u20ac", "% Used"]
    for si, sh in enumerate(total_sub):
        c = TOTAL_GROUP_START + si
        style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_GREEN_FILL, alignment=CENTER)
    r += 1

    SUMMARY_FIRST_WP_ROW = r

    # WP data rows
    for i, wp in enumerate(WORK_PACKAGES):
        row = r + i
        config_row = config_first_wp_row + i
        style_cell(summary, row, 1, font=GREEN_XREF).value = f"=Plan!A{config_row}"
        style_cell(summary, row, 2, font=GREEN_XREF).value = f"=Plan!B{config_row}"

        plan_wp_row = planned_first_wp_row + i
        act_wp_row = actual_first_wp_row + i

        for y in range(num_years):
            m_start = y * 12
            m_end = min((y + 1) * 12, PROJECT_MONTHS) - 1
            first_col_l = col_letter(MONTH_COL_START + m_start)
            last_col_l = col_letter(MONTH_COL_START + m_end)

            sc = YEAR_GROUP_START + y * COLS_PER_YEAR
            # Planned FTE for year
            style_cell(summary, row, sc, font=GREEN_XREF, fmt="0.00").value = f"=SUM(Plan!{first_col_l}{plan_wp_row}:{last_col_l}{plan_wp_row})"
            # Actual FTE for year
            style_cell(summary, row, sc + 1, font=GREEN_XREF, fmt="0.00").value = f"=SUM(Plan!{first_col_l}{act_wp_row}:{last_col_l}{act_wp_row})"
            # Variance
            style_cell(summary, row, sc + 2, font=BLACK_FORMULA, fmt="0.00").value = f"={col_letter(sc)}{row}-{col_letter(sc+1)}{row}"
            # Planned EUR
            style_cell(summary, row, sc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={col_letter(sc)}{row}*Plan!$E$2"
            # Actual EUR
            style_cell(summary, row, sc + 4, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={col_letter(sc+1)}{row}*Plan!$E$2"

        # Project totals
        tc = TOTAL_GROUP_START
        # Sum planned FTE across years
        plan_year_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR) for y in range(num_years)]
        plan_sum = "+".join([f"{cl}{row}" for cl in plan_year_cols])
        style_cell(summary, row, tc, font=BLACK_FORMULA, fmt="0.00").value = f"={plan_sum}"
        # Sum actual FTE
        act_year_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR + 1) for y in range(num_years)]
        act_sum = "+".join([f"{cl}{row}" for cl in act_year_cols])
        style_cell(summary, row, tc + 1, font=BLACK_FORMULA, fmt="0.00").value = f"={act_sum}"
        # Total variance
        style_cell(summary, row, tc + 2, font=BLACK_FORMULA, fmt="0.00").value = f"={col_letter(tc)}{row}-{col_letter(tc+1)}{row}"
        # Total planned EUR
        style_cell(summary, row, tc + 3, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={col_letter(tc)}{row}*Plan!$E$2"
        # Total actual EUR
        style_cell(summary, row, tc + 4, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"={col_letter(tc+1)}{row}*Plan!$E$2"
        # Remaining EUR
        style_cell(summary, row, tc + 5, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"=Plan!G{config_row}-{col_letter(tc+4)}{row}"
        # % consumed
        style_cell(summary, row, tc + 6, font=BLACK_FORMULA, fmt="0.0%").value = f"=IF(Plan!C{config_row}=0,0,{col_letter(tc+1)}{row}/Plan!C{config_row})"

    SUMMARY_LAST_WP_ROW = SUMMARY_FIRST_WP_ROW + NUM_WPS - 1
    r = SUMMARY_LAST_WP_ROW + 1

    # Totals row
    style_cell(summary, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(summary, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for y in range(num_years):
        for si in range(COLS_PER_YEAR):
            c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
            fmt = "\u20ac#,##0" if si >= 3 else "0.00"
            col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
            style_cell(summary, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

    for si in range(TOTAL_COLS):
        c = TOTAL_GROUP_START + si
        fmt = "0.0%" if si == 6 else ("\u20ac#,##0" if si >= 3 else "0.00")
        col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
        if si == 6:
            style_cell(summary, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = \
                f"=IF({col_letter(TOTAL_GROUP_START)}{r}=0,0,{col_letter(TOTAL_GROUP_START+1)}{r}/{col_letter(TOTAL_GROUP_START)}{r})"
        else:
            style_cell(summary, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

    # Conditional formatting: red on negative remaining EUR
    rem_col = col_letter(TOTAL_GROUP_START + 5)
    rem_range = f"{rem_col}{SUMMARY_FIRST_WP_ROW}:{rem_col}{SUMMARY_LAST_WP_ROW}"
    summary.conditional_formatting.add(rem_range, FormulaRule(
        formula=[f"{rem_col}{SUMMARY_FIRST_WP_ROW}<0"], fill=RED_FILL
    ))

    # Column widths
    summary.freeze_panes = f"C{SUMMARY_FIRST_WP_ROW}"
