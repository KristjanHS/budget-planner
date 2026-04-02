"""Actuals sheet builder — EUR actuals per WP per month from accountant."""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    RED_FILL, AMBER_FILL, CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_START, PROJECT_MONTHS, WORK_PACKAGES, NUM_WPS,
    MONTH_COL_START, MONTH_COL_END,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
    VAL_COL_CUM, VAL_COL_PCT_CEIL, VAL_COL_BREACH,
)


def build_actuals_sheet(wb, layout):
    """Build the Actuals sheet for EUR actuals per WP per month.

    Args:
        wb: The parent Workbook.
        layout: Layout dict from build_plan_sheet(), needs config_first_wp_row.

    Returns:
        Layout dict with keys: header_row, year_row, month_row,
        first_wp_row, last_wp_row, total_row.
    """
    config_first_wp_row = layout["config_first_wp_row"]

    ws = wb.create_sheet("Actuals")
    ws.sheet_properties.tabColor = "0000FF"

    r = 1

    # -- Section header --
    merge_end = VAL_COL_BREACH
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_end)
    style_cell(ws, r, 1, "ACTUAL EUR", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, merge_end + 1):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1
    year_row = r

    # -- Year header row --
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "", fill=LIGHT_BLUE_FILL)
    m = 0
    while m < PROJECT_MONTHS:
        year_start_col = MONTH_COL_START + m
        months_left = PROJECT_MONTHS - m
        span = min(12, months_left)
        year_end_col = year_start_col + span - 1
        cal_date = PROJECT_START + relativedelta(months=m)
        year_label = f"Year {m // 12 + 1} ({cal_date.year})"
        if span > 1:
            ws.merge_cells(start_row=r, start_column=year_start_col,
                           end_row=r, end_column=year_end_col)
        style_cell(ws, r, year_start_col, year_label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(year_start_col + 1, year_end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)
        m += span
    # Validation headers
    style_cell(ws, r, VAL_COL_SUM, "Sum", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_BUDGET, "Ceiling", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_DIFF, "Variance", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_CUM, "Cum. EUR", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_PCT_CEIL, "% Ceiling", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_BREACH, "Breach In", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1
    month_row = r

    # -- Month header row --
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    for m_idx in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m_idx
        cal_date = PROJECT_START + relativedelta(months=m_idx)
        mlabel = f"M{m_idx + 1}\n{cal_date.strftime('%b %y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8),
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_SUM, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_BUDGET, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_DIFF, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_CUM, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_PCT_CEIL, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_BREACH, "", fill=LIGHT_BLUE_FILL)
    r += 1
    first_wp_row = r

    # -- WP rows (blue input EUR cells) --
    for i, wp in enumerate(WORK_PACKAGES):
        row = r + i
        config_row = config_first_wp_row + i
        # Cross-sheet refs to Plan config
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"=Plan!A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"=Plan!B{config_row}"

        for m_idx in range(PROJECT_MONTHS):
            c = MONTH_COL_START + m_idx
            month_num = m_idx + 1
            if wp["start"] <= month_num <= wp["end"]:
                style_cell(ws, row, c, 0, font=BLUE_INPUT, fmt="\u20ac#,##0")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="\u20ac#,##0")

        # Validation: Sum
        sum_range = (f"{col_letter(MONTH_COL_START)}{row}:"
                     f"{col_letter(MONTH_COL_END)}{row}")
        style_cell(ws, row, VAL_COL_SUM, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
        # Validation: Ceiling (from Plan Personnel Ceiling column G)
        style_cell(ws, row, VAL_COL_BUDGET, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = f"=Plan!G{config_row}"
        # Validation: Variance = Ceiling - Sum
        style_cell(ws, row, VAL_COL_DIFF, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"={col_letter(VAL_COL_BUDGET)}{row}"
            f"-{col_letter(VAL_COL_SUM)}{row}"
        )

        # Cumulative EUR (up to current_month)
        ms = col_letter(MONTH_COL_START)
        me = col_letter(MONTH_COL_END)
        style_cell(ws, row, VAL_COL_CUM, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = (
            f"=SUMPRODUCT((COLUMN({ms}{row}:{me}{row})"
            f"-COLUMN(${ms}$1)+1<=current_month)"
            f"*({ms}{row}:{me}{row}))"
        )
        # % Ceiling = Cum. EUR / Ceiling
        cum_l = col_letter(VAL_COL_CUM)
        ceil_l = col_letter(VAL_COL_BUDGET)
        style_cell(ws, row, VAL_COL_PCT_CEIL, font=BLACK_FORMULA,
                   fmt="0.0%").value = f"={cum_l}{row}/{ceil_l}{row}"
        # Breach In = months until ceiling exhausted at current rate
        style_cell(ws, row, VAL_COL_BREACH, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({cum_l}{row}=0,"\u2014",'
            f"({ceil_l}{row}-{cum_l}{row})"
            f"/({cum_l}{row}/current_month))"
        )

    last_wp_row = first_wp_row + NUM_WPS - 1
    r = last_wp_row + 1

    # -- TOTAL row --
    total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for m_idx in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m_idx
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="\u20ac#,##0").value = f"=SUM({col_range})"
    # TOTAL Sum
    sum_range = (f"{col_letter(VAL_COL_SUM)}{first_wp_row}:"
                 f"{col_letter(VAL_COL_SUM)}{last_wp_row}")
    style_cell(ws, r, VAL_COL_SUM, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
    # TOTAL Ceiling = partner_ceiling named range
    style_cell(ws, r, VAL_COL_BUDGET, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = "=partner_ceiling"
    # TOTAL Variance = Ceiling - Sum
    style_cell(ws, r, VAL_COL_DIFF, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"={col_letter(VAL_COL_BUDGET)}{r}"
        f"-{col_letter(VAL_COL_SUM)}{r}"
    )
    # TOTAL Cumulative EUR
    ms = col_letter(MONTH_COL_START)
    me = col_letter(MONTH_COL_END)
    style_cell(ws, r, VAL_COL_CUM, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="\u20ac#,##0").value = (
        f"=SUMPRODUCT((COLUMN({ms}{r}:{me}{r})"
        f"-COLUMN(${ms}$1)+1<=current_month)"
        f"*({ms}{r}:{me}{r}))"
    )
    # TOTAL % Ceiling = Cum. EUR / partner_ceiling
    cum_l = col_letter(VAL_COL_CUM)
    style_cell(ws, r, VAL_COL_PCT_CEIL, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.0%").value = f"={cum_l}{r}/partner_ceiling"
    # TOTAL Breach In
    style_cell(ws, r, VAL_COL_BREACH, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({cum_l}{r}=0,"\u2014",'
        f"(partner_ceiling-{cum_l}{r})"
        f"/({cum_l}{r}/current_month))"
    )

    # -- Conditional formatting: red on negative variance --
    diff_col_l = col_letter(VAL_COL_DIFF)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=RED_FILL,
    ))

    # -- Conditional formatting: % Ceiling column (amber >85%, red >100%) --
    pct_col_l = col_letter(VAL_COL_PCT_CEIL)
    pct_range = f"{pct_col_l}{first_wp_row}:{pct_col_l}{last_wp_row}"
    # Amber first (lower priority), then red overrides
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>0.85"], fill=AMBER_FILL,
    ))
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>1"], fill=RED_FILL,
    ))

    # -- Conditional formatting: Breach In column (amber <12, red <6) --
    breach_col_l = col_letter(VAL_COL_BREACH)
    breach_range = f"{breach_col_l}{first_wp_row}:{breach_col_l}{last_wp_row}"
    # Amber first (lower priority), then red overrides
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<12"], fill=AMBER_FILL,
    ))
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<6"], fill=RED_FILL,
    ))

    # Freeze panes
    ws.freeze_panes = f"C{first_wp_row}"

    return {
        "header_row": 1,
        "year_row": year_row,
        "month_row": month_row,
        "first_wp_row": first_wp_row,
        "last_wp_row": last_wp_row,
        "total_row": total_row,
        "cum_col": VAL_COL_CUM,
        "pct_ceiling_col": VAL_COL_PCT_CEIL,
        "breach_col": VAL_COL_BREACH,
    }
