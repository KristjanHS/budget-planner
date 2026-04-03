"""Actuals sheet builder — EUR actuals per WP per month from accountant."""

from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_ACTUALS, SHEET_CONFIG,
    ACTUALS_AMBER_PCT_CEIL, ACTUALS_AMBER_BREACH_MONTHS, ACTUALS_RED_BREACH_MONTHS,
    build_year_month_headers, month_col_end as _month_col_end, val_columns,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_GREEN_FILL,
    RED_FILL, AMBER_FILL, CENTER, style_cell, col_letter,
)


def build_actuals_sheet(wb, layout, *, project_params=None, work_packages=None, preserved=None):
    """Build the Actuals sheet for EUR actuals per WP per month.

    Monthly EUR cells are always left empty (blue input) for manual entry
    or demo seeding. When ``preserved`` is provided (rebuild), those values
    are restored.

    Args:
        wb: The parent Workbook.
        layout: Layout dict from build_plan_sheet(), needs config_first_wp_row.

    Returns:
        Layout dict with keys: header_row, year_row, month_row,
        first_wp_row, last_wp_row, total_row.
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
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    val_col_cum = _vcols["cum"]
    val_col_pct_ceil = _vcols["pct_ceil"]
    val_col_breach = _vcols["breach"]

    config_first_wp_row = layout["config_first_wp_row"]

    ws = wb.create_sheet(SHEET_ACTUALS)
    ws.sheet_properties.tabColor = "0000FF"

    r = 1

    # -- Section header --
    merge_end = val_col_breach
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_end)
    style_cell(ws, r, 1, "ACTUAL EUR", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, merge_end + 1):
        style_cell(ws, r, c2, fill=SECTION_FILL)
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

        for m_idx in range(project_params["months"]):
            c = month_col_start + m_idx
            month_num = m_idx + 1
            if wp["start"] <= month_num <= wp["end"]:
                if preserved_values is not None and m_idx < len(preserved_values) and preserved_values[m_idx] is not None:
                    eur_value: int | float = preserved_values[m_idx]
                    style_cell(ws, row, c, eur_value, font=BLUE_INPUT, fmt="\u20ac#,##0")
                else:
                    style_cell(ws, row, c, 0, font=BLUE_INPUT, fmt="\u20ac#,##0")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="\u20ac#,##0")

        # Validation: Sum
        sum_range = (f"{col_letter(month_col_start)}{row}:"
                     f"{col_letter(month_col_end)}{row}")
        style_cell(ws, row, val_col_sum, font=BLACK_FORMULA,
                   fmt="\u20ac#,##0").value = f"=SUM({sum_range})"
        # Validation: Ceiling (from Config Personnel Ceiling column G)
        style_cell(ws, row, val_col_budget, font=GREEN_XREF,
                   fmt="\u20ac#,##0").value = f"={SHEET_CONFIG}!G{config_row}"
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
        # Breach In = months until ceiling exhausted at current rate
        style_cell(ws, row, val_col_breach, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({cum_l}{row}=0,"\u2014",'
            f"({ceil_l}{row}-{cum_l}{row})"
            f"/({cum_l}{row}/current_month))"
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
    # TOTAL Breach In
    style_cell(ws, r, val_col_breach, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({cum_l}{r}=0,"\u2014",'
        f"(partner_ceiling-{cum_l}{r})"
        f"/({cum_l}{r}/current_month))"
    )

    # -- Conditional formatting: red on negative variance --
    diff_col_l = col_letter(val_col_diff)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=RED_FILL,
    ))

    # -- Conditional formatting: % Ceiling column (amber >ACTUALS_AMBER_PCT_CEIL, red >100%) --
    pct_col_l = col_letter(val_col_pct_ceil)
    pct_range = f"{pct_col_l}{first_wp_row}:{pct_col_l}{last_wp_row}"
    # Amber first (lower priority), then red overrides
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"], fill=AMBER_FILL,
    ))
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>1"], fill=RED_FILL,
    ))

    # -- Conditional formatting: Breach In column (amber <ACTUALS_AMBER_BREACH_MONTHS, red <ACTUALS_RED_BREACH_MONTHS) --
    breach_col_l = col_letter(val_col_breach)
    breach_range = f"{breach_col_l}{first_wp_row}:{breach_col_l}{last_wp_row}"
    # Amber first (lower priority), then red overrides
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"], fill=AMBER_FILL,
    ))
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"], fill=RED_FILL,
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
        "cum_col": val_col_cum,
        "pct_ceiling_col": val_col_pct_ceil,
        "breach_col": val_col_breach,
    }
