"""Plan sheet builder — globals, config table, planned FTE block."""

from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    YELLOW_FILL, CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_TITLE, PROJECT_START, PROJECT_MONTHS, DEFAULT_RATE,
    WORK_PACKAGES, NUM_WPS, MONTH_COL_START, MONTH_COL_END,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
)


def _apply_planned_formulas(ws, layout, capacity_layout):
    """Overwrite planned FTE cells with cross-sheet SUM formulas from Capacity.

    Called after build_capacity_sheet() so that capacity row positions are known.
    Replaces seed values in the planned block with formulas referencing Capacity.

    Args:
        ws: The Plan worksheet.
        layout: Layout dict returned by build_plan_sheet().
        capacity_layout: Layout dict returned by build_capacity_sheet().
    """
    for i, wp in enumerate(WORK_PACKAGES):
        row = layout["planned_first_wp_row"] + i
        wp_id = wp["id"]

        # Find all Capacity rows for this WP across all persons
        cap_rows = []
        for person_name, group in capacity_layout["person_groups"].items():
            if wp_id in group["wp_rows"]:
                cap_rows.append(group["wp_rows"][wp_id])
        cap_rows.sort()  # deterministic order

        for m in range(PROJECT_MONTHS):
            c = MONTH_COL_START + m
            month_num = m + 1

            if wp["start"] <= month_num <= wp["end"] and cap_rows:
                cl = col_letter(c)
                refs = ",".join(f"Capacity!{cl}{cr}" for cr in cap_rows)
                cell = style_cell(ws, row, c, font=BLACK_FORMULA, fmt="0.00")
                cell.value = f"=SUM({refs})"
            else:
                # Inactive month: empty + grey fill
                style_cell(ws, row, c, "", font=BLACK_FORMULA, fill=GREY_FILL,
                           fmt="0.00")


def build_plan_sheet(ws, wb):
    """Build the entire Plan sheet: globals, config table, planned FTE block.

    Args:
        ws: The Plan worksheet to populate.
        wb: The parent Workbook (needed for named ranges).

    Returns:
        Layout dict with row numbers for all key positions.
    """
    r = 1

    # ── GLOBAL PARAMETERS (rows 1–4, columns A–B only) ──
    style_cell(ws, r, 1, "Project Title:", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, PROJECT_TITLE,
               font=Font(name="Arial", bold=True, size=14, color="0000FF"),
               border=False)
    r += 1
    style_cell(ws, r, 1, "Project Start:", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, PROJECT_START, font=BLUE_INPUT, border=False,
               fmt="MMM YYYY")
    r += 1
    style_cell(ws, r, 1, "FTE Rate (\u20ac):", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, DEFAULT_RATE, font=BLUE_INPUT, border=False,
               fmt="\u20ac#,##0")
    r += 1
    # Personnel Ceiling — formula filled after config table rows are known
    ceiling_row = r
    style_cell(ws, r, 1, "Personnel Ceiling:", font=HEADER_FONT, border=False)
    r += 1

    # Named ranges
    wb.defined_names.add(DefinedName("project_start", attr_text="Plan!$B$2"))
    wb.defined_names.add(DefinedName("fte_rate", attr_text="Plan!$B$3"))
    # partner_ceiling added after config table is built

    # ── CONFIG TABLE ──
    r += 1
    config_header_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    style_cell(ws, r, 1, "CONFIGURATION", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, 8):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1

    config_col_headers_row = r
    config_headers = [
        "WP ID", "WP Name", "Total FTE (PM)", "Start Month",
        "End Month", "Active Months", "Personnel Ceiling (\u20ac)",
    ]
    for ci, h in enumerate(config_headers, 1):
        style_cell(ws, r, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    config_first_wp_row = r
    for i, wp in enumerate(WORK_PACKAGES):
        row = r + i
        style_cell(ws, row, 1, wp["id"], font=BLUE_INPUT)
        style_cell(ws, row, 2, wp["name"], font=BLUE_INPUT)
        style_cell(ws, row, 3, wp["fte"], font=BLUE_INPUT, fmt="0.00")
        style_cell(ws, row, 4, wp["start"], font=BLUE_INPUT, fmt="0")
        style_cell(ws, row, 5, wp["end"], font=BLUE_INPUT, fmt="0")
        style_cell(ws, row, 6, font=BLACK_FORMULA, fmt="0").value = (
            f"=E{row}-D{row}+1"
        )
        style_cell(ws, row, 7, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
            f"=C{row}*fte_rate"
        )
    config_last_wp_row = config_first_wp_row + NUM_WPS - 1
    r = config_last_wp_row + 1

    config_total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 3, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
               fmt="0.00").value = (
        f"=SUM(C{config_first_wp_row}:C{config_last_wp_row})"
    )
    style_cell(ws, r, 4, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 5, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 6, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 7, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
               fmt="\u20ac#,##0").value = (
        f"=SUM(G{config_first_wp_row}:G{config_last_wp_row})"
    )

    # Fill Personnel Ceiling global param
    style_cell(ws, ceiling_row, 2, font=BLACK_FORMULA, border=False,
               fmt="\u20ac#,##0").value = (
        f"=SUM(G{config_first_wp_row}:G{config_last_wp_row})"
    )
    wb.defined_names.add(DefinedName("partner_ceiling", attr_text="Plan!$B$4"))

    r += 2

    # ── PLANNED FTE BLOCK ──
    planned = _build_month_block(
        ws, r, "PLANNED FTE (PERSON-MONTHS)",
        config_first_wp_row=config_first_wp_row,
    )

    # Freeze panes: freeze WP ID/Name columns and header rows for planned block
    ws.freeze_panes = f"C{planned['first_wp_row']}"

    return {
        "config_header_row": config_header_row,
        "config_col_headers_row": config_col_headers_row,
        "config_first_wp_row": config_first_wp_row,
        "config_last_wp_row": config_last_wp_row,
        "config_total_row": config_total_row,
        "planned_header_row": planned["header_row"],
        "planned_first_wp_row": planned["first_wp_row"],
        "planned_last_wp_row": planned["last_wp_row"],
        "planned_total_row": planned["total_row"],
    }


def _build_month_block(ws, start_row, block_title, config_first_wp_row):
    """Build a planned FTE month block.

    Args:
        ws: The Plan worksheet.
        start_row: Row to start writing at.
        block_title: Section header text.
        config_first_wp_row: First WP row in the config table (for cross-refs).

    Returns:
        Dict with header_row, year_row, month_row, first_wp_row, last_wp_row,
        total_row, next_row.
    """
    r = start_row

    # Section header
    merge_end = VAL_COL_DIFF
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_end)
    style_cell(ws, r, 1, block_title, font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, merge_end + 1):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1
    year_row = r

    # Year header row
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
    # Validation headers in year row
    style_cell(ws, r, VAL_COL_SUM, "Sum", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_BUDGET, "Budget FTE", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_DIFF, "Variance", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1
    month_row = r

    # Month header row
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    for m in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m
        cal_date = PROJECT_START + relativedelta(months=m)
        mlabel = f"M{m+1}\n{cal_date.strftime('%b %y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8),
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_SUM, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_BUDGET, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_DIFF, "", fill=LIGHT_BLUE_FILL)
    r += 1
    first_wp_row = r

    # WP rows
    for i, wp in enumerate(WORK_PACKAGES):
        row = r + i
        config_row = config_first_wp_row + i
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"=A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"=B{config_row}"

        active_months = wp["end"] - wp["start"] + 1
        even_fte = round(wp["fte"] / active_months, 4) if active_months > 0 else 0

        for m in range(PROJECT_MONTHS):
            c = MONTH_COL_START + m
            month_num = m + 1

            if wp["start"] <= month_num <= wp["end"]:
                style_cell(ws, row, c, even_fte, font=BLUE_INPUT, fmt="0.00")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="0.00")

        # Validation columns
        sum_range = (f"{col_letter(MONTH_COL_START)}{row}:"
                     f"{col_letter(MONTH_COL_END)}{row}")
        style_cell(ws, row, VAL_COL_SUM, font=BLACK_FORMULA,
                   fmt="0.00").value = f"=SUM({sum_range})"
        style_cell(ws, row, VAL_COL_BUDGET, font=GREEN_XREF,
                   fmt="0.00").value = f"=C{config_row}"
        style_cell(ws, row, VAL_COL_DIFF, font=BLACK_FORMULA,
                   fmt="0.00").value = (
            f"={col_letter(VAL_COL_BUDGET)}{row}"
            f"-{col_letter(VAL_COL_SUM)}{row}"
        )

    last_wp_row = first_wp_row + NUM_WPS - 1
    r = last_wp_row + 1

    # Totals row
    total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "MONTHLY TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    for m in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="0.00").value = f"=SUM({col_range})"
    sum_range = (f"{col_letter(VAL_COL_SUM)}{first_wp_row}:"
                 f"{col_letter(VAL_COL_SUM)}{last_wp_row}")
    style_cell(ws, r, VAL_COL_SUM, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({sum_range})"
    bud_range = (f"{col_letter(VAL_COL_BUDGET)}{first_wp_row}:"
                 f"{col_letter(VAL_COL_BUDGET)}{last_wp_row}")
    style_cell(ws, r, VAL_COL_BUDGET, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({bud_range})"
    diff_range = (f"{col_letter(VAL_COL_DIFF)}{first_wp_row}:"
                  f"{col_letter(VAL_COL_DIFF)}{last_wp_row}")
    style_cell(ws, r, VAL_COL_DIFF, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({diff_range})"

    # Conditional formatting: yellow fill on variance if nonzero
    diff_col_l = col_letter(VAL_COL_DIFF)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"ABS({diff_col_l}{first_wp_row})>0.01"], fill=YELLOW_FILL,
    ))

    return {
        "header_row": start_row,
        "year_row": year_row,
        "month_row": month_row,
        "first_wp_row": first_wp_row,
        "last_wp_row": last_wp_row,
        "total_row": total_row,
        "next_row": r + 2,
    }
