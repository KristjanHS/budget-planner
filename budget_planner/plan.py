"""Plan sheet builder — globals, config table, planned FTE block."""

from openpyxl.styles import Font, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    YELLOW_FILL, CENTER, style_cell, col_letter,
)


def _apply_planned_formulas(ws, layout, capacity_layout, *,
                            work_packages=None, project_params=None):
    """Overwrite planned FTE cells with cross-sheet SUM formulas from Capacity.

    Called after build_capacity_sheet() so that capacity row positions are known.
    Replaces seed values in the planned block with formulas referencing Capacity.

    Args:
        ws: The Plan worksheet.
        layout: Layout dict returned by build_plan_sheet().
        capacity_layout: Layout dict returned by build_capacity_sheet().
        work_packages: WP list; falls back to config.WORK_PACKAGES.
        project_params: Project params dict; falls back to config values.
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

    # Derived constants
    month_col_start = 3

    for i, wp in enumerate(work_packages):
        row = layout["planned_first_wp_row"] + i
        wp_id = wp["id"]

        # Find all Capacity rows for this WP across all persons
        cap_rows = []
        for person_name, group in capacity_layout["person_groups"].items():
            if wp_id in group["wp_rows"]:
                cap_rows.append(group["wp_rows"][wp_id])
        cap_rows.sort()  # deterministic order

        for m in range(project_params["months"]):
            c = month_col_start + m
            month_num = m + 1

            if wp["start"] <= month_num <= wp["end"] and cap_rows:
                cl = col_letter(c)
                refs = ",".join(f"Capacity!{cl}{cr}" for cr in cap_rows)
                cell = style_cell(ws, row, c, font=BLACK_FORMULA, fmt="0.00")
                cell.value = f"=SUM({refs})"
                # Re-lock: _build_month_block wrote these with BLUE_INPUT first
                cell.protection = Protection(locked=True)
            else:
                # Inactive month: empty + grey fill
                cell = style_cell(ws, row, c, "", font=BLACK_FORMULA,
                                  fill=GREY_FILL, fmt="0.00")
                cell.protection = Protection(locked=True)


def build_plan_sheet(ws, wb, *, project_params=None, work_packages=None):
    """Build the entire Plan sheet: globals, config table, planned FTE block.

    Args:
        ws: The Plan worksheet to populate.
        wb: The parent Workbook (needed for named ranges).
        project_params: dict with keys 'title', 'start', 'months', 'rate';
            falls back to config.py values when None.
        work_packages: list of WP dicts; falls back to config.WORK_PACKAGES.

    Returns:
        Layout dict with row numbers for all key positions.
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

    # Derived constants
    month_col_start = 3
    num_wps = len(work_packages)
    month_col_end = month_col_start + project_params["months"] - 1
    val_col_sum = month_col_end + 1
    val_col_budget = month_col_end + 2
    val_col_diff = month_col_end + 3

    r = 1

    # -- GLOBAL PARAMETERS (rows 1-4, columns A-B only) --
    style_cell(ws, r, 1, "Project Title:", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "=Config!B1",
               font=Font(name="Arial", bold=True, size=14, color="008000"),
               border=False)
    r += 1
    style_cell(ws, r, 1, "Project Start:", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "=Config!B2", font=GREEN_XREF, border=False,
               fmt="MMM YYYY")
    r += 1
    style_cell(ws, r, 1, "FTE Rate (\u20ac):", font=HEADER_FONT, border=False)
    style_cell(ws, r, 2, "=Config!B4", font=GREEN_XREF, border=False,
               fmt="\u20ac#,##0")
    r += 1
    # Personnel Ceiling -- formula filled after config table rows are known
    ceiling_row = r
    style_cell(ws, r, 1, "Personnel Ceiling:", font=HEADER_FONT, border=False)
    r += 1

    # partner_ceiling added after config table is built

    # -- CONFIG TABLE --
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
    for i, wp in enumerate(work_packages):
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
    config_last_wp_row = config_first_wp_row + num_wps - 1
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

    # -- PLANNED FTE BLOCK --
    planned = _build_month_block(
        ws, r, "PLANNED FTE (PERSON-MONTHS)",
        config_first_wp_row=config_first_wp_row,
        project_params=project_params,
        work_packages=work_packages,
        month_col_start=month_col_start,
        month_col_end=month_col_end,
        num_wps=num_wps,
        val_col_sum=val_col_sum,
        val_col_budget=val_col_budget,
        val_col_diff=val_col_diff,
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


def _build_month_block(ws, start_row, block_title, config_first_wp_row, *,
                       project_params, work_packages, month_col_start,
                       month_col_end, num_wps, val_col_sum, val_col_budget,
                       val_col_diff):
    """Build a planned FTE month block.

    Args:
        ws: The Plan worksheet.
        start_row: Row to start writing at.
        block_title: Section header text.
        config_first_wp_row: First WP row in the config table (for cross-refs).
        project_params: Project params dict with 'start', 'months'.
        work_packages: List of WP dicts.
        month_col_start: First month column number.
        month_col_end: Last month column number.
        num_wps: Number of work packages.
        val_col_sum: Sum validation column number.
        val_col_budget: Budget validation column number.
        val_col_diff: Variance validation column number.

    Returns:
        Dict with header_row, year_row, month_row, first_wp_row, last_wp_row,
        total_row, next_row.
    """
    r = start_row
    project_months = project_params["months"]
    project_start = project_params["start"]

    # Section header
    merge_end = val_col_diff
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
    while m < project_months:
        year_start_col = month_col_start + m
        months_left = project_months - m
        span = min(12, months_left)
        year_end_col = year_start_col + span - 1
        cal_date = project_start + relativedelta(months=m)
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
    style_cell(ws, r, val_col_sum, "Sum", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, val_col_budget, "Budget FTE", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, val_col_diff, "Variance", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1
    month_row = r

    # Month header row
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    for m in range(project_months):
        c = month_col_start + m
        cal_date = project_start + relativedelta(months=m)
        mlabel = f"M{m+1}\n{cal_date.strftime('%b %y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8),
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, val_col_sum, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, val_col_budget, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, val_col_diff, "", fill=LIGHT_BLUE_FILL)
    r += 1
    first_wp_row = r

    # WP rows
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = config_first_wp_row + i
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"=A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"=B{config_row}"

        active_months = wp["end"] - wp["start"] + 1
        even_fte = round(wp["fte"] / active_months, 4) if active_months > 0 else 0

        for m in range(project_months):
            c = month_col_start + m
            month_num = m + 1

            if wp["start"] <= month_num <= wp["end"]:
                style_cell(ws, row, c, even_fte, font=BLUE_INPUT, fmt="0.00")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="0.00")

        # Validation columns
        sum_range = (f"{col_letter(month_col_start)}{row}:"
                     f"{col_letter(month_col_end)}{row}")
        style_cell(ws, row, val_col_sum, font=BLACK_FORMULA,
                   fmt="0.00").value = f"=SUM({sum_range})"
        style_cell(ws, row, val_col_budget, font=GREEN_XREF,
                   fmt="0.00").value = f"=C{config_row}"
        style_cell(ws, row, val_col_diff, font=BLACK_FORMULA,
                   fmt="0.00").value = (
            f"={col_letter(val_col_budget)}{row}"
            f"-{col_letter(val_col_sum)}{row}"
        )

    last_wp_row = first_wp_row + num_wps - 1
    r = last_wp_row + 1

    # Totals row
    total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "MONTHLY TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    for m in range(project_months):
        c = month_col_start + m
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="0.00").value = f"=SUM({col_range})"
    sum_range = (f"{col_letter(val_col_sum)}{first_wp_row}:"
                 f"{col_letter(val_col_sum)}{last_wp_row}")
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({sum_range})"
    bud_range = (f"{col_letter(val_col_budget)}{first_wp_row}:"
                 f"{col_letter(val_col_budget)}{last_wp_row}")
    style_cell(ws, r, val_col_budget, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({bud_range})"
    diff_range = (f"{col_letter(val_col_diff)}{first_wp_row}:"
                  f"{col_letter(val_col_diff)}{last_wp_row}")
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({diff_range})"

    # Conditional formatting: yellow fill on variance if nonzero
    diff_col_l = col_letter(val_col_diff)
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
