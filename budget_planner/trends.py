"""Trends sheet builder — WP burn rate, cumulative chart, person load."""

from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_MONTHS, WORK_PACKAGES, NUM_WPS, TEAM,
    MONTH_COL_START, MONTH_COL_END, VAL_COL_BUDGET,
)

# Local fill constants for conditional formatting
_GREEN_CF_FILL = PatternFill("solid", fgColor="E2EFDA")
_AMBER_CF_FILL = PatternFill("solid", fgColor="FFD966")
_RED_CF_FILL = PatternFill("solid", fgColor="FFC7CE")


def build_trends_sheet(ws, wb, layout, actuals_layout, capacity_layout, team_layout):
    """Build the Trends sheet: WP burn rate, cumulative chart, person load.

    Args:
        ws: The Trends worksheet (already has current_month in B1).
        wb: The parent Workbook.
        layout: Plan layout dict (config_first_wp_row, etc).
        actuals_layout: Actuals layout dict (first_wp_row, cum_col, etc).
        capacity_layout: Capacity layout dict (person_groups with total_row per person).
        team_layout: Team layout dict (person_rows mapping name->row).
    """
    cum_col_letter = col_letter(actuals_layout["cum_col"])
    ceiling_col_letter = col_letter(VAL_COL_BUDGET)
    ms = col_letter(MONTH_COL_START)
    me = col_letter(MONTH_COL_END)

    # ================================================================
    # Section 1: WP Burn Rate & Exhaustion Table (rows 3-13)
    # ================================================================
    section1_row = 3
    ws.merge_cells(start_row=section1_row, start_column=1,
                   end_row=section1_row, end_column=7)
    style_cell(ws, section1_row, 1, "WP BURN RATE & EXHAUSTION",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER)
    for c in range(2, 8):
        style_cell(ws, section1_row, c, fill=SECTION_FILL)

    # Column headers (row 4)
    header_row = 4
    headers = [
        "WP ID", "WP Name", "% Time Elapsed", "% Budget EUR",
        "Burn Delta", "Status", "Exhaustion Mo.",
    ]
    for ci, h in enumerate(headers, 1):
        style_cell(ws, header_row, ci, h, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Cell comment on Exhaustion Mo. header
    ws.cell(row=header_row, column=7).comment = Comment(
        "Projection is linear based on quarterly actuals. "
        "Accuracy improves after Q2.",
        "Programme Budget Tool",
    )

    # WP data rows (rows 5-12)
    first_wp_data_row = 5
    last_wp_data_row = first_wp_data_row + NUM_WPS - 1

    for i, wp in enumerate(WORK_PACKAGES):
        row = first_wp_data_row + i
        config_row = layout["config_first_wp_row"] + i
        actuals_row = actuals_layout["first_wp_row"] + i
        actuals_month_range = f"Actuals!{ms}{actuals_row}:{me}{actuals_row}"

        # Col A: WP ID (green cross-sheet ref)
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"=Plan!A{config_row}"

        # Col B: WP Name (green cross-sheet ref)
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"=Plan!B{config_row}"

        # Col C: % Time Elapsed
        time_formula = (
            f"MIN(1,(current_month-Plan!D{config_row}+1)"
            f"/(Plan!E{config_row}-Plan!D{config_row}+1))"
        )
        style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="0.0%").value = (
            f'=IF(current_month<Plan!D{config_row},"\u2014",{time_formula})'
        )

        # Col D: % Budget EUR
        budget_formula = (
            f"Actuals!{cum_col_letter}{actuals_row}"
            f"/Actuals!{ceiling_col_letter}{actuals_row}"
        )
        style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="0.0%").value = (
            f'=IF(current_month<Plan!D{config_row},"\u2014",{budget_formula})'
        )

        # Col E: Burn Delta
        style_cell(ws, row, 5, font=BLACK_FORMULA, fmt="0.0%").value = (
            f'=IF(current_month<Plan!D{config_row},"\u2014",D{row}-C{row})'
        )

        # Col F: Status
        status_formula = (
            f'IF(ABS(E{row})<=0.05,"OK",'
            f'IF(E{row}>0.15,"OVER",'
            f'IF(E{row}>0.05,"WATCH",'
            f'IF(E{row}<-0.10,"SLOW","OK"))))'
        )
        style_cell(ws, row, 6, font=BLACK_FORMULA).value = (
            f'=IF(current_month<Plan!D{config_row},"\u2014",{status_formula})'
        )

        # Col G: Exhaustion Mo.
        exhaustion_formula = (
            f"IF(D{row}=0,\"\u2014\","
            f"IF(COUNTIF({actuals_month_range},\">0\")=0,\"No data\","
            f"Plan!D{config_row}+ROUND("
            f"Actuals!{ceiling_col_letter}{actuals_row}"
            f"/(Actuals!{cum_col_letter}{actuals_row}"
            f"/COUNTIF({actuals_month_range},\">0\")),0)-1))"
        )
        style_cell(ws, row, 7, font=BLACK_FORMULA, fmt="0").value = (
            f'=IF(current_month<Plan!D{config_row},"\u2014",{exhaustion_formula})'
        )

    # -- Conditional formatting: Burn Delta (col E, rows 5-12) --
    e_range = f"E{first_wp_data_row}:E{last_wp_data_row}"
    e_ref = f"E{first_wp_data_row}"

    # Green first (lowest priority) — ABS(delta) <= 0.05
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'ABS({e_ref})<=0.05'], fill=_GREEN_CF_FILL,
    ))
    # Amber: overspend (delta > 0.05 AND delta <= 0.15)
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'AND({e_ref}>0.05,{e_ref}<=0.15)'], fill=_AMBER_CF_FILL,
    ))
    # Amber: underspend (delta < -0.10)
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'{e_ref}<-0.10'], fill=_AMBER_CF_FILL,
    ))
    # Red: delta > 0.15
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'{e_ref}>0.15'], fill=_RED_CF_FILL,
    ))

    # -- Conditional formatting: Status (col F, rows 5-12) --
    f_range = f"F{first_wp_data_row}:F{last_wp_data_row}"
    f_ref = f"F{first_wp_data_row}"

    # Green: "OK"
    ws.conditional_formatting.add(f_range, FormulaRule(
        formula=[f'{f_ref}="OK"'], fill=_GREEN_CF_FILL,
    ))
    # Amber: "WATCH" or "SLOW"
    ws.conditional_formatting.add(f_range, FormulaRule(
        formula=[f'OR({f_ref}="WATCH",{f_ref}="SLOW")'], fill=_AMBER_CF_FILL,
    ))
    # Red: "OVER"
    ws.conditional_formatting.add(f_range, FormulaRule(
        formula=[f'{f_ref}="OVER"'], fill=_RED_CF_FILL,
    ))

    # -- Conditional formatting: Exhaustion Month (col G) — per WP row --
    # Each WP has a different end month, so CF must reference the correct config row
    for i in range(NUM_WPS):
        row = first_wp_data_row + i
        config_row = layout["config_first_wp_row"] + i
        g_cell = f"G{row}:G{row}"
        g_ref = f"G{row}"
        # Green first: exhaustion >= wp_end
        ws.conditional_formatting.add(g_cell, FormulaRule(
            formula=[f'AND(ISNUMBER({g_ref}),{g_ref}>=Plan!E{config_row})'],
            fill=_GREEN_CF_FILL,
        ))
        # Amber: exhaustion within 3 months of wp_end
        ws.conditional_formatting.add(g_cell, FormulaRule(
            formula=[
                f'AND(ISNUMBER({g_ref}),'
                f'{g_ref}<Plan!E{config_row},'
                f'{g_ref}>=Plan!E{config_row}-3)'
            ],
            fill=_AMBER_CF_FILL,
        ))
        # Red: exhaustion < wp_end - 3
        ws.conditional_formatting.add(g_cell, FormulaRule(
            formula=[
                f'AND(ISNUMBER({g_ref}),'
                f'{g_ref}<Plan!E{config_row}-3)'
            ],
            fill=_RED_CF_FILL,
        ))

    # -- TOTAL row (row 13) --
    total_row = last_wp_data_row + 1
    act_first = actuals_layout["first_wp_row"]
    act_last = actuals_layout["last_wp_row"]
    style_cell(ws, total_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 2, "TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 3, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 4, "", fill=LIGHT_GREEN_FILL)
    # Weighted average burn delta (ISNUMBER guard for WPs showing "—")
    style_cell(ws, total_row, 5, font=BLACK_FORMULA,
               fill=LIGHT_GREEN_FILL, fmt="0.0%").value = (
        f"=SUMPRODUCT(IF(ISNUMBER(E{first_wp_data_row}:E{last_wp_data_row}),"
        f"E{first_wp_data_row}:E{last_wp_data_row},0),"
        f"Actuals!{ceiling_col_letter}{act_first}"
        f":{ceiling_col_letter}{act_last})"
        f"/partner_ceiling"
    )
    style_cell(ws, total_row, 6, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 7, "", fill=LIGHT_GREEN_FILL)

    # ================================================================
    # Section 2: Person Load Table (rows 32-37+)
    # ================================================================
    person_section_row = 32
    ws.merge_cells(start_row=person_section_row, start_column=1,
                   end_row=person_section_row, end_column=5)
    style_cell(ws, person_section_row, 1, "PERSON LOAD",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER)
    for c in range(2, 6):
        style_cell(ws, person_section_row, c, fill=SECTION_FILL)

    # Column headers (row 33)
    person_header_row = 33
    person_headers = [
        "Person", "Avg FTE (to date)", "Avg FTE (remaining)",
        "Peak Month FTE", "Status",
    ]
    for ci, h in enumerate(person_headers, 1):
        style_cell(ws, person_header_row, ci, h, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Person data rows (row 34+)
    person_first_row = 34
    person_groups = capacity_layout["person_groups"]
    person_rows_map = team_layout["person_rows"]

    for pi, member in enumerate(TEAM):
        name = member["name"]
        row = person_first_row + pi
        tr = person_groups[name]["total_row"]
        roster_row = person_rows_map[name]

        # Col A: Person name (green cross-sheet ref)
        style_cell(ws, row, 1, font=GREEN_XREF).value = (
            f"=Team!A{roster_row}"
        )

        # Col B: Avg FTE (to date)
        style_cell(ws, row, 2, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=IF(current_month=0,0,"
            f"SUMPRODUCT((COLUMN(Capacity!{ms}{tr}:{me}{tr})"
            f"-COLUMN(Capacity!${ms}$1)+1<=current_month)"
            f"*(Capacity!{ms}{tr}:{me}{tr}))"
            f"/current_month)"
        )

        # Col C: Avg FTE (remaining)
        style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=IF(current_month>={PROJECT_MONTHS},0,"
            f"SUMPRODUCT((COLUMN(Capacity!{ms}{tr}:{me}{tr})"
            f"-COLUMN(Capacity!${ms}$1)+1>current_month)"
            f"*(Capacity!{ms}{tr}:{me}{tr}))"
            f"/({PROJECT_MONTHS}-current_month))"
        )

        # Col D: Peak Month FTE
        style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=MAX(Capacity!{ms}{tr}:{me}{tr})"
        )

        # Col E: Status
        style_cell(ws, row, 5, font=BLACK_FORMULA).value = (
            f'=IF(C{row}>Team!B{roster_row},"OVER",'
            f'IF(C{row}>0.8*Team!B{roster_row},"WATCH","OK"))'
        )

    person_last_row = person_first_row + len(TEAM) - 1

    # -- Conditional formatting: Person Status (col E) + Peak (col D) — per person --
    # Each person has a different max_fte, so CF must use the correct roster row
    for pi, member in enumerate(TEAM):
        row = person_first_row + pi
        roster_row = person_rows_map[member["name"]]
        e_cell = f"E{row}:E{row}"
        d_cell = f"D{row}:D{row}"

        # Status CF (green/amber/red based on col C vs max_fte)
        ws.conditional_formatting.add(e_cell, FormulaRule(
            formula=[f'C{row}<=0.8*Team!$B${roster_row}'],
            fill=_GREEN_CF_FILL,
        ))
        ws.conditional_formatting.add(e_cell, FormulaRule(
            formula=[f'C{row}>0.8*Team!$B${roster_row}'],
            fill=_AMBER_CF_FILL,
        ))
        ws.conditional_formatting.add(e_cell, FormulaRule(
            formula=[f'C{row}>Team!$B${roster_row}'],
            fill=_RED_CF_FILL,
        ))

        # Peak Month CF (red if > max_fte)
        ws.conditional_formatting.add(d_cell, FormulaRule(
            formula=[f'D{row}>Team!$B${roster_row}'],
            fill=_RED_CF_FILL,
        ))

    # ================================================================
    # Section 3: Cumulative LineChart + hidden helper rows
    # ================================================================
    # Helper rows placed at row 40+ (below person load section)
    helper_row_planned = 40
    helper_row_actual = 41
    helper_row_ceiling = 42

    plan_first_wp = layout["planned_first_wp_row"]
    plan_last_wp = layout["planned_last_wp_row"]
    act_total = actuals_layout["total_row"]

    for m_idx in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m_idx
        cl = col_letter(c)

        # Helper Row 1: Cumulative planned EUR
        if m_idx == 0:
            # First month: just the monthly total * fte_rate
            style_cell(ws, helper_row_planned, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"=SUM(Plan!{cl}{plan_first_wp}:{cl}{plan_last_wp})*fte_rate"
            )
        else:
            prev_cl = col_letter(c - 1)
            style_cell(ws, helper_row_planned, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={prev_cl}{helper_row_planned}"
                f"+SUM(Plan!{cl}{plan_first_wp}:{cl}{plan_last_wp})*fte_rate"
            )

        # Helper Row 2: Cumulative actual EUR
        if m_idx == 0:
            style_cell(ws, helper_row_actual, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"=Actuals!{cl}{act_total}"
            )
        else:
            prev_cl = col_letter(c - 1)
            style_cell(ws, helper_row_actual, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={prev_cl}{helper_row_actual}+Actuals!{cl}{act_total}"
            )

        # Helper Row 3: Ceiling line (flat)
        style_cell(ws, helper_row_ceiling, c,
                   font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
            "=partner_ceiling"
        )

    # Hide helper rows
    ws.row_dimensions[helper_row_planned].hidden = True
    ws.row_dimensions[helper_row_actual].hidden = True
    ws.row_dimensions[helper_row_ceiling].hidden = True

    # -- Build the LineChart --
    chart = LineChart()
    chart.title = "Cumulative EUR: Planned vs Actual"
    chart.style = 10
    chart.y_axis.title = "EUR"
    chart.x_axis.title = "Month"
    chart.y_axis.numFmt = "\u20ac#,##0"

    # Data references (columns C through BJ = cols 3..62)
    min_col = MONTH_COL_START
    max_col = MONTH_COL_END

    # Series 1: Cumulative planned (blue solid)
    planned_ref = Reference(ws, min_col=min_col, max_col=max_col,
                            min_row=helper_row_planned,
                            max_row=helper_row_planned)
    chart.add_data(planned_ref, titles_from_data=False)
    s_planned = chart.series[0]
    s_planned.tx = SeriesLabel(v="Planned")
    s_planned.graphicalProperties.line.solidFill = "4472C4"

    # Series 2: Cumulative actual (red solid)
    actual_ref = Reference(ws, min_col=min_col, max_col=max_col,
                           min_row=helper_row_actual,
                           max_row=helper_row_actual)
    chart.add_data(actual_ref, titles_from_data=False)
    s_actual = chart.series[1]
    s_actual.tx = SeriesLabel(v="Actual")
    s_actual.graphicalProperties.line.solidFill = "FF0000"

    # Series 3: Ceiling (grey dashed)
    ceiling_ref = Reference(ws, min_col=min_col, max_col=max_col,
                            min_row=helper_row_ceiling,
                            max_row=helper_row_ceiling)
    chart.add_data(ceiling_ref, titles_from_data=False)
    s_ceiling = chart.series[2]
    s_ceiling.tx = SeriesLabel(v="Ceiling")
    s_ceiling.graphicalProperties.line.solidFill = "808080"
    s_ceiling.graphicalProperties.line.dashStyle = "dash"

    # X-axis categories: M1..M60 labels
    # Create labels in a helper area (row 43, hidden)
    helper_row_labels = 43
    for m_idx in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m_idx
        ws.cell(row=helper_row_labels, column=c, value=f"M{m_idx + 1}")
    ws.row_dimensions[helper_row_labels].hidden = True

    cats = Reference(ws, min_col=min_col, max_col=max_col,
                     min_row=helper_row_labels, max_row=helper_row_labels)
    chart.set_categories(cats)

    # Chart dimensions and placement
    chart.width = 30
    chart.height = 15
    ws.add_chart(chart, "A15")

    # ================================================================
    # Return layout dict
    # ================================================================
    return {
        "section_header_row": section1_row,
        "burn_rate_header_row": header_row,
        "burn_rate_first_wp_row": first_wp_data_row,
        "burn_rate_last_wp_row": last_wp_data_row,
        "burn_rate_total_row": total_row,
        "chart_anchor_row": 15,
        "person_load_section_row": person_section_row,
        "person_load_header_row": person_header_row,
        "person_load_first_row": person_first_row,
        "person_load_last_row": person_last_row,
        "helper_row_planned": helper_row_planned,
        "helper_row_actual": helper_row_actual,
        "helper_row_ceiling": helper_row_ceiling,
        "helper_row_labels": helper_row_labels,
    }
