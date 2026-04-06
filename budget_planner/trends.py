"""Trends sheet builder — WP burn rate, cumulative chart, person load."""

from dateutil.relativedelta import relativedelta
from openpyxl.chart import LineChart
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_PLAN, SHEET_CONFIG, SHEET_ACTUALS, SHEET_CAPACITY,
    TREND_OK_BAND, TREND_OVER_THRESHOLD, TREND_WATCH_THRESHOLD, TREND_SLOW_THRESHOLD,
    CAPACITY_AMBER_LOAD, month_col_end as _month_col_end, val_columns,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CENTER, style_cell, col_letter, write_section_header,
    auto_fit_columns,
)

# CF fills imported from styles — DXF context requires bgColor
from budget_planner.styles import CF_GREEN_FILL as _GREEN_CF_FILL
from budget_planner.styles import CF_AMBER_FILL as _AMBER_CF_FILL
from budget_planner.styles import CF_RED_FILL as _RED_CF_FILL


def build_trends_sheet(ws, wb, layout, actuals_layout, capacity_layout, team_layout, *,
                       project_params=None, work_packages=None, team=None,
                       reporting_periods=None):
    """Build the Trends sheet: WP burn rate, cumulative chart, person load.

    Args:
        ws: The Trends worksheet (already has current_month in B1).
        wb: The parent Workbook.
        layout: Plan layout dict (config_first_wp_row, etc).
        actuals_layout: Actuals layout dict (first_wp_row, cum_col, etc).
        capacity_layout: Capacity layout dict (person_groups with total_row per person).
        team_layout: Team layout dict (person_rows mapping name->row).
        project_params: Dict with keys title, start, months, rate. Falls back to config.
        work_packages: List of WP dicts. Falls back to config.
        team: List of team member dicts. Falls back to config.
    """
    if any(v is None for v in [project_params, work_packages, team]):
        from budget_planner import config
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES
        if team is None:
            team = config.TEAM

    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS

    assert project_params is not None
    assert work_packages is not None
    assert team is not None

    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    month_col_end = _month_col_end(project_params["months"])
    val_col_budget = val_columns(project_params["months"])["budget"]

    cum_col_letter = col_letter(actuals_layout["cum_col"])
    ceiling_col_letter = col_letter(val_col_budget)
    ms = col_letter(month_col_start)
    me = col_letter(month_col_end)

    # ================================================================
    # Section 1: WP Burn Rate & Exhaustion Table
    # ================================================================
    section1_row = 3
    write_section_header(ws, section1_row, "WP BURN RATE & EXHAUSTION", merge_end=7)

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

    # WP data rows
    first_wp_data_row = header_row + 1
    last_wp_data_row = first_wp_data_row + num_wps - 1

    for i, wp in enumerate(work_packages):
        row = first_wp_data_row + i
        config_row = layout["config_first_wp_row"] + i
        actuals_row = actuals_layout["first_wp_row"] + i
        actuals_month_range = f"{SHEET_ACTUALS}!{ms}{actuals_row}:{me}{actuals_row}"

        # Col A: WP ID (green cross-sheet ref)
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"

        # Col B: WP Name (green cross-sheet ref)
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        # Col C: % Time Elapsed
        time_formula = (
            f"MIN(1,(current_month-{SHEET_CONFIG}!D{config_row}+1)"
            f"/({SHEET_CONFIG}!E{config_row}-{SHEET_CONFIG}!D{config_row}+1))"
        )
        style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="0.0%").value = (
            f'=IF(current_month<{SHEET_CONFIG}!D{config_row},"\u2014",{time_formula})'
        )

        # Col D: % Budget EUR
        budget_formula = (
            f"{SHEET_ACTUALS}!{cum_col_letter}{actuals_row}"
            f"/{SHEET_ACTUALS}!{ceiling_col_letter}{actuals_row}"
        )
        style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="0.0%").value = (
            f'=IF(current_month<{SHEET_CONFIG}!D{config_row},"\u2014",{budget_formula})'
        )

        # Col E: Burn Delta
        style_cell(ws, row, 5, font=BLACK_FORMULA, fmt="0.0%").value = (
            f'=IF(current_month<{SHEET_CONFIG}!D{config_row},"\u2014",D{row}-C{row})'
        )

        # Col F: Status
        status_formula = (
            f'IF(ABS(E{row})<={TREND_OK_BAND},"OK",'
            f'IF(E{row}>{TREND_OVER_THRESHOLD},"OVER",'
            f'IF(E{row}>{TREND_WATCH_THRESHOLD},"WATCH",'
            f'IF(E{row}<{TREND_SLOW_THRESHOLD},"SLOW","OK"))))'
        )
        style_cell(ws, row, 6, font=BLACK_FORMULA).value = (
            f'=IF(current_month<{SHEET_CONFIG}!D{config_row},"\u2014",{status_formula})'
        )

        # Col G: Exhaustion Mo.
        exhaustion_formula = (
            f"IF(D{row}=0,\"\u2014\","
            f"IF(COUNTIF({actuals_month_range},\">0\")=0,\"No data\","
            f"{SHEET_CONFIG}!D{config_row}+ROUND("
            f"{SHEET_ACTUALS}!{ceiling_col_letter}{actuals_row}"
            f"/({SHEET_ACTUALS}!{cum_col_letter}{actuals_row}"
            f"/COUNTIF({actuals_month_range},\">0\")),0)-1))"
        )
        style_cell(ws, row, 7, font=BLACK_FORMULA, fmt="0").value = (
            f'=IF(current_month<{SHEET_CONFIG}!D{config_row},"\u2014",{exhaustion_formula})'
        )

    # -- Conditional formatting: Burn Delta (col E, rows 5-12) --
    e_range = f"E{first_wp_data_row}:E{last_wp_data_row}"
    e_ref = f"E{first_wp_data_row}"

    # Green first (lowest priority) — ABS(delta) <= TREND_OK_BAND
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'ABS({e_ref})<={TREND_OK_BAND}'], fill=_GREEN_CF_FILL,
    ))
    # Amber: overspend (delta > TREND_WATCH_THRESHOLD AND delta <= TREND_OVER_THRESHOLD)
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'AND({e_ref}>{TREND_WATCH_THRESHOLD},{e_ref}<={TREND_OVER_THRESHOLD})'], fill=_AMBER_CF_FILL,
    ))
    # Amber: underspend (delta < TREND_SLOW_THRESHOLD)
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'{e_ref}<{TREND_SLOW_THRESHOLD}'], fill=_AMBER_CF_FILL,
    ))
    # Red: delta > TREND_OVER_THRESHOLD
    ws.conditional_formatting.add(e_range, FormulaRule(
        formula=[f'{e_ref}>{TREND_OVER_THRESHOLD}'], fill=_RED_CF_FILL,
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
    for i in range(num_wps):
        row = first_wp_data_row + i
        config_row = layout["config_first_wp_row"] + i
        g_cell = f"G{row}:G{row}"
        g_ref = f"G{row}"
        # Green first: exhaustion >= wp_end
        ws.conditional_formatting.add(g_cell, FormulaRule(
            formula=[f'AND(ISNUMBER({g_ref}),{g_ref}>={SHEET_CONFIG}!E{config_row})'],
            fill=_GREEN_CF_FILL,
        ))
        # Amber: exhaustion within 3 months of wp_end
        ws.conditional_formatting.add(g_cell, FormulaRule(
            formula=[
                f'AND(ISNUMBER({g_ref}),'
                f'{g_ref}<{SHEET_CONFIG}!E{config_row},'
                f'{g_ref}>={SHEET_CONFIG}!E{config_row}-3)'
            ],
            fill=_AMBER_CF_FILL,
        ))
        # Red: exhaustion < wp_end - 3
        ws.conditional_formatting.add(g_cell, FormulaRule(
            formula=[
                f'AND(ISNUMBER({g_ref}),'
                f'{g_ref}<{SHEET_CONFIG}!E{config_row}-3)'
            ],
            fill=_RED_CF_FILL,
        ))

    # -- TOTAL row --
    total_row = last_wp_data_row + 1
    act_first = actuals_layout["first_wp_row"]
    style_cell(ws, total_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 2, "TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 3, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 4, "", fill=LIGHT_GREEN_FILL)
    # Weighted average burn delta — individual IF() terms avoid text*number #VALUE!
    terms = []
    for i in range(num_wps):
        wr = first_wp_data_row + i
        ar = act_first + i
        terms.append(
            f"IF(ISNUMBER(E{wr}),E{wr}*{SHEET_ACTUALS}!"
            f"{ceiling_col_letter}{ar},0)"
        )
    style_cell(ws, total_row, 5, font=BLACK_FORMULA,
               fill=LIGHT_GREEN_FILL, fmt="0.0%").value = (
        "=(" + "+".join(terms) + ")/partner_ceiling"
    )
    style_cell(ws, total_row, 6, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 7, "", fill=LIGHT_GREEN_FILL)

    # ================================================================
    # Section 2: Person Load Table
    # ================================================================
    # Chart is anchored 2 rows below total and occupies ~15 rows of height
    chart_anchor_row = total_row + 2
    person_section_row = chart_anchor_row + 16  # chart height ~15 rows
    write_section_header(ws, person_section_row, "PERSON LOAD", merge_end=5)

    # Column headers
    person_header_row = person_section_row + 1
    person_headers = [
        "Person", "Avg FTE (to date)", "Avg FTE (remaining)",
        "Peak Month FTE", "Status",
    ]
    for ci, h in enumerate(person_headers, 1):
        style_cell(ws, person_header_row, ci, h, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Person data rows
    person_first_row = person_header_row + 1
    person_groups = capacity_layout["person_groups"]
    person_rows_map = team_layout["person_rows"]

    for pi, member in enumerate(team):
        name = member["name"]
        row = person_first_row + pi
        tr = person_groups[name]["total_row"]
        roster_row = person_rows_map[name]

        # Col A: Person name (green cross-sheet ref)
        style_cell(ws, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{roster_row}"
        )

        # Col B: Avg FTE (to date)
        style_cell(ws, row, 2, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=IF(current_month=0,0,"
            f"SUMPRODUCT((COLUMN({SHEET_CAPACITY}!{ms}{tr}:{me}{tr})"
            f"-COLUMN({SHEET_CAPACITY}!${ms}$1)+1<=current_month)"
            f"*({SHEET_CAPACITY}!{ms}{tr}:{me}{tr}))"
            f"/current_month)"
        )

        # Col C: Avg FTE (remaining)
        style_cell(ws, row, 3, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=IF(current_month>={project_params['months']},0,"
            f"SUMPRODUCT((COLUMN({SHEET_CAPACITY}!{ms}{tr}:{me}{tr})"
            f"-COLUMN({SHEET_CAPACITY}!${ms}$1)+1>current_month)"
            f"*({SHEET_CAPACITY}!{ms}{tr}:{me}{tr}))"
            f"/({project_params['months']}-current_month))"
        )

        # Col D: Peak Month FTE
        style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=MAX({SHEET_CAPACITY}!{ms}{tr}:{me}{tr})"
        )

        # Col E: Status
        style_cell(ws, row, 5, font=BLACK_FORMULA).value = (
            f'=IF(C{row}>{SHEET_CONFIG}!B{roster_row},"OVER",'
            f'IF(C{row}>{CAPACITY_AMBER_LOAD}*{SHEET_CONFIG}!B{roster_row},"WATCH","OK"))'
        )

    person_last_row = person_first_row + len(team) - 1

    # -- Conditional formatting: Person Status (col E) + Peak (col D) — per person --
    # Each person has a different max_fte, so CF must use the correct roster row
    for pi, member in enumerate(team):
        row = person_first_row + pi
        roster_row = person_rows_map[member["name"]]
        e_cell = f"E{row}:E{row}"
        d_cell = f"D{row}:D{row}"

        # Status CF (green/amber/red based on col C vs max_fte)
        ws.conditional_formatting.add(e_cell, FormulaRule(
            formula=[f'C{row}<={CAPACITY_AMBER_LOAD}*{SHEET_CONFIG}!$B${roster_row}'],
            fill=_GREEN_CF_FILL,
        ))
        ws.conditional_formatting.add(e_cell, FormulaRule(
            formula=[f'C{row}>{CAPACITY_AMBER_LOAD}*{SHEET_CONFIG}!$B${roster_row}'],
            fill=_AMBER_CF_FILL,
        ))
        ws.conditional_formatting.add(e_cell, FormulaRule(
            formula=[f'C{row}>{SHEET_CONFIG}!$B${roster_row}'],
            fill=_RED_CF_FILL,
        ))

        # Peak Month CF (red if > max_fte)
        ws.conditional_formatting.add(d_cell, FormulaRule(
            formula=[f'D{row}>{SHEET_CONFIG}!$B${roster_row}'],
            fill=_RED_CF_FILL,
        ))

    # ================================================================
    # Section 3: Cumulative LineChart + hidden helper rows
    # ================================================================
    # Helper rows placed below person load section (2-row gap)
    helper_start_row = person_last_row + 2
    helper_row_planned = helper_start_row
    helper_row_actual = helper_start_row + 1
    helper_row_ceiling = helper_start_row + 2

    plan_first_wp = layout["planned_first_wp_row"]
    plan_last_wp = layout["planned_last_wp_row"]
    act_total = actuals_layout["total_row"]

    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        cl = col_letter(c)

        # Helper Row 1: Cumulative planned EUR
        if m_idx == 0:
            # First month: just the monthly total * fte_rate
            style_cell(ws, helper_row_planned, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"=SUM({SHEET_PLAN}!{cl}{plan_first_wp}:{cl}{plan_last_wp})*fte_rate"
            )
        else:
            prev_cl = col_letter(c - 1)
            style_cell(ws, helper_row_planned, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={prev_cl}{helper_row_planned}"
                f"+SUM({SHEET_PLAN}!{cl}{plan_first_wp}:{cl}{plan_last_wp})*fte_rate"
            )

        # Helper Row 2: Cumulative actual EUR
        if m_idx == 0:
            style_cell(ws, helper_row_actual, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={SHEET_ACTUALS}!{cl}{act_total}"
            )
        else:
            prev_cl = col_letter(c - 1)
            style_cell(ws, helper_row_actual, c,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"={prev_cl}{helper_row_actual}+{SHEET_ACTUALS}!{cl}{act_total}"
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
    # NOTE: Series are constructed manually (not via chart.add_data) because
    # LibreOffice's recalc macro strips val/tx/cat from add_data-created series.
    chart = LineChart()
    chart.title = "Cumulative EUR: Planned vs Actual"
    chart.style = 10
    chart.visible_cells_only = False  # helper rows are hidden
    chart.y_axis.numFmt = "\u20ac#,##0"
    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.delete = False

    sheet_name = "Trends"
    range_str = f"'{sheet_name}'!${ms}${{}}:${me}${{}}"

    # X-axis category labels (month-year, e.g. "Nov 25")
    helper_row_labels = helper_row_ceiling + 1
    project_start = project_params["start"]
    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        cal_date = project_start + relativedelta(months=m_idx)
        ws.cell(row=helper_row_labels, column=c,
                value=f"M{m_idx + 1}-{cal_date.strftime('%b%y')}")
    ws.row_dimensions[helper_row_labels].hidden = True

    # Helper Row 5: RP boundary markers (ceiling value at each RP end month)
    helper_row_rp = helper_row_labels + 1
    for rp in reporting_periods:
        boundary_col = month_col_start + rp["end"] - 1
        if boundary_col <= month_col_end:
            style_cell(ws, helper_row_rp, boundary_col,
                       font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                "=partner_ceiling"
            )
    ws.row_dimensions[helper_row_rp].hidden = True

    cat_ref = AxDataSource(strRef=StrRef(
        f=range_str.format(helper_row_labels, helper_row_labels)))

    # Series 1: Cumulative planned (blue solid)
    s_planned = Series(idx=0, order=0)
    s_planned.tx = SeriesLabel(v="Planned")
    s_planned.val = NumDataSource(numRef=NumRef(
        f=range_str.format(helper_row_planned, helper_row_planned)))
    s_planned.cat = cat_ref
    s_planned.graphicalProperties.line.solidFill = "4472C4"
    chart.series.append(s_planned)

    # Series 2: Cumulative actual (red solid)
    s_actual = Series(idx=1, order=1)
    s_actual.tx = SeriesLabel(v="Actual")
    s_actual.val = NumDataSource(numRef=NumRef(
        f=range_str.format(helper_row_actual, helper_row_actual)))
    s_actual.cat = cat_ref
    s_actual.graphicalProperties.line.solidFill = "FF0000"
    chart.series.append(s_actual)

    # Series 3: Ceiling (grey dashed)
    s_ceiling = Series(idx=2, order=2)
    s_ceiling.tx = SeriesLabel(v="Ceiling")
    s_ceiling.val = NumDataSource(numRef=NumRef(
        f=range_str.format(helper_row_ceiling, helper_row_ceiling)))
    s_ceiling.cat = cat_ref
    s_ceiling.graphicalProperties.line.solidFill = "808080"
    s_ceiling.graphicalProperties.line.dashStyle = "dash"
    chart.series.append(s_ceiling)

    # Series 4: RP boundaries (green diamond markers at ceiling level)
    s_rp = Series(idx=3, order=3)
    s_rp.tx = SeriesLabel(v="RP End")
    s_rp.val = NumDataSource(numRef=NumRef(
        f=range_str.format(helper_row_rp, helper_row_rp)))
    s_rp.cat = cat_ref
    s_rp.graphicalProperties.line.solidFill = "00B050"
    s_rp.graphicalProperties.line.dashStyle = "dot"
    rp_marker = Marker(symbol="diamond", size=10)
    assert rp_marker.graphicalProperties is not None
    rp_marker.graphicalProperties.solidFill = "00B050"
    s_rp.marker = rp_marker
    chart.series.append(s_rp)

    # Chart dimensions and placement
    chart.width = 30
    chart.height = 15
    ws.add_chart(chart, f"A{chart_anchor_row}")

    # ================================================================
    # Column widths
    # ================================================================
    auto_fit_columns(ws)
    ws.column_dimensions["B"].width = 25

    # ================================================================
    # Return layout dict
    # ================================================================
    return {
        "section_header_row": section1_row,
        "burn_rate_header_row": header_row,
        "burn_rate_first_wp_row": first_wp_data_row,
        "burn_rate_last_wp_row": last_wp_data_row,
        "burn_rate_total_row": total_row,
        "chart_anchor_row": chart_anchor_row,
        "person_load_section_row": person_section_row,
        "person_load_header_row": person_header_row,
        "person_load_first_row": person_first_row,
        "person_load_last_row": person_last_row,
        "helper_row_planned": helper_row_planned,
        "helper_row_actual": helper_row_actual,
        "helper_row_ceiling": helper_row_ceiling,
        "helper_row_labels": helper_row_labels,
        "helper_row_rp": helper_row_rp,
    }
