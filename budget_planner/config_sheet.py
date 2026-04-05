"""Config sheet builder — project parameters, work packages, and team roster."""

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import SHEET_CONFIG
from budget_planner.styles import (
    BLACK_FORMULA,
    BLUE_INPUT,
    HEADER_FONT,
    SECTION_FONT,
    SECTION_FILL,
    TITLE_FONT,
    LIGHT_BLUE_FILL,
    auto_fit_columns,
    CF_YELLOW_FILL,
    CENTER,
    style_cell,
)


def build_config_sheet(ws, wb, project_params, work_packages, team=None,
                       reporting_periods=None):
    """Build the Config sheet with project parameters, work packages table, and team roster.

    Args:
        ws: Config worksheet to populate.
        wb: Parent Workbook (for named ranges).
        project_params: dict with keys 'title', 'start', 'months', 'rate'.
        work_packages: list of dicts with keys 'id', 'name', 'fte', 'start', 'end'.
        team: list of dicts with keys 'name', 'max_fte'. When None, lazy-imports from config.
        reporting_periods: list of dicts with keys 'id', 'start', 'end'.
            When None, lazy-imports from config.

    Returns:
        config_layout dict with keys:
            'param_rows': dict mapping param name to row number,
            'wp_header_row': int,
            'wp_first_row': int,
            'wp_last_row': int,
            'wp_total_row': int,
            'roster_header_row': int,
            'roster_first_row': int,
            'roster_last_row': int,
            'person_rows': dict mapping person name to row number,
            'rp_first_row': int,
            'rp_last_row': int,
    """
    if team is None:
        from budget_planner import config
        team = config.TEAM
    # ── Section A: Project Parameters (rows 1-4) ──
    param_defs = [
        ("Project Title", project_params["title"], None, None),
        ("Project Start", project_params["start"], "MMM YYYY", None),
        ("Project Months", project_params["months"], "0", None),
        ("FTE Rate (EUR)", project_params["rate"], "€#,##0", None),
        ("EU Annual Ceiling (PM)", project_params.get("eu_ceiling_pm", 12.0), "0.00", None),
    ]
    param_rows = {}
    for i, (label, value, fmt, _) in enumerate(param_defs, start=1):
        # Label in column A — HEADER_FONT, no border
        style_cell(ws, i, 1, value=label, font=HEADER_FONT, border=False)
        # Value in column B — BLUE_INPUT, no border
        font = TITLE_FONT if i == 1 else BLUE_INPUT
        style_cell(ws, i, 2, value=value, font=font, fmt=fmt, border=False)
        param_rows[label] = i

    # Empty separator row after last param (no action needed)

    # ── Section B: Work Packages Table ──
    section_row = len(param_defs) + 2   # gap row after last param
    header_row = section_row + 1
    first_data_row = header_row + 1

    # Section header — merged A6:H6
    ws.merge_cells(start_row=section_row, start_column=1,
                   end_row=section_row, end_column=8)
    style_cell(ws, section_row, 1, value="WORK PACKAGES",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)

    # Column headers
    headers = [
        "WP ID", "WP Name", "FTE Budget", "Start Month", "End Month",
        "Type", "Active Months", "Personnel Ceiling (\u20ac)",
    ]
    for col, hdr in enumerate(headers, start=1):
        style_cell(ws, header_row, col, value=hdr,
                   font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # WP data rows
    wp_formats = [None, None, "0.00", "0", "0"]
    last_data_row = first_data_row + len(work_packages) - 1
    for row_idx, wp in enumerate(work_packages, start=first_data_row):
        values = [wp["id"], wp["name"], wp["fte"], wp["start"], wp["end"]]
        for col, (val, fmt) in enumerate(zip(values, wp_formats), start=1):
            style_cell(ws, row_idx, col, value=val, font=BLUE_INPUT, fmt=fmt,
                       border=False)
        # Col F: Type (blue input)
        wp_type = wp.get("type", "flat")
        style_cell(ws, row_idx, 6, value=wp_type, font=BLUE_INPUT, border=False)
        # Col G: Active Months = End - Start + 1 (black formula)
        style_cell(ws, row_idx, 7, font=BLACK_FORMULA, fmt="0",
                   border=False).value = f"=E{row_idx}-D{row_idx}+1"
        # Col H: Personnel Ceiling = FTE Budget * fte_rate (black formula)
        style_cell(ws, row_idx, 8, font=BLACK_FORMULA, fmt="\u20ac#,##0",
                   border=False).value = f"=C{row_idx}*fte_rate"

    # TOTAL row
    wp_total_row = last_data_row + 1
    style_cell(ws, wp_total_row, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, wp_total_row, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)
    style_cell(ws, wp_total_row, 3, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
               fmt="0.00").value = (
        f"=SUM(C{first_data_row}:C{last_data_row})"
    )
    style_cell(ws, wp_total_row, 4, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, wp_total_row, 5, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, wp_total_row, 6, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, wp_total_row, 7, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, wp_total_row, 8, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
               fmt="\u20ac#,##0").value = (
        f"=SUM(H{first_data_row}:H{last_data_row})"
    )

    # Excel Table over header + data range (includes derived columns G-H)
    table_ref = f"A{header_row}:H{last_data_row}"
    table_style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=False, showColumnStripes=False,
    )
    table = Table(displayName="tbl_WorkPackages", ref=table_ref)
    table.tableStyleInfo = table_style
    ws.add_table(table)

    # Conditional formatting: yellow fill if Start > End
    for row_idx in range(first_data_row, last_data_row + 1):
        rule = FormulaRule(
            formula=[f"$D{row_idx}>$E{row_idx}"],
            fill=CF_YELLOW_FILL,
        )
        ws.conditional_formatting.add(f"A{row_idx}:H{row_idx}", rule)

    # ── Section C: Team Roster (below WP table + TOTAL row) ──
    r = wp_total_row + 2  # blank separator row after TOTAL row

    # Section header — merged A:B
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, value="TEAM ROSTER",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    style_cell(ws, r, 2, fill=SECTION_FILL)
    r += 1

    # Column headers
    roster_header_row = r
    style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "Max FTE", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    r += 1

    # Team member data rows
    roster_first_row = r
    person_rows: dict[str, int] = {}
    for member in team:
        style_cell(ws, r, 1, member["name"], font=BLUE_INPUT)
        style_cell(ws, r, 2, member["max_fte"], font=BLUE_INPUT, fmt="0.00")
        person_rows[member["name"]] = r
        r += 1
    roster_last_row = r - 1

    # ── Section D: Reporting Periods ──
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS

    r = roster_last_row + 2  # gap row after roster

    # Section header — merged A:C
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    style_cell(ws, r, 1, value="REPORTING PERIODS",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    for c in range(2, 4):
        style_cell(ws, r, c, fill=SECTION_FILL)
    r += 1

    # Column headers
    for col, hdr in enumerate(["Period", "Start Month", "End Month"], start=1):
        style_cell(ws, r, col, hdr, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    # RP data rows
    rp_first_row = r
    for rp in reporting_periods:
        style_cell(ws, r, 1, rp["id"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, 2, rp["start"], font=BLUE_INPUT, fmt="0", border=False)
        style_cell(ws, r, 3, rp["end"], font=BLUE_INPUT, fmt="0", border=False)
        r += 1
    rp_last_row = r - 1

    # ── Named Ranges ──
    start_row = param_rows["Project Start"]
    dn_start = DefinedName("project_start", attr_text=f"{SHEET_CONFIG}!$B${start_row}")
    wb.defined_names.add(dn_start)

    rate_row = param_rows["FTE Rate (EUR)"]
    dn_rate = DefinedName("fte_rate", attr_text=f"{SHEET_CONFIG}!$B${rate_row}")
    wb.defined_names.add(dn_rate)

    ceiling_row = param_rows["EU Annual Ceiling (PM)"]
    dn_ceiling = DefinedName("eu_annual_ceiling_pm",
                              attr_text=f"{SHEET_CONFIG}!$B${ceiling_row}")
    wb.defined_names.add(dn_ceiling)

    # ── Sheet Appearance ──
    ws.sheet_properties.tabColor = "4472C4"
    ws.freeze_panes = f"A{section_row}"
    auto_fit_columns(ws)

    return {
        "param_rows": param_rows,
        "wp_header_row": header_row,
        "wp_first_row": first_data_row,
        "wp_last_row": last_data_row,
        "wp_total_row": wp_total_row,
        "roster_header_row": roster_header_row,
        "roster_first_row": roster_first_row,
        "roster_last_row": roster_last_row,
        "person_rows": person_rows,
        "rp_first_row": rp_first_row,
        "rp_last_row": rp_last_row,
    }
