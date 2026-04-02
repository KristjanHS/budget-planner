"""Config sheet builder — project parameters and work package definitions."""

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

from budget_planner.styles import (
    BLUE_INPUT,
    HEADER_FONT,
    SECTION_FONT,
    SECTION_FILL,
    TITLE_FONT,
    LIGHT_BLUE_FILL,
    YELLOW_FILL,
    CENTER,
    style_cell,
)


def build_config_sheet(ws, wb, project_params, work_packages):
    """Build the Config sheet with project parameters and work packages table.

    Args:
        ws: Config worksheet to populate.
        wb: Parent Workbook (for named ranges).
        project_params: dict with keys 'title', 'start', 'months', 'rate'.
        work_packages: list of dicts with keys 'id', 'name', 'fte', 'start', 'end'.

    Returns:
        config_layout dict with keys:
            'param_rows': dict mapping param name to row number,
            'wp_header_row': int,
            'wp_first_row': int,
            'wp_last_row': int,
    """
    # ── Section A: Project Parameters (rows 1-4) ──
    param_defs = [
        ("Project Title", project_params["title"], None, None),
        ("Project Start", project_params["start"], "MMM YYYY", None),
        ("Project Months", project_params["months"], "0", None),
        ("FTE Rate (EUR)", project_params["rate"], "€#,##0", None),
    ]
    param_rows = {}
    for i, (label, value, fmt, _) in enumerate(param_defs, start=1):
        # Label in column A — HEADER_FONT, no border
        style_cell(ws, i, 1, value=label, font=HEADER_FONT, border=False)
        # Value in column B — BLUE_INPUT, no border
        font = TITLE_FONT if i == 1 else BLUE_INPUT
        style_cell(ws, i, 2, value=value, font=font, fmt=fmt, border=False)
        param_rows[label] = i

    # Row 5 is empty separator (no action needed)

    # ── Section B: Work Packages Table (rows 6+) ──
    section_row = 6
    header_row = 7
    first_data_row = 8

    # Section header — merged A6:E6
    ws.merge_cells(start_row=section_row, start_column=1,
                   end_row=section_row, end_column=5)
    style_cell(ws, section_row, 1, value="WORK PACKAGES",
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)

    # Column headers
    headers = ["WP ID", "WP Name", "FTE Budget", "Start Month", "End Month"]
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

    # Excel Table over header + data range
    table_ref = f"A{header_row}:E{last_data_row}"
    table_style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=False,
    )
    table = Table(displayName="tbl_WorkPackages", ref=table_ref)
    table.tableStyleInfo = table_style
    ws.add_table(table)

    # Conditional formatting: yellow fill if Start > End
    for row_idx in range(first_data_row, last_data_row + 1):
        rule = FormulaRule(
            formula=[f"$D{row_idx}>$E{row_idx}"],
            fill=YELLOW_FILL,
        )
        ws.conditional_formatting.add(f"A{row_idx}:E{row_idx}", rule)

    # ── Named Ranges ──
    dn_start = DefinedName("project_start", attr_text="Config!$B$2")
    wb.defined_names.add(dn_start)

    dn_rate = DefinedName("fte_rate", attr_text="Config!$B$4")
    wb.defined_names.add(dn_rate)

    # ── Sheet Appearance ──
    ws.sheet_properties.tabColor = "4472C4"
    ws.freeze_panes = "A6"

    return {
        "param_rows": param_rows,
        "wp_header_row": header_row,
        "wp_first_row": first_data_row,
        "wp_last_row": last_data_row,
    }
