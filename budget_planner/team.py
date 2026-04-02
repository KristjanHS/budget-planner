"""Team sheet builder — roster and WP share seed matrix."""

from openpyxl.formatting.rule import FormulaRule

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, LIGHT_BLUE_FILL, YELLOW_FILL,
    CENTER, style_cell, col_letter,
)


def build_team_sheet(ws, wb, *, team=None, work_packages=None):
    """Build the Team sheet: roster + WP share seed matrix.

    Args:
        ws: The Team worksheet to populate.
        wb: The parent Workbook (needed for potential named ranges).
        team: List of team member dicts (lazy-imports from config if None).
        work_packages: List of WP dicts (lazy-imports from config if None).

    Returns:
        team_layout dict with row/column positions for all key elements.
    """
    if team is None or work_packages is None:
        from budget_planner import config
        if team is None:
            team = config.TEAM
        if work_packages is None:
            work_packages = config.WORK_PACKAGES

    ws.sheet_properties.tabColor = "FFA500"
    r = 1

    # ── BLOCK 1: TEAM ROSTER ──
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, "TEAM ROSTER", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, fill=SECTION_FILL)
    r += 1

    roster_header_row = r
    style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "Max FTE", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    r += 1

    roster_first_row = r
    person_rows = {}
    for member in team:
        style_cell(ws, r, 1, member["name"], font=BLUE_INPUT)
        style_cell(ws, r, 2, member["max_fte"], font=BLUE_INPUT, fmt="0.00")
        person_rows[member["name"]] = r
        r += 1
    roster_last_row = r - 1

    # Gap row
    r += 1

    # ── BLOCK 2: WP SHARE SEEDS ──
    ws.merge_cells(start_row=r, start_column=1, end_row=r,
                   end_column=1 + len(work_packages))
    style_cell(ws, r, 1, "WP SHARE SEEDS", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, 2 + len(work_packages)):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1

    share_header_row = r
    style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    share_wp_columns = {}
    for wi, wp in enumerate(work_packages):
        col = 2 + wi
        share_wp_columns[wp["id"]] = col
        style_cell(ws, r, col, wp["id"], font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    share_first_row = r
    for member in team:
        roster_row = person_rows[member["name"]]
        # Person name as formula ref to roster
        style_cell(ws, r, 1, font=BLUE_INPUT).value = f"=A{roster_row}"
        for wp in work_packages:
            col = share_wp_columns[wp["id"]]
            share_val = member["wps"].get(wp["id"])
            if share_val is not None:
                style_cell(ws, r, col, share_val, font=BLUE_INPUT, fmt="0.00")
            else:
                style_cell(ws, r, col, font=BLUE_INPUT, fmt="0.00")
        r += 1
    share_last_row = r - 1

    # Validation row: SUM per WP column
    share_validation_row = r
    style_cell(ws, r, 1, "SUM", font=HEADER_FONT)
    for wp in work_packages:
        col = share_wp_columns[wp["id"]]
        cl = col_letter(col)
        style_cell(ws, r, col, font=BLACK_FORMULA, fmt="0.00").value = (
            f"=SUM({cl}{share_first_row}:{cl}{share_last_row})"
        )
        # Conditional formatting: yellow if SUM != 1.0
        cell_ref = f"{cl}{share_validation_row}"
        ws.conditional_formatting.add(cell_ref, FormulaRule(
            formula=[f"ABS({cell_ref}-1)>0.001"], fill=YELLOW_FILL,
        ))

    # Freeze panes: freeze header rows
    ws.freeze_panes = "A3"

    return {
        "roster_header_row": roster_header_row,
        "roster_first_row": roster_first_row,
        "roster_last_row": roster_last_row,
        "share_header_row": share_header_row,
        "share_first_row": share_first_row,
        "share_last_row": share_last_row,
        "share_validation_row": share_validation_row,
        "share_wp_columns": share_wp_columns,
        "person_rows": person_rows,
    }
