"""Team sheet builder — roster and WP share seed matrix."""

from openpyxl.formatting.rule import FormulaRule

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, LIGHT_BLUE_FILL, YELLOW_FILL,
    CENTER, style_cell, col_letter,
)
from budget_planner.config import TEAM, WORK_PACKAGES


def build_team_sheet(ws, wb, team_data=None):
    """Build the Team sheet: roster + WP share seed matrix.

    Args:
        ws: The Team worksheet to populate.
        wb: The parent Workbook (needed for potential named ranges).
        team_data: Optional preserved team data from reader. If provided,
                   uses this instead of config.py TEAM. Format:
                   [{"name": "Alice", "max_fte": 1.0, "wps": {"WP1": 0.25}}, ...]

    Returns:
        team_layout dict with row/column positions for all key elements.
    """
    team = team_data if team_data is not None else TEAM

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
                   end_column=1 + len(WORK_PACKAGES))
    style_cell(ws, r, 1, "WP SHARE SEEDS", font=SECTION_FONT, fill=SECTION_FILL,
               alignment=CENTER)
    for c2 in range(2, 2 + len(WORK_PACKAGES)):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1

    share_header_row = r
    style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    share_wp_columns = {}
    for wi, wp in enumerate(WORK_PACKAGES):
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
        for wp in WORK_PACKAGES:
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
    for wp in WORK_PACKAGES:
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

    # ── COLUMN WIDTHS ──
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10
    for wp in WORK_PACKAGES:
        col = share_wp_columns[wp["id"]]
        ws.column_dimensions[col_letter(col)].width = 10

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
