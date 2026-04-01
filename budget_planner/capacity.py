"""Capacity sheet builder — per-person monthly FTE allocation."""

from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from dateutil.relativedelta import relativedelta

from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    AMBER_FILL, RED_FILL, CENTER, style_cell, col_letter,
)
from budget_planner.config import (
    PROJECT_START, PROJECT_MONTHS, WORK_PACKAGES,
    TEAM, EXTERNAL_COMMITMENTS,
    MONTH_COL_START, MONTH_COL_END,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
)


def _wp_config_row(wp_id, config_first_wp_row=6):
    """Return the Plan config table row for a WP ID."""
    for i, wp in enumerate(WORK_PACKAGES):
        if wp["id"] == wp_id:
            return config_first_wp_row + i
    raise ValueError(f"Unknown WP ID: {wp_id}")


def build_capacity_sheet(ws, wb, team_layout, preserved_data=None, mode="demo"):
    """Build the Capacity sheet: external commitments config + per-person groups.

    Args:
        ws: The Capacity worksheet to populate.
        wb: The parent Workbook.
        team_layout: dict from build_team_sheet with person_rows, share_wp_columns, etc.
        preserved_data: Optional dict with capacity data to preserve on --update.
        mode: "demo" seeds monthly cells with calculated values;
              "skeleton" leaves them empty for manual entry.

    Returns:
        capacity_layout dict with row/column positions for all key elements.
    """
    ws.sheet_properties.tabColor = "7030A0"
    r = 1

    person_rows = team_layout["person_rows"]
    share_wp_columns = team_layout["share_wp_columns"]
    share_first_row = team_layout["share_first_row"]

    # ── CONFIG BLOCK: EXTERNAL COMMITMENTS ──
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    style_cell(ws, r, 1, "EXTERNAL COMMITMENTS", font=SECTION_FONT,
               fill=SECTION_FILL, alignment=CENTER)
    for c2 in range(2, 6):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1

    config_col_headers = ["Person", "Project", "Total FTE", "Start Month", "End Month"]
    for ci, h in enumerate(config_col_headers, 1):
        style_cell(ws, r, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    config_first_row = r
    external_config_rows = {}
    for entry in EXTERNAL_COMMITMENTS:
        style_cell(ws, r, 1, entry["person"], font=BLUE_INPUT)
        style_cell(ws, r, 2, entry["project"], font=BLUE_INPUT)
        style_cell(ws, r, 3, entry["fte"], font=BLUE_INPUT, fmt="0.00")
        style_cell(ws, r, 4, entry["start"], font=BLUE_INPUT, fmt="0")
        style_cell(ws, r, 5, entry["end"], font=BLUE_INPUT, fmt="0")
        external_config_rows[(entry["person"], entry["project"])] = r
        r += 1
    config_last_row = r - 1 if EXTERNAL_COMMITMENTS else r

    # Gap row
    r += 1

    # ── YEAR + MONTH HEADER ROWS (once, before first person group) ──
    year_header_row = r
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
    style_cell(ws, r, VAL_COL_BUDGET, "Budget", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_DIFF, "Variance", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1

    month_header_row = r
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    for mi in range(PROJECT_MONTHS):
        c = MONTH_COL_START + mi
        cal_date = PROJECT_START + relativedelta(months=mi)
        mlabel = f"M{mi+1}\n{cal_date.strftime('%b %y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8),
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_SUM, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_BUDGET, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_DIFF, "", fill=LIGHT_BLUE_FILL)
    r += 1

    first_person_header_row = r

    # ── PER-PERSON GROUPS ──
    person_groups = {}

    for member in TEAM:
        name = member["name"]
        roster_row = person_rows[name]

        # Person header row
        header_row = r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        style_cell(ws, r, 1, font=HEADER_FONT, fill=SECTION_FILL)
        ws.cell(row=r, column=1).value = f"=Team!A{roster_row}"
        ws.cell(row=r, column=1).font = Font(name="Arial", bold=True, size=10,
                                              color="FFFFFF")
        style_cell(ws, r, 2, fill=SECTION_FILL)
        r += 1

        # Track all data rows for this person (for TOTAL formula)
        data_row_start = r
        wp_rows = {}
        external_rows = {}

        # Programme WP rows
        # Find the person's index in TEAM for share matrix row lookup
        person_idx = next(i for i, m in enumerate(TEAM) if m["name"] == name)
        share_row = share_first_row + person_idx

        # Look up preserved capacity data for this person
        person_preserved = None
        if preserved_data and name in preserved_data.get("person_groups", {}):
            person_preserved = preserved_data["person_groups"][name]

        first_wp_row_for_person = True
        for wp_id, share in member["wps"].items():
            wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
            config_row = _wp_config_row(wp_id)
            share_col = share_wp_columns[wp_id]

            # Col A: formula ref to Plan config WP ID
            style_cell(ws, r, 1, font=GREEN_XREF).value = f"=Plan!A{config_row}"
            # Col B: formula ref to Plan config WP Name
            style_cell(ws, r, 2, font=GREEN_XREF).value = f"=Plan!B{config_row}"

            # Monthly cells
            active_months = wp["end"] - wp["start"] + 1
            if mode == "skeleton":
                seed_value = None
            else:
                seed_value = round(share * wp["fte"] / active_months, 4) if active_months > 0 else 0

            # Get preserved WP values for this person-WP pair
            wp_preserved = None
            if person_preserved is not None:
                wp_preserved = person_preserved.get("wp_values", {}).get(wp_id)

            comment_added = False
            for mi in range(PROJECT_MONTHS):
                c = MONTH_COL_START + mi
                month_num = mi + 1

                # Check for preserved value
                cap_preserved = None
                if wp_preserved and mi < len(wp_preserved) and wp_preserved[mi] is not None:
                    cap_preserved = wp_preserved[mi]

                if wp["start"] <= month_num <= wp["end"]:
                    if cap_preserved is not None:
                        val = cap_preserved
                    elif seed_value is not None:
                        val = seed_value
                    else:
                        val = None

                    if val is not None:
                        style_cell(ws, r, c, val, font=BLUE_INPUT, fmt="0.00")
                    else:
                        style_cell(ws, r, c, font=BLUE_INPUT, fmt="0.00")
                    # Add comment on first active cell of first WP row per person
                    if (mode == "skeleton" and first_wp_row_for_person
                            and not comment_added):
                        ws.cell(row=r, column=c).comment = Comment(
                            "Enter monthly FTE allocation for each work "
                            "package and external project.",
                            "Programme Budget Tool",
                        )
                        comment_added = True
                else:
                    style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                               fmt="0.00")

            # Validation columns
            sum_range = (f"{col_letter(MONTH_COL_START)}{r}:"
                         f"{col_letter(MONTH_COL_END)}{r}")
            style_cell(ws, r, VAL_COL_SUM, font=BLACK_FORMULA,
                       fmt="0.00").value = f"=SUM({sum_range})"

            # Budget: person's share x WP total FTE
            share_cl = col_letter(share_col)
            style_cell(ws, r, VAL_COL_BUDGET, font=GREEN_XREF,
                       fmt="0.00").value = (
                f"=Team!{share_cl}{share_row}*Plan!C{config_row}"
            )

            # Variance: Budget - Sum
            style_cell(ws, r, VAL_COL_DIFF, font=BLACK_FORMULA,
                       fmt="0.00").value = (
                f"={col_letter(VAL_COL_BUDGET)}{r}"
                f"-{col_letter(VAL_COL_SUM)}{r}"
            )

            wp_rows[wp_id] = r
            first_wp_row_for_person = False
            r += 1

        # External project rows
        person_externals = [e for e in EXTERNAL_COMMITMENTS
                            if e["person"] == name]
        for entry in person_externals:
            ext_config_row = external_config_rows[(name, entry["project"])]

            # Col A: project name (blue input)
            style_cell(ws, r, 1, entry["project"], font=BLUE_INPUT)
            # Col B: empty
            style_cell(ws, r, 2, font=BLUE_INPUT)

            # Monthly cells
            active_months = entry["end"] - entry["start"] + 1
            if mode == "skeleton":
                ext_seed = None
            else:
                ext_seed = round(entry["fte"] / active_months, 4) if active_months > 0 else 0

            # Get preserved external values for this person-project pair
            ext_preserved = None
            if person_preserved is not None:
                ext_preserved = person_preserved.get("external_values", {}).get(entry["project"])

            for mi in range(PROJECT_MONTHS):
                c = MONTH_COL_START + mi
                month_num = mi + 1

                # Check for preserved value
                ext_cap_preserved = None
                if ext_preserved and mi < len(ext_preserved) and ext_preserved[mi] is not None:
                    ext_cap_preserved = ext_preserved[mi]

                if entry["start"] <= month_num <= entry["end"]:
                    if ext_cap_preserved is not None:
                        val = ext_cap_preserved
                    elif ext_seed is not None:
                        val = ext_seed
                    else:
                        val = None

                    if val is not None:
                        style_cell(ws, r, c, val, font=BLUE_INPUT, fmt="0.00")
                    else:
                        style_cell(ws, r, c, font=BLUE_INPUT, fmt="0.00")
                else:
                    style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                               fmt="0.00")

            # Validation columns
            sum_range = (f"{col_letter(MONTH_COL_START)}{r}:"
                         f"{col_letter(MONTH_COL_END)}{r}")
            style_cell(ws, r, VAL_COL_SUM, font=BLACK_FORMULA,
                       fmt="0.00").value = f"=SUM({sum_range})"

            # Budget: ref to config block Total FTE cell
            style_cell(ws, r, VAL_COL_BUDGET, font=GREEN_XREF,
                       fmt="0.00").value = f"=C{ext_config_row}"

            # Variance: Budget - Sum
            style_cell(ws, r, VAL_COL_DIFF, font=BLACK_FORMULA,
                       fmt="0.00").value = (
                f"={col_letter(VAL_COL_BUDGET)}{r}"
                f"-{col_letter(VAL_COL_SUM)}{r}"
            )

            external_rows[entry["project"]] = r
            r += 1

        data_row_end = r - 1

        # TOTAL row
        total_row = r
        style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
        style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
        for mi in range(PROJECT_MONTHS):
            c = MONTH_COL_START + mi
            cl = col_letter(c)
            style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                       fmt="0.00").value = (
                f"=SUM({cl}{data_row_start}:{cl}{data_row_end})"
            )
        r += 1

        # REMAINING row
        remaining_row = r
        style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
        style_cell(ws, r, 2, "REMAINING", font=HEADER_FONT,
                   fill=LIGHT_GREEN_FILL)
        for mi in range(PROJECT_MONTHS):
            c = MONTH_COL_START + mi
            cl = col_letter(c)
            style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                       fmt="0.00").value = (
                f"=Team!B{roster_row}-{cl}{total_row}"
            )
        r += 1

        # Conditional formatting for TOTAL row
        total_range = (f"{col_letter(MONTH_COL_START)}{total_row}:"
                       f"{col_letter(MONTH_COL_END)}{total_row}")
        # Amber if > 0.8 x Max FTE (added first, lower priority)
        ws.conditional_formatting.add(total_range, FormulaRule(
            formula=[f"{col_letter(MONTH_COL_START)}{total_row}>0.8*Team!$B${roster_row}"],
            fill=AMBER_FILL,
        ))
        # Red if > Max FTE (added second, overrides amber for overcommitted cells)
        ws.conditional_formatting.add(total_range, FormulaRule(
            formula=[f"{col_letter(MONTH_COL_START)}{total_row}>Team!$B${roster_row}"],
            fill=RED_FILL,
        ))

        # Conditional formatting for REMAINING row — red if < 0
        remaining_range = (f"{col_letter(MONTH_COL_START)}{remaining_row}:"
                           f"{col_letter(MONTH_COL_END)}{remaining_row}")
        ws.conditional_formatting.add(remaining_range, FormulaRule(
            formula=[f"{col_letter(MONTH_COL_START)}{remaining_row}<0"],
            fill=RED_FILL,
        ))

        person_groups[name] = {
            "header_row": header_row,
            "wp_rows": wp_rows,
            "external_rows": external_rows,
            "total_row": total_row,
            "remaining_row": remaining_row,
        }

    # Freeze panes
    ws.freeze_panes = f"C{first_person_header_row + 1}"

    return {
        "config_first_row": config_first_row,
        "config_last_row": config_last_row,
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "person_groups": person_groups,
        "external_config_rows": external_config_rows,
    }
