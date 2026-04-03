"""Capacity sheet builder — per-person monthly FTE allocation."""

from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG, CAPACITY_AMBER_LOAD,
    build_year_month_headers, month_col_end as _month_col_end, val_columns,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FONT, SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    YELLOW_FILL, AMBER_FILL, RED_FILL, CENTER, style_cell, col_letter,
)


def _wp_config_row(wp_id, work_packages, config_first_wp_row):
    """Return the Plan config table row for a WP ID."""
    for i, wp in enumerate(work_packages):
        if wp["id"] == wp_id:
            return config_first_wp_row + i
    raise ValueError(f"Unknown WP ID: {wp_id}")


def build_capacity_sheet(ws, wb, team_layout, *,
                         project_params=None, work_packages=None,
                         team=None, external_commitments=None,
                         preserved=None, config_first_wp_row=None):
    """Build the Capacity sheet: external commitments config + per-person groups.

    Monthly FTE cells are always left empty for manual entry (or demo seeding).
    When ``preserved`` is provided (rebuild), those values are restored.

    Args:
        ws: The Capacity worksheet to populate.
        wb: The parent Workbook.
        team_layout: dict from build_config_sheet with person_rows, etc.
        project_params: dict with keys title, start, months, rate.
            When None, lazy-imports from config.
        work_packages: list of WP dicts. When None, lazy-imports from config.
        team: list of team member dicts. When None, lazy-imports from config.
        external_commitments: list of external commitment dicts.
            When None, lazy-imports from config.
        preserved: dict keyed by (person_name, wp_id) -> list of monthly
            float values for round-trip rebuild mode.
        config_first_wp_row: First WP row in Plan config table. Required.

    Returns:
        capacity_layout dict with row/column positions for all key elements.
    """
    # ── Lazy-import fallback for backward compatibility ──
    if any(v is None for v in [project_params, work_packages, team, external_commitments]):
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
        if team is None:
            team = config.TEAM
        if external_commitments is None:
            external_commitments = config.EXTERNAL_COMMITMENTS

    assert project_params is not None
    assert work_packages is not None
    assert team is not None
    assert external_commitments is not None
    assert config_first_wp_row is not None

    month_col_start = MONTH_COL_START
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]

    ws.sheet_properties.tabColor = "7030A0"
    r = 1

    person_rows = team_layout["person_rows"]

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
    for entry in external_commitments:
        style_cell(ws, r, 1, entry["person"], font=BLUE_INPUT)
        style_cell(ws, r, 2, entry["project"], font=BLUE_INPUT)
        style_cell(ws, r, 3, entry["fte"], font=BLUE_INPUT, fmt="0.00")
        style_cell(ws, r, 4, entry["start"], font=BLUE_INPUT, fmt="0")
        style_cell(ws, r, 5, entry["end"], font=BLUE_INPUT, fmt="0")
        external_config_rows[(entry["person"], entry["project"])] = r
        r += 1
    config_last_row = r - 1 if external_commitments else r

    # Gap row
    r += 1

    # ── YEAR + MONTH HEADER ROWS (once, before first person group) ──
    year_header_row, month_header_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            (val_col_sum, "Sum"),
            (val_col_budget, "Budget"),
            (val_col_diff, "Variance"),
        ],
    )

    first_person_header_row = r

    # ── PER-PERSON GROUPS ──
    person_groups = {}

    for member in team:
        name = member["name"]
        roster_row = person_rows[name]

        # Person header row
        header_row = r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        style_cell(ws, r, 1, font=HEADER_FONT, fill=SECTION_FILL)
        ws.cell(row=r, column=1).value = f"={SHEET_CONFIG}!A{roster_row}"
        ws.cell(row=r, column=1).font = Font(name="Arial", bold=True, size=10,
                                              color="FFFFFF")
        style_cell(ws, r, 2, fill=SECTION_FILL)
        r += 1

        # Track all data rows for this person (for TOTAL formula)
        data_row_start = r
        wp_rows = {}
        external_rows = {}

        # Programme WP rows — one row per WP for every person
        first_wp_row_for_person = True
        for wp in work_packages:
            wp_id = wp["id"]
            config_row = _wp_config_row(wp_id, work_packages, config_first_wp_row)

            # Col A: formula ref to Config WP ID
            style_cell(ws, r, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
            # Col B: formula ref to Config WP Name
            style_cell(ws, r, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

            # Monthly cells
            # Check for preserved round-trip data
            if preserved and (name, wp_id) in preserved:
                monthly_values = preserved[(name, wp_id)]
            else:
                monthly_values = None

            comment_added = False
            for mi in range(project_params["months"]):
                c = month_col_start + mi
                month_num = mi + 1

                if wp["start"] <= month_num <= wp["end"]:
                    if monthly_values is not None and mi < len(monthly_values) and monthly_values[mi] is not None:
                        seed_value = monthly_values[mi]
                    else:
                        seed_value = None

                    if seed_value is not None:
                        style_cell(ws, r, c, seed_value, font=BLUE_INPUT, fmt="0.00")
                    else:
                        style_cell(ws, r, c, font=BLUE_INPUT, fmt="0.00")
                    # Add comment on first active cell of first WP row per person
                    if (first_wp_row_for_person and not comment_added
                            and seed_value is None):
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
            sum_range = (f"{col_letter(month_col_start)}{r}:"
                         f"{col_letter(month_col_end)}{r}")
            style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                       fmt="0.00").value = f"=SUM({sum_range})"

            # Budget: WP total FTE from Config WP table
            style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                       fmt="0.00").value = (
                f"={SHEET_CONFIG}!C{config_row}"
            )

            # Variance: Budget - Sum
            style_cell(ws, r, val_col_diff, font=BLACK_FORMULA,
                       fmt="0.00").value = (
                f"={col_letter(val_col_budget)}{r}"
                f"-{col_letter(val_col_sum)}{r}"
            )

            # CF: yellow if Sum > Budget (variance < 0)
            diff_cl = col_letter(val_col_diff)
            ws.conditional_formatting.add(
                f"{diff_cl}{r}:{diff_cl}{r}",
                FormulaRule(formula=[f"{diff_cl}{r}<0"], fill=YELLOW_FILL),
            )

            wp_rows[wp_id] = r
            first_wp_row_for_person = False
            r += 1

        # External project rows
        person_externals = [e for e in external_commitments
                            if e["person"] == name]
        for entry in person_externals:
            ext_config_row = external_config_rows[(name, entry["project"])]

            # Col A: project name (blue input)
            style_cell(ws, r, 1, entry["project"], font=BLUE_INPUT)
            # Col B: empty
            style_cell(ws, r, 2, font=BLUE_INPUT)

            # Monthly cells — always empty for manual entry
            for mi in range(project_params["months"]):
                c = month_col_start + mi
                month_num = mi + 1

                if entry["start"] <= month_num <= entry["end"]:
                    style_cell(ws, r, c, font=BLUE_INPUT, fmt="0.00")
                else:
                    style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                               fmt="0.00")

            # Validation columns
            sum_range = (f"{col_letter(month_col_start)}{r}:"
                         f"{col_letter(month_col_end)}{r}")
            style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                       fmt="0.00").value = f"=SUM({sum_range})"

            # Budget: ref to config block Total FTE cell
            style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                       fmt="0.00").value = f"=C{ext_config_row}"

            # Variance: Budget - Sum
            style_cell(ws, r, val_col_diff, font=BLACK_FORMULA,
                       fmt="0.00").value = (
                f"={col_letter(val_col_budget)}{r}"
                f"-{col_letter(val_col_sum)}{r}"
            )

            external_rows[entry["project"]] = r
            r += 1

        data_row_end = r - 1

        # TOTAL row
        total_row = r
        style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
        style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
        for mi in range(project_params["months"]):
            c = month_col_start + mi
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
        for mi in range(project_params["months"]):
            c = month_col_start + mi
            cl = col_letter(c)
            style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                       fmt="0.00").value = (
                f"={SHEET_CONFIG}!B{roster_row}-{cl}{total_row}"
            )
        r += 1

        # Conditional formatting for TOTAL row
        total_range = (f"{col_letter(month_col_start)}{total_row}:"
                       f"{col_letter(month_col_end)}{total_row}")
        # Amber if > CAPACITY_AMBER_LOAD x Max FTE (added first, lower priority)
        ws.conditional_formatting.add(total_range, FormulaRule(
            formula=[f"{col_letter(month_col_start)}{total_row}>{CAPACITY_AMBER_LOAD}*{SHEET_CONFIG}!$B${roster_row}"],
            fill=AMBER_FILL,
        ))
        # Red if > Max FTE (added second, overrides amber for overcommitted cells)
        ws.conditional_formatting.add(total_range, FormulaRule(
            formula=[f"{col_letter(month_col_start)}{total_row}>{SHEET_CONFIG}!$B${roster_row}"],
            fill=RED_FILL,
        ))

        # Conditional formatting for REMAINING row — red if < 0
        remaining_range = (f"{col_letter(month_col_start)}{remaining_row}:"
                           f"{col_letter(month_col_end)}{remaining_row}")
        ws.conditional_formatting.add(remaining_range, FormulaRule(
            formula=[f"{col_letter(month_col_start)}{remaining_row}<0"],
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
