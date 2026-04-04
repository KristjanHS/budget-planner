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
    SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    YELLOW_FILL, AMBER_FILL, RED_FILL, CENTER, style_cell, col_letter,
    write_section_header,
)


def _wp_config_row(wp_id, work_packages, config_first_wp_row):
    """Return the Config sheet WP table row for a WP ID."""
    for i, wp in enumerate(work_packages):
        if wp["id"] == wp_id:
            return config_first_wp_row + i
    raise ValueError(f"Unknown WP ID: {wp_id}")


# ── Distribution curves ──

def distribute(total_fte: float, num_months: int, curve: str) -> list[float]:
    """Distribute total_fte across num_months using the given curve.

    All distributions sum exactly to total_fte.

    Args:
        total_fte: Total person-months to distribute.
        num_months: Number of active months.
        curve: One of "flat", "front-loaded", "bell", "back-loaded".

    Returns:
        List of float values, one per month, summing to total_fte.
    """
    if num_months <= 0:
        return []
    if num_months == 1:
        return [total_fte]

    if curve == "flat":
        per_month = total_fte / num_months
        return [per_month] * num_months

    if curve == "front-loaded":
        # Linearly decreasing: weight[i] = num_months - i
        weights = [num_months - i for i in range(num_months)]
        total_w = sum(weights)
        return [total_fte * w / total_w for w in weights]

    if curve == "back-loaded":
        # Linearly increasing: weight[i] = i + 1
        weights = [i + 1 for i in range(num_months)]
        total_w = sum(weights)
        return [total_fte * w / total_w for w in weights]

    if curve == "bell":
        # Triangular: ramp up to midpoint, symmetric ramp down
        mid = (num_months - 1) / 2.0
        weights = [mid - abs(i - mid) + 1 for i in range(num_months)]
        total_w = sum(weights)
        return [total_fte * w / total_w for w in weights]

    # Unknown curve — fall back to flat
    per_month = total_fte / num_months
    return [per_month] * num_months


def build_capacity_sheet(ws, wb, team_layout, *,
                         project_params=None, work_packages=None,
                         team=None, external_commitments=None,
                         assignments=None,
                         preserved=None, config_first_wp_row=None):
    """Build the Capacity sheet: config blocks + per-person groups.

    When ``assignments`` is provided, per-person grids show only assigned
    WPs with auto-distributed monthly values. When empty or None, persons
    have no programme WP rows (empty template for manual assignment).

    Args:
        ws: The Capacity worksheet to populate.
        wb: The parent Workbook.
        team_layout: dict from build_config_sheet with person_rows, etc.
        project_params: dict with keys title, start, months, rate.
        work_packages: list of WP dicts (now includes 'type').
        team: list of team member dicts.
        external_commitments: list of external commitment dicts.
        assignments: list of person-WP assignment dicts with keys
            person, wp, total_fte, start, end. When None, no WP rows.
        preserved: dict keyed by (person_name, wp_id) -> list of monthly
            float values for round-trip rebuild mode. NOT USED in new design
            (monthly values are always recalculated from assignments).
        config_first_wp_row: First WP row in Config WP table. Required.

    Returns:
        capacity_layout dict with row/column positions for all key elements.
    """
    # ── Lazy-import fallback for backward compatibility ──
    if any(v is None for v in [project_params, work_packages, team, external_commitments]):
        from budget_planner import config
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES
        if team is None:
            team = config.TEAM
        if external_commitments is None:
            external_commitments = config.EXTERNAL_COMMITMENTS

    if assignments is None:
        assignments = []

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

    # Build WP type lookup
    wp_type_map = {wp["id"]: wp.get("type", "flat") for wp in work_packages}

    ws.sheet_properties.tabColor = "7030A0"
    r = 1

    person_rows = team_layout["person_rows"]

    # ── CONFIG BLOCK 1: PERSON-WP ASSIGNMENTS ──
    write_section_header(ws, r, "PERSON-WP ASSIGNMENTS", merge_end=5)
    r += 1

    assign_col_headers = ["Person", "WP", "Total FTE", "Start Month", "End Month"]
    for ci, h in enumerate(assign_col_headers, 1):
        style_cell(ws, r, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    assign_first_row = r
    assign_config_rows: dict[tuple[str, str], int] = {}
    for entry in assignments:
        style_cell(ws, r, 1, entry["person"], font=BLUE_INPUT)
        style_cell(ws, r, 2, entry["wp"], font=BLUE_INPUT)
        style_cell(ws, r, 3, entry["total_fte"], font=BLUE_INPUT, fmt="0.00")
        style_cell(ws, r, 4, entry["start"], font=BLUE_INPUT, fmt="0")
        style_cell(ws, r, 5, entry["end"], font=BLUE_INPUT, fmt="0")
        assign_config_rows[(entry["person"], entry["wp"])] = r
        r += 1
    assign_last_row = r - 1 if assignments else r

    # Collapse assignments block
    if assignments:
        ws.row_dimensions.group(assign_first_row - 1, assign_last_row, hidden=True)

    # Gap row
    r += 1

    # ── CONFIG BLOCK 2: EXTERNAL COMMITMENTS ──
    write_section_header(ws, r, "EXTERNAL COMMITMENTS", merge_end=5)
    r += 1

    ext_col_headers = ["Person", "Project", "Total FTE", "Start Month", "End Month"]
    for ci, h in enumerate(ext_col_headers, 1):
        style_cell(ws, r, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    ext_first_row = r
    external_config_rows: dict[tuple[str, str], int] = {}
    for entry in external_commitments:
        style_cell(ws, r, 1, entry["person"], font=BLUE_INPUT)
        style_cell(ws, r, 2, entry["project"], font=BLUE_INPUT)
        style_cell(ws, r, 3, entry["fte"], font=BLUE_INPUT, fmt="0.00")
        style_cell(ws, r, 4, entry["start"], font=BLUE_INPUT, fmt="0")
        style_cell(ws, r, 5, entry["end"], font=BLUE_INPUT, fmt="0")
        external_config_rows[(entry["person"], entry["project"])] = r
        r += 1
    ext_last_row = r - 1 if external_commitments else r

    # Collapse external commitments block
    if external_commitments:
        ws.row_dimensions.group(ext_first_row - 1, ext_last_row, hidden=True)

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

    # ── Build per-person assignment lookup ──
    person_assignments: dict[str, list[dict]] = {}
    for entry in assignments:
        person_assignments.setdefault(entry["person"], []).append(entry)

    # ── PER-PERSON GROUPS ──
    person_groups: dict[str, dict] = {}

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
        wp_rows: dict[str, int] = {}
        external_rows: dict[str, int] = {}

        # Programme WP rows — only assigned WPs for this person
        my_assignments = person_assignments.get(name, [])
        first_wp_row_for_person = True
        for asgn in my_assignments:
            wp_id = asgn["wp"]
            wp = next((w for w in work_packages if w["id"] == wp_id), None)
            if wp is None:
                continue
            config_row = _wp_config_row(wp_id, work_packages, config_first_wp_row)

            # Col A: formula ref to Config WP ID
            style_cell(ws, r, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
            # Col B: formula ref to Config WP Name
            style_cell(ws, r, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

            # Calculate distribution
            asgn_start = asgn["start"]
            asgn_end = asgn["end"]
            num_active = asgn_end - asgn_start + 1
            wp_curve = wp_type_map.get(wp_id, "flat")
            monthly_values = distribute(asgn["total_fte"], num_active, wp_curve)

            # Write monthly cells
            comment_added = False
            active_idx = 0
            for mi in range(project_params["months"]):
                c = month_col_start + mi
                month_num = mi + 1

                if asgn_start <= month_num <= asgn_end:
                    val = monthly_values[active_idx] if active_idx < len(monthly_values) else 0
                    active_idx += 1
                    style_cell(ws, r, c, round(val, 4), font=BLUE_INPUT, fmt="0.00")
                    if first_wp_row_for_person and not comment_added and val == 0:
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

            # Budget: ref to assignment config block Total FTE
            asgn_row = assign_config_rows.get((name, wp_id))
            if asgn_row is not None:
                style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                           fmt="0.00").value = f"=C{asgn_row}"
            else:
                style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                           fmt="0.00").value = f"={SHEET_CONFIG}!C{config_row}"

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

            # Distribute external commitment using flat curve
            ext_start = entry["start"]
            ext_end = entry["end"]
            ext_months = ext_end - ext_start + 1
            ext_values = distribute(entry["fte"], ext_months, "flat")

            ext_idx = 0
            for mi in range(project_params["months"]):
                c = month_col_start + mi
                month_num = mi + 1

                if ext_start <= month_num <= ext_end:
                    val = ext_values[ext_idx] if ext_idx < len(ext_values) else 0
                    ext_idx += 1
                    style_cell(ws, r, c, round(val, 4), font=BLUE_INPUT, fmt="0.00")
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

        # Handle case where person has no data rows
        if data_row_end < data_row_start:
            data_row_end = data_row_start
            # Write an empty placeholder row so TOTAL SUM has a range
            r += 1

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
        "assign_first_row": assign_first_row,
        "assign_last_row": assign_last_row,
        "config_first_row": ext_first_row,
        "config_last_row": ext_last_row,
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "person_groups": person_groups,
        "external_config_rows": external_config_rows,
        "assign_config_rows": assign_config_rows,
    }
