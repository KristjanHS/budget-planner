"""Capacity sheet builder — per-person monthly FTE allocation."""

from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG, CAPACITY_AMBER_LOAD,
    build_year_month_headers, month_col_end as _month_col_end, val_columns,
    _calendar_year_groups,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CF_AMBER_FILL, CF_RED_FILL, style_cell, col_letter,
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


def _write_person_grid(
    ws, r: int, member: dict, *,
    person_assignments: list[dict],
    work_packages: list[dict],
    external_commitments: list[dict],
    wp_type_map: dict[str, str],
    assign_config_rows: dict[tuple[str, str], int],
    external_config_rows: dict[tuple[str, str], int],
    person_rows: dict[str, int],
    project_params: dict,
    month_col_start: int,
    month_col_end: int,
    val_col_sum: int,
    val_col_budget: int,
    val_col_diff: int,
    config_first_wp_row: int,
    team_layout: dict,
) -> tuple[int, dict]:
    """Write one person's grid section (header, WP rows, externals, totals).

    Returns:
        (next_row, person_group_dict) where person_group_dict has keys
        header_row, wp_rows, external_rows, total_row, remaining_row,
        eu_total_row.
    """
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
    first_wp_row_for_person = True
    for asgn in person_assignments:
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
                if first_wp_row_for_person and not comment_added:
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

        # Budget: ref to Config sheet assignment Total FTE
        asgn_row = assign_config_rows.get((name, wp_id))
        if asgn_row is not None:
            style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                       fmt="0.00").value = f"={SHEET_CONFIG}!C{asgn_row}"
        else:
            style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                       fmt="0.00").value = f"={SHEET_CONFIG}!C{config_row}"

        # Variance: Budget - Sum
        style_cell(ws, r, val_col_diff, font=BLACK_FORMULA,
                   fmt="0.00").value = (
            f"={col_letter(val_col_budget)}{r}"
            f"-{col_letter(val_col_sum)}{r}"
        )

        # CF: red if variance < 0 (overallocated)
        diff_cl = col_letter(val_col_diff)
        ws.conditional_formatting.add(
            f"{diff_cl}{r}:{diff_cl}{r}",
            FormulaRule(formula=[f"{diff_cl}{r}<0"], fill=CF_RED_FILL),
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

        # Budget: ref to Config sheet external commitment Total FTE cell
        style_cell(ws, r, val_col_budget, font=GREEN_XREF,
                   fmt="0.00").value = f"={SHEET_CONFIG}!C{ext_config_row}"

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
    # Sum/Budget/Variance for TOTAL row
    scl, bcl = col_letter(val_col_sum), col_letter(val_col_budget)
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({scl}{data_row_start}:{scl}{data_row_end})"
    style_cell(ws, r, val_col_budget, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({bcl}{data_row_start}:{bcl}{data_row_end})"
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"={bcl}{r}-{scl}{r}"
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
    # Sum/Budget/Variance for REMAINING row (all empty — per-month only)
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00")
    style_cell(ws, r, val_col_budget, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00")
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00")
    r += 1

    # EU TOTAL row — sum of programme WP rows + EU-flagged external rows
    eu_total_row = r
    style_cell(ws, r, 1, fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "EU TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)

    # Build list of rows that count toward EU total
    eu_row_refs = list(wp_rows.values())  # All programme WPs are EU
    for proj, ext_row in external_rows.items():
        # Check if this external commitment is EU
        entry = next((e for e in person_externals
                      if e["project"] == proj), None)
        if entry and entry.get("is_eu", False):
            eu_row_refs.append(ext_row)

    for mi in range(project_params["months"]):
        c = month_col_start + mi
        cl = col_letter(c)
        if eu_row_refs:
            cell_refs = ",".join(f"{cl}{row}" for row in eu_row_refs)
            formula = f"=SUM({cell_refs})"
        else:
            formula = "=0"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL,
                   fmt="0.00").value = formula

    # Sum validation column for EU TOTAL
    if eu_row_refs:
        sum_refs = ",".join(f"{col_letter(val_col_sum)}{row}"
                            for row in eu_row_refs)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fill=LIGHT_BLUE_FILL, fmt="0.00").value = f"=SUM({sum_refs})"
    else:
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fill=LIGHT_BLUE_FILL, fmt="0.00").value = "=0"
    # Budget and Variance columns left empty for EU TOTAL
    style_cell(ws, r, val_col_budget, fill=LIGHT_BLUE_FILL, fmt="0.00")
    style_cell(ws, r, val_col_diff, fill=LIGHT_BLUE_FILL, fmt="0.00")
    r += 1

    # Conditional formatting for TOTAL row
    # Excel CF priority: lower number = higher priority = wins conflicts.
    # openpyxl assigns priority sequentially, so add red first (wins when both match).
    total_range = (f"{col_letter(month_col_start)}{total_row}:"
                   f"{col_letter(month_col_end)}{total_row}")
    # Red if > Max FTE (added first = highest priority, wins over amber)
    ws.conditional_formatting.add(total_range, FormulaRule(
        formula=[f"{col_letter(month_col_start)}{total_row}>{SHEET_CONFIG}!$B${roster_row}"],
        fill=CF_RED_FILL,
    ))
    # Amber if > CAPACITY_AMBER_LOAD x Max FTE (lower priority, only visible when red doesn't match)
    ws.conditional_formatting.add(total_range, FormulaRule(
        formula=[f"{col_letter(month_col_start)}{total_row}>{CAPACITY_AMBER_LOAD}*{SHEET_CONFIG}!$B${roster_row}"],
        fill=CF_AMBER_FILL,
    ))

    # Conditional formatting for REMAINING row
    remaining_range = (f"{col_letter(month_col_start)}{remaining_row}:"
                       f"{col_letter(month_col_end)}{remaining_row}")
    remaining_amber_threshold = 1 - CAPACITY_AMBER_LOAD
    # Red if < 0 (added first = highest priority, wins over amber)
    ws.conditional_formatting.add(remaining_range, FormulaRule(
        formula=[f"{col_letter(month_col_start)}{remaining_row}<0"],
        fill=CF_RED_FILL,
    ))
    # Amber if remaining < 20% of Max FTE (lower priority, only visible when red doesn't match)
    ws.conditional_formatting.add(remaining_range, FormulaRule(
        formula=[f"{col_letter(month_col_start)}{remaining_row}<{remaining_amber_threshold}*{SHEET_CONFIG}!$B${roster_row}"],
        fill=CF_AMBER_FILL,
    ))

    # Conditional formatting for EU TOTAL — per calendar year
    ceiling_row = team_layout["param_rows"]["EU Annual Ceiling (PM/FTE/year)"]
    for group_start, span in _calendar_year_groups(
        project_params["start"], project_params["months"]
    ):
        yr_first_col = month_col_start + group_start
        yr_last_col = yr_first_col + span - 1
        yr_range = (f"{col_letter(yr_first_col)}{eu_total_row}:"
                    f"{col_letter(yr_last_col)}{eu_total_row}")
        # Absolute refs so formula doesn't shift across the range
        sum_expr = (f"SUM(${col_letter(yr_first_col)}${eu_total_row}"
                    f":${col_letter(yr_last_col)}${eu_total_row})")
        # Red if annual EU total > ceiling (added first = higher priority)
        ws.conditional_formatting.add(yr_range, FormulaRule(
            formula=[f"{sum_expr}>{SHEET_CONFIG}!$B${ceiling_row}"],
            fill=CF_RED_FILL,
        ))
        # Amber if annual EU total > ceiling - 1
        ws.conditional_formatting.add(yr_range, FormulaRule(
            formula=[f"{sum_expr}>{SHEET_CONFIG}!$B${ceiling_row}-1"],
            fill=CF_AMBER_FILL,
        ))

    return r, {
        "header_row": header_row,
        "wp_rows": wp_rows,
        "external_rows": external_rows,
        "total_row": total_row,
        "remaining_row": remaining_row,
        "eu_total_row": eu_total_row,
    }


def build_capacity_sheet(ws, team_layout, *,
                         project_params=None, work_packages=None,
                         wp_schedule=None,
                         team=None, external_commitments=None,
                         assignments=None,
                         config_first_wp_row=None):
    """Build the Capacity sheet: year/month headers + per-person groups.

    When ``assignments`` is provided, per-person grids show only assigned
    WPs with auto-distributed monthly values. When empty or None, persons
    have no programme WP rows (empty template for manual assignment).

    Assignment and external-commitment tables live on the Config sheet;
    the Capacity sheet only holds the planning grid and references
    Config via =Config!C<row> for per-row FTE budgets.

    Args:
        ws: The Capacity worksheet to populate.
        team_layout: dict from build_config_sheet. Must include
            person_rows, assign_config_rows, external_config_rows.
        project_params: dict with keys title, start, months, rate.
        work_packages: list of WP dicts with keys 'id', 'name', 'fte', 'rate'.
        team: list of team member dicts.
        external_commitments: list of external commitment dicts.
        assignments: list of person-WP assignment dicts with keys
            person, wp, total_fte, start, end. When None, no WP rows.
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
    if wp_schedule is None:
        from budget_planner.config import WP_SCHEDULE
        wp_schedule = WP_SCHEDULE

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

    # Build WP type lookup from schedule (first matching entry per WP)
    wp_type_map: dict[str, str] = {}
    for entry in wp_schedule:
        if entry["id"] not in wp_type_map:
            wp_type_map[entry["id"]] = entry.get("type", "flat")

    ws.sheet_properties.tabColor = "7030A0"

    person_rows = team_layout["person_rows"]
    assign_config_rows = team_layout["assign_config_rows"]
    external_config_rows = team_layout["external_config_rows"]

    # ── YEAR + MONTH HEADER ROWS (once, before first person group) ──
    r = 1
    year_header_row, month_header_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            (val_col_sum, "Sum"),
            (val_col_budget, "Budget"),
            (val_col_diff, "Variance"),
        ],
    )

    first_person_header_row = r

    # ── Build per-person assignment lookup (deduplicate by person+wp) ──
    _dedup: dict[tuple[str, str], dict] = {}
    for entry in assignments:
        _dedup[(entry["person"], entry["wp"])] = entry
    per_person_assignments: dict[str, list[dict]] = {}
    for entry in _dedup.values():
        per_person_assignments.setdefault(entry["person"], []).append(entry)

    # ── PER-PERSON GROUPS ──
    person_groups: dict[str, dict] = {}

    for member in team:
        name = member["name"]
        r, group = _write_person_grid(
            ws, r, member,
            person_assignments=per_person_assignments.get(name, []),
            work_packages=work_packages,
            external_commitments=external_commitments,
            wp_type_map=wp_type_map,
            assign_config_rows=assign_config_rows,
            external_config_rows=external_config_rows,
            person_rows=person_rows,
            project_params=project_params,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            val_col_sum=val_col_sum,
            val_col_budget=val_col_budget,
            val_col_diff=val_col_diff,
            config_first_wp_row=config_first_wp_row,
            team_layout=team_layout,
        )
        person_groups[name] = group

    # Freeze panes
    ws.freeze_panes = f"C{first_person_header_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "person_groups": person_groups,
        "external_config_rows": external_config_rows,
        "assign_config_rows": assign_config_rows,
    }
