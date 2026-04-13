"""Actuals-Hours sheet builder — per-person monthly hours timelog.

Mirrors the Capacity sheet geometry 1:1 (same WPs, same externals, same
months, same header rows) but holds **hours** rather than FTE. Hour input
cells are blank blue inputs the user fills in at month-end; inactive WP
months are greyed and locked, matching the Capacity behavior the auditors
expect.

Scope of this module (T3 of the hours-timelog plan): hour inputs, per-row
validation Sum, per-person TOTAL and EU TOTAL hour rows, and the EU ceiling
conditional formatting per calendar year. No forecast-EUR columns, no
REMAINING row, no rebuild-preservation plumbing yet — those land in later
tasks.
"""

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG,
    build_year_month_headers,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
    _calendar_year_groups,
)
from budget_planner.person_section import plan_person_layout
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CF_AMBER_FILL, CF_RED_FILL, style_cell, col_letter,
)

HOUR_FMT = "#,##0"


def _wp_config_row(wp_id, work_packages, config_first_wp_row):
    """Return the Config sheet WP table row for a WP ID."""
    for i, wp in enumerate(work_packages):
        if wp["id"] == wp_id:
            return config_first_wp_row + i
    raise ValueError(f"Unknown WP ID: {wp_id}")


def _write_person_hours_grid(
    ws, r: int, member: dict, *,
    person_assignments: list[dict],
    work_packages: list[dict],
    external_commitments: list[dict],
    wp_active_map: dict[str, frozenset[int]],
    person_rows: dict[str, int],
    project_params: dict,
    month_col_start: int,
    month_col_end: int,
    val_col_sum: int,
    config_first_wp_row: int,
    preserved_for_person: dict[str, list[float | None]] | None,
) -> tuple[int, dict]:
    """Write one person's hour-grid section (header, WP rows, externals, totals).

    Returns:
        (next_row, person_group_dict)
    """
    name = member["name"]
    roster_row = person_rows[name]

    # Pre-compute ordered WP ids and externals so plan_person_layout can
    # assign row numbers up front — mirror the Capacity behavior.
    ordered_wp_ids: list[str] = []
    for asgn in person_assignments:
        wp_id = asgn["wp"]
        if any(w["id"] == wp_id for w in work_packages):
            if wp_id not in ordered_wp_ids:
                ordered_wp_ids.append(wp_id)
    person_externals = [e for e in external_commitments if e["person"] == name]
    ordered_external_names = [e["project"] for e in person_externals]

    layout_start = r
    if not ordered_wp_ids and not ordered_external_names:
        layout_start = r + 1
    layout = plan_person_layout(
        start_row=layout_start,
        wp_ids=ordered_wp_ids,
        external_names=ordered_external_names,
        include_remaining=False,
        include_eu_total=True,
    )

    # ── Person header row ──
    header_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, font=HEADER_FONT, fill=SECTION_FILL)
    ws.cell(row=r, column=1).value = f"={SHEET_CONFIG}!A{roster_row}"
    ws.cell(row=r, column=1).font = Font(name="Arial", bold=True, size=10,
                                          color="FFFFFF")
    style_cell(ws, r, 2, fill=SECTION_FILL)
    r += 1

    data_row_start = r
    wp_rows: dict[str, int] = {}
    external_rows: dict[str, int] = {}

    # ── programme WP rows ──
    first_wp_row_for_person = True
    for wp_id in ordered_wp_ids:
        config_row = _wp_config_row(wp_id, work_packages, config_first_wp_row)
        assert r == layout.wp_rows[wp_id]

        style_cell(ws, r, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(ws, r, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        wp_active = wp_active_map.get(wp_id, frozenset())
        preserved_months = (
            preserved_for_person.get(wp_id) if preserved_for_person else None
        )

        comment_added = False
        for mi in range(project_params["months"]):
            c = month_col_start + mi
            month_num = mi + 1

            if month_num in wp_active:
                # Start blank — user fills at month-end. Preserved values
                # (from rebuild, Task 13) overwrite blanks when provided.
                val: float | str = ""
                if preserved_months is not None and mi < len(preserved_months):
                    p = preserved_months[mi]
                    if p is not None:
                        val = p
                style_cell(ws, r, c, val, font=BLUE_INPUT, fmt=HOUR_FMT)
                if first_wp_row_for_person and not comment_added:
                    ws.cell(row=r, column=c).comment = Comment(
                        "Enter hours worked per month for this work package. "
                        "Leave blank until the month closes.",
                        "Programme Budget Tool",
                    )
                    comment_added = True
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt=HOUR_FMT)

        # Validation: Sum column only (no Budget/Variance — hours log is a
        # record, not a budget-vs-actual view; design §1).
        sum_range = (f"{col_letter(month_col_start)}{r}:"
                     f"{col_letter(month_col_end)}{r}")
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fmt=HOUR_FMT).value = f"=SUM({sum_range})"

        wp_rows[wp_id] = r
        first_wp_row_for_person = False
        r += 1

    # ── External project rows ──
    for entry in person_externals:
        assert r == layout.external_rows[entry["project"]]

        style_cell(ws, r, 1, entry["project"], font=BLUE_INPUT)
        style_cell(ws, r, 2, font=BLUE_INPUT)

        ext_start = entry["start"]
        ext_end = entry["end"]
        preserved_months = (
            preserved_for_person.get(entry["project"])
            if preserved_for_person else None
        )

        for mi in range(project_params["months"]):
            c = month_col_start + mi
            month_num = mi + 1

            if ext_start <= month_num <= ext_end:
                val: float | str = ""
                if preserved_months is not None and mi < len(preserved_months):
                    p = preserved_months[mi]
                    if p is not None:
                        val = p
                style_cell(ws, r, c, val, font=BLUE_INPUT, fmt=HOUR_FMT)
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt=HOUR_FMT)

        sum_range = (f"{col_letter(month_col_start)}{r}:"
                     f"{col_letter(month_col_end)}{r}")
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fmt=HOUR_FMT).value = f"=SUM({sum_range})"

        external_rows[entry["project"]] = r
        r += 1

    data_row_end = r - 1

    # Person with no data rows: keep placeholder row so TOTAL sits where
    # plan_person_layout said it would.
    if data_row_end < data_row_start:
        data_row_end = data_row_start
        r += 1

    # ── TOTAL row (SUM of all data rows, hours per month) ──
    assert r == layout.total_row
    total_row = layout.total_row
    style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    ws.cell(row=r, column=2).comment = Comment(
        "Per-month SUM of all hours logged across programme WP rows and external "
        "project rows for this person.",
        "Programme Budget Tool",
    )
    for mi in range(project_params["months"]):
        c = month_col_start + mi
        cl = col_letter(c)
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt=HOUR_FMT).value = (
            f"=SUM({cl}{data_row_start}:{cl}{data_row_end})"
        )
    scl = col_letter(val_col_sum)
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt=HOUR_FMT).value = (
        f"=SUM({scl}{data_row_start}:{scl}{data_row_end})"
    )
    r += 1

    # ── EU TOTAL row (SUM of programme WP rows + EU-flagged externals) ──
    assert layout.eu_total_row is not None
    assert r == layout.eu_total_row
    eu_total_row = layout.eu_total_row
    style_cell(ws, r, 1, fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "EU TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)
    ws.cell(row=r, column=2).comment = Comment(
        "Per-month SUM of programme WP rows + EU-flagged external project rows "
        "(hours). Compared per calendar year against "
        "eu_annual_ceiling_pm × hours_per_fte_month.",
        "Programme Budget Tool",
    )

    eu_row_refs = list(wp_rows.values())
    for proj, ext_row in external_rows.items():
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
                   fmt=HOUR_FMT).value = formula

    # Sum column for EU TOTAL
    if eu_row_refs:
        sum_refs = ",".join(f"{col_letter(val_col_sum)}{row}"
                            for row in eu_row_refs)
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fill=LIGHT_BLUE_FILL,
                   fmt=HOUR_FMT).value = f"=SUM({sum_refs})"
    else:
        style_cell(ws, r, val_col_sum, font=BLACK_FORMULA,
                   fill=LIGHT_BLUE_FILL, fmt=HOUR_FMT).value = "=0"
    r += 1

    # ── EU ceiling CF per calendar year on EU TOTAL row ──
    # Red: annual_hours > eu_annual_ceiling_pm * hours_per_fte_month
    # Amber: within one hours_per_fte_month unit of the ceiling
    for group_start, span in _calendar_year_groups(
        project_params["start"], project_params["months"]
    ):
        yr_first_col = month_col_start + group_start
        yr_last_col = yr_first_col + span - 1
        yr_range = (f"{col_letter(yr_first_col)}{eu_total_row}:"
                    f"{col_letter(yr_last_col)}{eu_total_row}")
        sum_expr = (f"SUM(${col_letter(yr_first_col)}${eu_total_row}"
                    f":${col_letter(yr_last_col)}${eu_total_row})")
        # Red first — highest priority in openpyxl CF ordering.
        ws.conditional_formatting.add(yr_range, FormulaRule(
            formula=[f"{sum_expr}>eu_annual_ceiling_pm*hours_per_fte_month"],
            fill=CF_RED_FILL,
        ))
        ws.conditional_formatting.add(yr_range, FormulaRule(
            formula=[
                f"{sum_expr}>eu_annual_ceiling_pm*hours_per_fte_month"
                f"-hours_per_fte_month"
            ],
            fill=CF_AMBER_FILL,
        ))

    assert r == layout.next_row
    return layout.next_row, {
        "header_row": header_row,
        "wp_rows": wp_rows,
        "external_rows": external_rows,
        "total_row": total_row,
        "eu_total_row": eu_total_row,
    }


def build_actuals_hours_sheet(
    ws,
    team_layout,
    *,
    project_params,
    work_packages,
    wp_schedule,
    team,
    external_commitments,
    assignments,
    config_first_wp_row,
    config_layout,
    preserved: dict[str, dict[str, list[float | None]]] | None = None,
) -> dict:
    """Build the Actuals-Hours sheet: per-person monthly hour timelog grid.

    Args:
        ws: The Actuals-Hours worksheet to populate.
        team_layout: Config layout dict (needs person_rows).
        project_params: dict with 'title', 'start', 'months', 'rate'.
        work_packages: list of WP dicts.
        wp_schedule: list of WP schedule entries (for active-month greying).
        team: list of team member dicts.
        external_commitments: list of external commitment dicts.
        assignments: list of person-WP assignment dicts (drives which WP
            rows appear per person — same intersection as Capacity).
        config_first_wp_row: First WP row in Config WP table.
        config_layout: Config layout dict (reserved for future cross-refs).
        preserved: Optional rebuild-preservation payload
            ``{person_name: {wp_id_or_project: [hours_per_month, ...]}}``.
            When provided, blue-input hour cells are pre-filled from this
            map (None entries are skipped). Task 13 wires the reader.

    Returns:
        actuals_hours_layout dict: ``{year_header_row, month_header_row,
        person_groups: {name: {header_row, wp_rows, external_rows,
        total_row, eu_total_row}}}``.
    """
    if assignments is None:
        assignments = []

    month_col_start = MONTH_COL_START
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    val_col_sum = _vcols["sum"]

    ws.sheet_properties.tabColor = "C00000"  # distinct red — hours timelog

    person_rows = team_layout["person_rows"]

    # ── Header rows (Sum only — no Budget/Variance) ──
    r = 1
    year_header_row, month_header_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[(val_col_sum, "Sum")],
    )

    ws.cell(row=year_header_row, column=val_col_sum).comment = Comment(
        "Row-wide SUM of hours logged this row. Per-person TOTAL and EU "
        "TOTAL rows sum across these.",
        "Programme Budget Tool",
    )

    first_person_header_row = r

    # ── Per-person assignment lookup (dedup by person+wp) ──
    _dedup: dict[tuple[str, str], dict] = {}
    for entry in assignments:
        _dedup[(entry["person"], entry["wp"])] = entry
    per_person_assignments: dict[str, list[dict]] = {}
    for entry in _dedup.values():
        per_person_assignments.setdefault(entry["person"], []).append(entry)

    wp_active_map: dict[str, frozenset[int]] = {
        wp["id"]: wp_active_months(wp["id"], wp_schedule)
        for wp in work_packages
    }

    person_groups: dict[str, dict] = {}

    for member in team:
        name = member["name"]
        preserved_for_person = preserved.get(name) if preserved else None
        r, group = _write_person_hours_grid(
            ws, r, member,
            person_assignments=per_person_assignments.get(name, []),
            work_packages=work_packages,
            external_commitments=external_commitments,
            wp_active_map=wp_active_map,
            person_rows=person_rows,
            project_params=project_params,
            month_col_start=month_col_start,
            month_col_end=month_col_end,
            val_col_sum=val_col_sum,
            config_first_wp_row=config_first_wp_row,
            preserved_for_person=preserved_for_person,
        )
        person_groups[name] = group

    ws.freeze_panes = f"C{first_person_header_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "person_groups": person_groups,
    }
