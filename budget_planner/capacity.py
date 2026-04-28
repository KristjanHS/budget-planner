"""Capacity sheet builder — per-person monthly FTE allocation.

Per-person groups with WP rows (blue input, distribution-curve static
values) and a TOTAL row (SUM across the WP rows for that person). REMAINING
and EU TOTAL rows were dropped 2026-04-28 — REMAINING < 0 folded onto the
TOTAL > Max FTE red CF, and the per-year EU ceiling now lives only on
Actuals-Hours (the auditor enforcement surface — see ``actuals_hours.py``).

Non-obvious conventions:

- Per-person WP rows cover **only the assigned WPs** (from Config
  PERSON-WP ASSIGNMENTS), not every WP.
- Distribution curve type comes from the Type col on the merged WP BUDGET
  row.
- The right-side raw-monthly band sits at ``outline_level=1`` (default
  collapsed).
- The always-visible TO-DATE ``Variance PM`` column is populated only on
  the TOTAL row.

Returns ``capacity_layout`` with ``person_groups`` + ``todate_col_actual`` /
``todate_col_var``.

Role: monthly cells are auto-filled by distribution curves on rebuild;
users can override for fine-tuning (overrides are lost on next rebuild).

## Conditional formatting

- TOTAL row per person: **amber** when ``TOTAL > 0.8 * Max FTE``,
  **red** when ``TOTAL > Max FTE`` (the red also signals what the prior
  REMAINING < 0 rule signalled).
- Raw-monthly Variance-FTE band: **red** when
  ``|variance| > CAPACITY_VARIANCE_FTE_RED``.
- Raw-monthly Actual-TOTAL utilisation vs Config Max FTE: amber/red same
  thresholds as planned TOTAL.
- To-date block ``Variance PM`` (TOTAL only): **red** when ``< 0``. Block
  reduced to 2 cols (2026-04-14 simplification) — the prior per-WP
  Variance-% and to-date util/ceil CF rules were dropped as duplicates of
  planned-side signals.
"""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.properties import Outline

from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG, SHEET_ACTUALS_HOURS,
    CAPACITY_AMBER_LOAD, CAPACITY_VARIANCE_FTE_RED, COL_ROSTER_FTE,
    build_year_month_headers, cell_ref, month_range_ref,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
    _calendar_year_groups,
)
from budget_planner.person_section import (
    plan_person_layout, PersonGridContext, apply_person_row_outline,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    SECTION_FILL, GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    CF_AMBER_FILL, CF_RED_FILL, CENTER, style_cell, col_letter,
    sized_comment,
)
from budget_planner import config

# Band labels for the raw monthly Actual/Variance bands (row 1 merged).
# Both bands default to collapsed (see the column_dimensions setup in
# build_capacity_sheet); the warning text flags that per-month values are
# only trustworthy when hours are entered month-by-month.
ACTUAL_BAND_LABEL = "ACTUAL FTE (raw monthly — unreliable unless hours entered per-month)"
VARIANCE_BAND_LABEL = "VARIANCE FTE (raw monthly — unreliable unless hours entered per-month)"


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


def _resolve_grid_state(
    *, start_row: int, member: dict, ctx: PersonGridContext,
    person_assignments: list[dict], person_rows: dict[str, int],
    team_layout: dict,
) -> dict:
    """Bundle loop-invariant state for one person's grid into a single dict.

    Skips assignments whose WP is not in ``work_packages`` (matches the
    legacy inline loop that ``continue``d past missing WPs without
    incrementing ``r``). When the person has no data rows, bumps the
    layout start by 1 so TOTAL lands one row lower (legacy placeholder).

    Returns dict with: name, roster_row, wp_by_id, wp_idx,
    person_externals, layout.
    """
    name = member["name"]
    roster_row = person_rows[name]

    # O(1) WP lookups — avoid O(N²) linear scans in the per-assignment loops
    # below (one scan for filtering, another for data-write, plus a third for
    # the Config-row offset).
    wp_by_id = {w["id"]: w for w in ctx.work_packages}
    wp_idx = {w["id"]: i for i, w in enumerate(ctx.work_packages)}

    person_externals = [e for e in ctx.external_commitments
                        if e["person"] == name]

    ordered_wp_ids: list[str] = []
    for asgn in person_assignments:
        wp_id = asgn["wp"]
        if wp_id in wp_by_id:
            ordered_wp_ids.append(wp_id)
    ordered_external_names = [e["project"] for e in person_externals]

    layout_start = start_row
    if not ordered_wp_ids and not ordered_external_names:
        layout_start = start_row + 1
    # Capacity drops REMAINING + EU TOTAL: REMAINING CF folded onto TOTAL,
    # EU-ceiling enforcement lives on Actuals-Hours (the per-year CF on the
    # EU TOTAL row over there is the auditor surface).
    layout = plan_person_layout(
        start_row=layout_start,
        wp_ids=ordered_wp_ids,
        external_names=ordered_external_names,
        include_remaining=False,
        include_eu_total=False,
    )
    return {
        "name": name,
        "roster_row": roster_row,
        "wp_by_id": wp_by_id,
        "wp_idx": wp_idx,
        "person_externals": person_externals,
        "layout": layout,
    }


def _write_header_row(ws, r: int, roster_row: int) -> None:
    """Write the merged person-header row (col A formula → Config!A<roster>)."""
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    style_cell(ws, r, 1, font=HEADER_FONT, fill=SECTION_FILL)
    ws.cell(row=r, column=1).value = f"={SHEET_CONFIG}!A{roster_row}"
    ws.cell(row=r, column=1).font = Font(name="Arial", bold=True, size=10,
                                          color="FFFFFF")
    style_cell(ws, r, 2, fill=SECTION_FILL)


def _write_wp_rows(
    ws, *, start_row: int, name: str, person_assignments: list[dict],
    wp_by_id: dict, wp_idx: dict[str, int], wp_active_map: dict,
    wp_type_map: dict[str, str],
    assign_config_rows: dict[tuple[str, str], int],
    config_first_wp_row: int, layout,
    project_params: dict, month_col_start: int, month_col_end: int,
    val_col_sum: int, val_col_budget: int, val_col_diff: int,
) -> tuple[int, dict[str, int]]:
    """Write programme WP rows (only assigned WPs); return (next_row, wp_rows)."""
    r = start_row
    wp_rows: dict[str, int] = {}
    first_wp_row_for_person = True
    for asgn in person_assignments:
        wp_id = asgn["wp"]
        wp = wp_by_id.get(wp_id)
        if wp is None:
            continue
        config_row = config_first_wp_row + wp_idx[wp_id]
        # Row position is pre-assigned by plan_person_layout; keep `r` in sync
        # with the inline cell-write sequence below.
        assert r == layout.wp_rows[wp_id]

        # Col A: formula ref to Config WP ID
        style_cell(ws, r, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
        # Col B: formula ref to Config WP Name
        style_cell(ws, r, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        # Distribute across the intersection of the assignment window with
        # the WP's active months — each WP has a single [start, end] range
        # on WP BUDGET, and FTE must not be allocated to months outside it
        # (auditors flag that).
        asgn_start = asgn["start"]
        asgn_end = asgn["end"]
        wp_active = wp_active_map.get(wp_id, frozenset())
        valid_months = frozenset(
            m for m in range(asgn_start, asgn_end + 1) if m in wp_active
        )
        wp_curve = wp_type_map.get(wp_id, "flat")
        monthly_values = distribute(asgn["total_fte"], len(valid_months), wp_curve)

        # Write monthly cells
        comment_added = False
        active_idx = 0
        for mi in range(project_params["months"]):
            c = month_col_start + mi
            month_num = mi + 1

            if month_num in valid_months:
                val = monthly_values[active_idx] if active_idx < len(monthly_values) else 0
                active_idx += 1
                style_cell(ws, r, c, round(val, 4), font=BLUE_INPUT, fmt="0.00")
                if first_wp_row_for_person and not comment_added:
                    ws.cell(row=r, column=c).comment = sized_comment("Enter monthly FTE allocation for each work "
                        "package and external project.")
                    comment_added = True
            else:
                style_cell(ws, r, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="0.00")

        # Validation columns
        sum_range = month_range_ref(r, month_col_start, month_col_end)
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
    return r, wp_rows


def _write_external_rows(
    ws, *, start_row: int, name: str, person_externals: list[dict],
    external_config_rows: dict[tuple[str, str], int], layout,
    project_params: dict, month_col_start: int, month_col_end: int,
    val_col_sum: int, val_col_budget: int, val_col_diff: int,
) -> tuple[int, dict[str, int]]:
    """Write external project rows; return (next_row, external_rows)."""
    r = start_row
    external_rows: dict[str, int] = {}
    for entry in person_externals:
        ext_config_row = external_config_rows[(name, entry["project"])]
        assert r == layout.external_rows[entry["project"]]

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
        sum_range = month_range_ref(r, month_col_start, month_col_end)
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
    return r, external_rows


def _write_total_row(
    ws, *, r: int, layout, data_row_start: int, data_row_end: int,
    project_params: dict, month_col_start: int,
    val_col_sum: int, val_col_budget: int, val_col_diff: int,
) -> int:
    """Write the per-person TOTAL row; return total_row."""
    assert r == layout.total_row
    total_row = layout.total_row
    style_cell(ws, r, 1, fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    ws.cell(row=r, column=2).comment = sized_comment("Per-month SUM of all programme WP rows above for this person. External "
        "project rows are NOT included here — see EU TOTAL for the "
        "EU-ceiling view that includes EU-flagged externals.")
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
    return total_row


def _apply_planned_total_overalloc_cf(
    ws, *, total_row: int, roster_row: int,
    month_col_start: int, month_col_end: int,
) -> None:
    """Apply over-allocation CF for the planned TOTAL row (red before amber).

    Folds the previous REMAINING-row CF onto TOTAL: red when monthly TOTAL
    exceeds Max FTE (i.e. REMAINING < 0), amber when TOTAL crosses the
    CAPACITY_AMBER_LOAD * Max FTE warning band.

    Excel CF priority: lower number = higher priority = wins conflicts.
    openpyxl assigns priority sequentially, so add red first (wins when
    both match).
    """
    total_range = month_range_ref(total_row, month_col_start, month_col_end)
    max_fte_ref = cell_ref(SHEET_CONFIG, roster_row, COL_ROSTER_FTE)
    # Red if > Max FTE (added first = highest priority, wins over amber)
    ws.conditional_formatting.add(total_range, FormulaRule(
        formula=[f"{col_letter(month_col_start)}{total_row}>{max_fte_ref}"],
        fill=CF_RED_FILL,
    ))
    # Amber if > CAPACITY_AMBER_LOAD x Max FTE (lower priority).
    ws.conditional_formatting.add(total_range, FormulaRule(
        formula=[f"{col_letter(month_col_start)}{total_row}>{CAPACITY_AMBER_LOAD}*{max_fte_ref}"],
        fill=CF_AMBER_FILL,
    ))


def _setup_actual_band_header(
    ws, *, header_row: int, actual_col_start: int, variance_col_start: int,
    months: int,
) -> None:
    """Style the actual + variance band cells on the person-header row."""
    for mi in range(months):
        style_cell(ws, header_row, actual_col_start + mi, fill=SECTION_FILL)
        style_cell(ws, header_row, variance_col_start + mi, fill=SECTION_FILL)


def _write_actual_pair(
    ws, row: int, *, is_active_month, months: int,
    month_col_start: int, actual_col_start: int, variance_col_start: int,
) -> None:
    """Write Actual FTE + Variance cells for a single row.

    is_active_month(month_num_1based) -> bool determines greying.
    """
    for mi in range(months):
        month_num = mi + 1
        planned_c = month_col_start + mi
        actual_c = actual_col_start + mi
        variance_c = variance_col_start + mi
        if is_active_month(month_num):
            hours_cell = (
                f"'{SHEET_ACTUALS_HOURS}'!"
                f"{col_letter(planned_c)}{row}"
            )
            # Actual FTE — green cross-sheet font (pulls from Actuals-Hours).
            style_cell(
                ws, row, actual_c, font=GREEN_XREF, fmt="0.00",
            ).value = (
                f'=IF({hours_cell}="","",'
                f"{hours_cell}/hours_per_fte_month)"
            )
            # Variance FTE — Planned - Actual, black formula font.
            planned_ref = f"{col_letter(planned_c)}{row}"
            actual_ref = f"{col_letter(actual_c)}{row}"
            style_cell(
                ws, row, variance_c, font=BLACK_FORMULA, fmt="0.00",
            ).value = (
                f'=IF({actual_ref}="","",{planned_ref}-{actual_ref})'
            )
        else:
            style_cell(ws, row, actual_c, "", font=GREEN_XREF,
                       fill=GREY_FILL, fmt="0.00")
            style_cell(ws, row, variance_c, "", font=BLACK_FORMULA,
                       fill=GREY_FILL, fmt="0.00")


def _write_actual_pairs_for_data_rows(
    ws, *, wp_rows: dict[str, int], external_rows: dict[str, int],
    person_assignments: list[dict], person_externals: list[dict],
    wp_active_map: dict, project_params: dict,
    months: int, month_col_start: int, actual_col_start: int,
    variance_col_start: int,
) -> None:
    """Fill Actual+Variance band cells for WP rows + external rows.

    Active-month predicates mirror planned-side greying:
      - WP rows: WP active schedule ∩ assignment window.
      - External rows: external commitment start/end window.
    """
    wp_windows: dict[str, tuple[int, int, frozenset[int]]] = {}
    for asgn in person_assignments:
        wp_id = asgn["wp"]
        if wp_id not in wp_rows:
            continue
        wp_active = wp_active_map.get(wp_id, frozenset())
        wp_windows[wp_id] = (asgn["start"], asgn["end"], wp_active)

    for wp_id, wp_row in wp_rows.items():
        asgn_start, asgn_end, wp_active = wp_windows.get(
            wp_id, (1, project_params["months"], frozenset())
        )

        def _wp_active(mn, _s=asgn_start, _e=asgn_end, _w=wp_active):
            return _s <= mn <= _e and mn in _w

        _write_actual_pair(
            ws, wp_row, is_active_month=_wp_active, months=months,
            month_col_start=month_col_start,
            actual_col_start=actual_col_start,
            variance_col_start=variance_col_start,
        )

    # External rows — active months = external commitment start/end window.
    for proj, ext_row in external_rows.items():
        entry = next((e for e in person_externals
                      if e["project"] == proj), None)
        if entry is None:
            continue
        ext_start, ext_end = entry["start"], entry["end"]

        def _ext_active(mn, _s=ext_start, _e=ext_end):
            return _s <= mn <= _e

        _write_actual_pair(
            ws, ext_row, is_active_month=_ext_active, months=months,
            month_col_start=month_col_start,
            actual_col_start=actual_col_start,
            variance_col_start=variance_col_start,
        )


def _write_actual_total_band(
    ws, *, total_row: int, wp_rows: dict[str, int],
    external_rows: dict[str, int], months: int,
    actual_col_start: int, variance_col_start: int,
) -> None:
    """Fill Actual TOTAL row band cells (SUM across wp_rows ∪ external_rows)."""
    data_rows_for_total = list(wp_rows.values()) + list(external_rows.values())
    for mi in range(months):
        ac = actual_col_start + mi
        vc = variance_col_start + mi
        acl = col_letter(ac)
        vcl = col_letter(vc)
        if data_rows_for_total:
            # IFERROR guards the case where all Actual cells are empty strings
            # (Actuals-Hours has no hours yet) → SUM of empty strings errors;
            # we want a clean 0 in that case.
            sum_terms = ",".join(
                f'IF({acl}{dr}="",0,{acl}{dr})'
                for dr in data_rows_for_total
            )
            actual_total_formula = f"=SUM({sum_terms})"
            var_sum_terms = ",".join(
                f'IF({vcl}{dr}="",0,{vcl}{dr})'
                for dr in data_rows_for_total
            )
            var_total_formula = f"=SUM({var_sum_terms})"
        else:
            actual_total_formula = "=0"
            var_total_formula = "=0"
        style_cell(ws, total_row, ac, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL, fmt="0.00").value = actual_total_formula
        style_cell(ws, total_row, vc, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL, fmt="0.00").value = var_total_formula


def _apply_actual_band_cf(
    ws, *, wp_rows: dict[str, int], external_rows: dict[str, int],
    total_row: int, roster_row: int,
    months: int,
    actual_col_start: int, variance_col_start: int,
) -> None:
    """Apply CF for actual band: variance reds + Actual TOTAL utilisation.

    EU per-year ceiling CF moved to Actuals-Hours (the enforcement surface);
    Capacity no longer carries the EU TOTAL row.
    """
    # ── CF: Variance band — red if ABS(variance) > CAPACITY_VARIANCE_FTE_RED ──
    all_variance_rows = (list(wp_rows.values())
                         + list(external_rows.values()))
    if all_variance_rows:
        var_start_cl = col_letter(variance_col_start)
        var_end_cl = col_letter(variance_col_start + months - 1)
        for vrow in all_variance_rows:
            var_range = f"{var_start_cl}{vrow}:{var_end_cl}{vrow}"
            ws.conditional_formatting.add(var_range, FormulaRule(
                formula=[
                    f"AND(ISNUMBER({var_start_cl}{vrow}),"
                    f"ABS({var_start_cl}{vrow})>{CAPACITY_VARIANCE_FTE_RED})"
                ],
                fill=CF_RED_FILL,
            ))

    # ── CF: Utilisation on Actual TOTAL row — red ≥ Max FTE, amber ≥ 0.8 × Max FTE ──
    max_fte_ref = cell_ref(SHEET_CONFIG, roster_row, COL_ROSTER_FTE)
    actual_start_cl = col_letter(actual_col_start)
    actual_end_cl = col_letter(actual_col_start + months - 1)
    actual_total_range = (
        f"{actual_start_cl}{total_row}:{actual_end_cl}{total_row}"
    )
    # Red first (highest CF priority → wins amber).
    ws.conditional_formatting.add(actual_total_range, FormulaRule(
        formula=[f"{actual_start_cl}{total_row}>={max_fte_ref}"],
        fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(actual_total_range, FormulaRule(
        formula=[
            f"{actual_start_cl}{total_row}"
            f">={CAPACITY_AMBER_LOAD}*{max_fte_ref}"
        ],
        fill=CF_AMBER_FILL,
    ))


def _write_person_grid(
    ws, r: int, member: dict, *,
    ctx: PersonGridContext,
    person_assignments: list[dict],
    wp_type_map: dict[str, str],
    assign_config_rows: dict[tuple[str, str], int],
    external_config_rows: dict[tuple[str, str], int],
    person_rows: dict[str, int],
    val_col_budget: int,
    val_col_diff: int,
    team_layout: dict,
    actual_col_start: int,
) -> tuple[int, dict]:
    """Write one person's grid section (header, WP rows, externals, TOTAL).

    Returns:
        (next_row, person_group_dict) where person_group_dict has keys
        header_row, wp_rows, external_rows, total_row.
    """
    state = _resolve_grid_state(
        start_row=r, member=member, ctx=ctx,
        person_assignments=person_assignments, person_rows=person_rows,
        team_layout=team_layout,
    )
    name = state["name"]
    roster_row = state["roster_row"]
    person_externals = state["person_externals"]
    layout = state["layout"]
    project_params = ctx.project_params
    month_col_start = ctx.month_col_start
    month_col_end = ctx.month_col_end
    val_col_sum = ctx.val_col_sum

    header_row = r
    _write_header_row(ws, r, roster_row)
    r += 1
    data_row_start = r

    r, wp_rows = _write_wp_rows(
        ws, start_row=r, name=name, person_assignments=person_assignments,
        wp_by_id=state["wp_by_id"], wp_idx=state["wp_idx"],
        wp_active_map=ctx.wp_active_map,
        wp_type_map=wp_type_map, assign_config_rows=assign_config_rows,
        config_first_wp_row=ctx.config_first_wp_row, layout=layout,
        project_params=project_params, month_col_start=month_col_start,
        month_col_end=month_col_end, val_col_sum=val_col_sum,
        val_col_budget=val_col_budget, val_col_diff=val_col_diff,
    )
    r, external_rows = _write_external_rows(
        ws, start_row=r, name=name, person_externals=person_externals,
        external_config_rows=external_config_rows, layout=layout,
        project_params=project_params, month_col_start=month_col_start,
        month_col_end=month_col_end, val_col_sum=val_col_sum,
        val_col_budget=val_col_budget, val_col_diff=val_col_diff,
    )

    # Empty-data placeholder: layout_start was bumped by 1 in
    # _resolve_grid_state so TOTAL sits one row lower.
    data_row_end = r - 1
    if data_row_end < data_row_start:
        data_row_end = data_row_start
        r += 1

    total_row = _write_total_row(
        ws, r=r, layout=layout, data_row_start=data_row_start,
        data_row_end=data_row_end, project_params=project_params,
        month_col_start=month_col_start, val_col_sum=val_col_sum,
        val_col_budget=val_col_budget, val_col_diff=val_col_diff,
    )
    r += 1

    _apply_planned_total_overalloc_cf(
        ws, total_row=total_row, roster_row=roster_row,
        month_col_start=month_col_start, month_col_end=month_col_end,
    )

    # Right-side band: Actual FTE | gutter | Variance. Variance offset by
    # months+1 (1-col gutter); utilisation CF sits on Actual TOTAL.
    months = project_params["months"]
    variance_col_start = actual_col_start + months + 1
    _setup_actual_band_header(
        ws, header_row=header_row, actual_col_start=actual_col_start,
        variance_col_start=variance_col_start, months=months,
    )
    _write_actual_pairs_for_data_rows(
        ws, wp_rows=wp_rows, external_rows=external_rows,
        person_assignments=person_assignments,
        person_externals=person_externals, wp_active_map=ctx.wp_active_map,
        project_params=project_params, months=months,
        month_col_start=month_col_start, actual_col_start=actual_col_start,
        variance_col_start=variance_col_start,
    )
    _write_actual_total_band(
        ws, total_row=total_row, wp_rows=wp_rows, external_rows=external_rows,
        months=months, actual_col_start=actual_col_start,
        variance_col_start=variance_col_start,
    )
    _apply_actual_band_cf(
        ws, wp_rows=wp_rows, external_rows=external_rows,
        total_row=total_row, roster_row=roster_row,
        months=months,
        actual_col_start=actual_col_start,
        variance_col_start=variance_col_start,
    )

    assert r == layout.next_row
    apply_person_row_outline(ws, layout)
    return layout.next_row, {
        "header_row": header_row,
        "wp_rows": wp_rows,
        "external_rows": external_rows,
        "total_row": total_row,
        "actual_col_start": actual_col_start,
        "variance_col_start": variance_col_start,
    }


def build_capacity_sheet(ws, team_layout, *,
                         project_params=None, work_packages=None,
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
    # ── Default-arg fallback (load-bearing for bare-call test fixtures) ──
    if any(v is None for v in [project_params, work_packages, team, external_commitments]):
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
    months = project_params["months"]
    month_col_end = _month_col_end(months)
    _vcols = val_columns(months)
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    # Right-side grouped band: Actual FTE columns, then a 1-col gutter, then
    # Variance FTE columns. Both bands are outline_level=1 so they collapse
    # via the Capacity sheet's summaryRight outline.
    actual_col_start = val_col_diff + 2  # 1-col gutter after Variance
    actual_col_end = actual_col_start + months - 1
    variance_col_start = actual_col_end + 2  # 1-col gutter after Actual
    variance_col_end = variance_col_start + months - 1
    # To-date block: 2 always-visible summary columns at the far right.
    # Sits outside the raw-monthly outline group so auditors always see
    # Actual/Variance-to-date even when the monthly bands are collapsed.
    # The former 4-col (+ util + ceil) layout was dropped because util/ceil
    # duplicate the planned-side CF signals already on this sheet.
    todate_col_start = variance_col_end + 2  # 1-col gutter after Variance
    todate_col_actual = todate_col_start
    todate_col_var = todate_col_start + 1
    todate_col_end = todate_col_var

    # Build WP type lookup — single entry per WP under the merged schema.
    wp_type_map: dict[str, str] = {
        wp["id"]: wp.get("type", "flat") for wp in work_packages
    }

    ws.sheet_properties.tabColor = "7030A0"

    # Outline groups: right-side Actual + Variance bands collapse right.
    # summaryBelow=False places the row-outline collapse button next to each
    # person header (above the block), so collapsing hides that person's
    # data rows beneath the header rather than above the next header.
    outline_pr = ws.sheet_properties.outlinePr
    if outline_pr is None:
        ws.sheet_properties.outlinePr = Outline(
            summaryRight=True, summaryBelow=False,
        )
    else:
        outline_pr.summaryRight = True
        outline_pr.summaryBelow = False

    # Raw monthly Actual + Variance bands are outline level 1 AND default
    # collapsed: per-month values are unreliable when the PI enters hours
    # with skipped months (T6). The to-date block provides the reliable
    # summary. User expands on-demand via the outline button.
    for mi in range(months):
        for col in (actual_col_start + mi, variance_col_start + mi):
            cd = ws.column_dimensions[col_letter(col)]
            cd.outline_level = 1
            cd.hidden = True

    person_rows = team_layout["person_rows"]
    assign_config_rows = team_layout["assign_config_rows"]
    external_config_rows = team_layout["external_config_rows"]

    # ── Band label row (row 1): "ACTUAL FTE (from hours)" and "VARIANCE FTE" ──
    band_label_row = 1
    ws.merge_cells(start_row=band_label_row, start_column=actual_col_start,
                   end_row=band_label_row, end_column=actual_col_end)
    style_cell(ws, band_label_row, actual_col_start,
               value=ACTUAL_BAND_LABEL,
               font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    ws.merge_cells(start_row=band_label_row, start_column=variance_col_start,
                   end_row=band_label_row, end_column=variance_col_end)
    style_cell(ws, band_label_row, variance_col_start,
               value=VARIANCE_BAND_LABEL,
               font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    # To-date block label (always-visible summary; outside the raw-monthly
    # outline group).
    ws.merge_cells(start_row=band_label_row, start_column=todate_col_start,
                   end_row=band_label_row, end_column=todate_col_end)
    style_cell(ws, band_label_row, todate_col_start,
               value="TO-DATE",
               font=HEADER_FONT, fill=LIGHT_GREEN_FILL, alignment=CENTER)
    ws.cell(row=band_label_row, column=todate_col_start).comment = sized_comment(
        "Progress snapshot through the `hours_actuals_through_month` "
        "named range (`'Actuals-Hours'!$D$1`). "
        "Actual PM: logged hours through that month / hours_per_fte_month "
        "(reads from Actuals-Hours cum-hours col). "
        "Variance PM: Plan − Actual on TOTAL row only; red when < 0. "
        "WP and external rows show values only where they carry signal."
    )

    # ── YEAR + MONTH HEADER ROWS (once, before first person group) ──
    r = band_label_row + 1
    year_header_row, month_header_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            (val_col_sum, "Sum"),
            (val_col_budget, "Budget"),
            (val_col_diff, "Variance"),
        ],
    )

    # Mirror year + month headers onto both Actual and Variance bands so
    # auditors see matching month labels across the right-side grouped band.
    _small = Font(name="Arial", size=8)
    for group_start, span in _calendar_year_groups(
        project_params["start"], months
    ):
        src_year_col = month_col_start + group_start
        label = ws.cell(row=year_header_row, column=src_year_col).value
        for band_start in (actual_col_start, variance_col_start):
            yr_start_col = band_start + group_start
            yr_end_col = yr_start_col + span - 1
            if span > 1:
                ws.merge_cells(start_row=year_header_row,
                               start_column=yr_start_col,
                               end_row=year_header_row,
                               end_column=yr_end_col)
            style_cell(ws, year_header_row, yr_start_col,
                       value=label, font=HEADER_FONT,
                       fill=LIGHT_BLUE_FILL)
            for c2 in range(yr_start_col + 1, yr_end_col + 1):
                style_cell(ws, year_header_row, c2, fill=LIGHT_BLUE_FILL)
    for mi in range(months):
        src_month_cell = ws.cell(row=month_header_row,
                                 column=month_col_start + mi)
        for band_start in (actual_col_start, variance_col_start):
            style_cell(ws, month_header_row, band_start + mi,
                       value=src_month_cell.value,
                       font=_small, fill=LIGHT_BLUE_FILL)

    # To-date block headers — one column label per metric.
    todate_labels = [
        (todate_col_actual, "Actual PM"),
        (todate_col_var, "Variance PM"),
    ]
    for col, label in todate_labels:
        style_cell(ws, year_header_row, col, fill=LIGHT_BLUE_FILL)
        style_cell(ws, month_header_row, col, value=label,
                   font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Explanatory comments on validation column headers
    ws.cell(row=year_header_row, column=val_col_sum).comment = sized_comment("Row-wide SUM of the monthly FTE cells for this WP/person. "
        "Should equal Budget — if not, Variance flags the gap.")
    ws.cell(row=year_header_row, column=val_col_budget).comment = sized_comment("Cross-ref (green font) to this assignment's Total FTE on the "
        "Config sheet. Edit Total FTE on Config, not here.")
    ws.cell(row=year_header_row, column=val_col_diff).comment = sized_comment("Budget − Sum. Should be zero when the distribution matches the "
        "assignment. Red highlight if negative (monthly total exceeds the "
        "committed Total FTE).")

    first_person_header_row = r

    # ── Build per-person assignment lookup (deduplicate by person+wp) ──
    _dedup: dict[tuple[str, str], dict] = {}
    for entry in assignments:
        _dedup[(entry["person"], entry["wp"])] = entry
    per_person_assignments: dict[str, list[dict]] = {}
    for entry in _dedup.values():
        per_person_assignments.setdefault(entry["person"], []).append(entry)

    # Pre-compute active months per WP so the per-person grid can clip each
    # assignment window to the WP's single contiguous active range.
    wp_active_map: dict[str, frozenset[int]] = {
        wp["id"]: wp_active_months(wp["id"], work_packages)
        for wp in work_packages
    }

    # ── PER-PERSON GROUPS ──
    person_groups: dict[str, dict] = {}

    ctx = PersonGridContext(
        work_packages=work_packages,
        external_commitments=external_commitments,
        wp_active_map=wp_active_map,
        project_params=project_params,
        month_col_start=month_col_start,
        month_col_end=month_col_end,
        val_col_sum=val_col_sum,
        config_first_wp_row=config_first_wp_row,
    )
    for member in team:
        name = member["name"]
        r, group = _write_person_grid(
            ws, r, member,
            ctx=ctx,
            person_assignments=per_person_assignments.get(name, []),
            wp_type_map=wp_type_map,
            assign_config_rows=assign_config_rows,
            external_config_rows=external_config_rows,
            person_rows=person_rows,
            val_col_budget=val_col_budget,
            val_col_diff=val_col_diff,
            team_layout=team_layout,
            actual_col_start=actual_col_start,
        )
        person_groups[name] = group

    # Freeze panes
    ws.freeze_panes = f"C{first_person_header_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "year_header_row": year_header_row,
        "month_header_row": month_header_row,
        "band_label_row": band_label_row,
        "actual_col_start": actual_col_start,
        "actual_col_end": actual_col_end,
        "variance_col_start": variance_col_start,
        "variance_col_end": variance_col_end,
        "todate_col_start": todate_col_start,
        "todate_col_actual": todate_col_actual,
        "todate_col_var": todate_col_var,
        "todate_col_end": todate_col_end,
        "person_groups": person_groups,
        "external_config_rows": external_config_rows,
        "assign_config_rows": assign_config_rows,
    }


def apply_capacity_todate_formulas(
    ws, capacity_layout, actuals_hours_layout, *, project_params,
):
    """Write Actual PM / Variance PM formulas onto the 2-col to-date block
    reserved by ``build_capacity_sheet``.

    Called after both Capacity and Actuals-Hours sheets have been built,
    so per-person row numbers are known on both sides. The block mirrors
    the ``hours_actuals_through_month`` window (`'Actuals-Hours'!$D$1`): everything
    is summed up to that month inclusive.

    Per-row fill rules (2-col to-date block):
      * WP row: Actual PM reads the Actuals-Hours cum-hours helper col
        divided by hours_per_fte_month (green cross-sheet ref).
        Variance PM is blank + grey — only meaningful at person-level.
      * External row: blank + grey across both cols (non-programme).
      * TOTAL row: Actual PM = SUM(WP Actual cells); Variance PM =
        (windowed plan sum on same-row planned cells) − Actual.

    CF: Variance PM < 0 red on TOTAL only (1·TEAM rules).
    """
    months = project_params["months"]
    month_col_start_l = col_letter(MONTH_COL_START)
    month_col_end_l = col_letter(MONTH_COL_START + months - 1)

    actual_col = capacity_layout["todate_col_actual"]
    var_col = capacity_layout["todate_col_var"]
    actual_cl = col_letter(actual_col)
    var_cl = col_letter(var_col)

    ah_groups = actuals_hours_layout.get("person_groups", {})
    cum_hours_col = actuals_hours_layout["cum_hours_col"]
    cum_cl = col_letter(cum_hours_col)

    def _wp_actual_formula(ah_row: int) -> str:
        # Single cross-sheet ref to Actuals-Hours helper col, divided by
        # hours_per_fte_month. The windowed SUM+INDEX lives on that helper
        # col, not here — see design doc.
        return (
            f"='{SHEET_ACTUALS_HOURS}'!${cum_cl}${ah_row}"
            f"/hours_per_fte_month"
        )

    def _windowed_plan_sum_expr(row: int) -> str:
        # Returns an *expression fragment* (no leading ``=``) so callers can
        # splice it into a compound formula. SUMs the planned monthly cells
        # on this row (typically a TOTAL or EU TOTAL row, whose per-month
        # cells already sum all WPs) across the first
        # ``hours_actuals_through_month`` columns.
        rng = f"${month_col_start_l}${row}:${month_col_end_l}${row}"
        return (
            f"IF(hours_actuals_through_month<=0,0,"
            f"SUM(${month_col_start_l}${row}:"
            f"INDEX({rng},1,hours_actuals_through_month)))"
        )

    for name, cap_group in capacity_layout["person_groups"].items():
        ah_group = ah_groups.get(name, {})
        ah_wp_rows: dict[str, int] = ah_group.get("wp_rows", {})

        # Programme WP rows — Actual PM only (cross-sheet); Variance PM blank.
        for wp_id, cap_row in cap_group["wp_rows"].items():
            ah_row = ah_wp_rows.get(wp_id)
            if ah_row is not None:
                style_cell(ws, cap_row, actual_col, font=GREEN_XREF,
                           fmt="0.00").value = _wp_actual_formula(ah_row)
            else:
                style_cell(ws, cap_row, actual_col, value=0,
                           font=BLACK_FORMULA, fmt="0.00")
            style_cell(ws, cap_row, var_col, value="",
                       font=BLACK_FORMULA, fill=GREY_FILL)

        # External rows — non-programme: blank/grey across the block.
        for ext_row in cap_group["external_rows"].values():
            for c in (actual_col, var_col):
                style_cell(ws, ext_row, c, value="",
                           font=BLACK_FORMULA, fill=GREY_FILL)

        # TOTAL row — Actual = SUM(WP actuals);
        # Variance = windowed plan sum − Actual.
        wp_rows = list(cap_group["wp_rows"].values())

        def _sum_wp_actuals(rows: list[int]) -> str:
            if not rows:
                return "=0"
            refs = ",".join(f"{actual_cl}{rr}" for rr in rows)
            return f"=SUM({refs})"

        total_row = cap_group["total_row"]

        style_cell(ws, total_row, actual_col, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL,
                   fmt="0.00").value = _sum_wp_actuals(wp_rows)
        var_formula = (
            f"={_windowed_plan_sum_expr(total_row)}"
            f"-{actual_cl}{total_row}"
        )
        style_cell(ws, total_row, var_col, font=BLACK_FORMULA,
                   fill=LIGHT_GREEN_FILL,
                   fmt="0.00").value = var_formula

        # CF: Variance PM red when < 0 on TOTAL row only.
        cell_ref_str = f"{var_cl}{total_row}"
        ws.conditional_formatting.add(
            cell_ref_str,
            FormulaRule(
                formula=[f"{cell_ref_str}<0"],
                fill=CF_RED_FILL,
            ),
        )
