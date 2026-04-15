"""Plan-Forecast sheet builder — forecast EUR aggregated from Actuals-Hours.

Mirrors the structure of ``plan.py`` but each monthly cell sums the
forecast-EUR leaf cells on the Actuals-Hours sheet for the persons
assigned to that WP.

Formula shape per WP × active-month cell::

    =SUM('Actuals-Hours'!<leafA>, 'Actuals-Hours'!<leafB>, ...)

- Leaves come from ``actuals_hours_layout["person_groups"]``: for each
  person assigned to the WP, locate that person's WP row, then pick
  the column at ``forecast_col_start + month_idx``.
- When no persons are assigned to the WP, the cell formula is ``=0``
  (explicit — avoids an empty ``SUM()``).
- Inactive months (WP not scheduled) are greyed out and empty.
"""

from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START, SHEET_CONFIG, SHEET_ACTUALS_HOURS, SHEET_PLAN_FORECAST,
    build_year_month_headers, month_col_end as _month_col_end, val_columns,
    wp_active_months,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    GREY_FILL, LIGHT_GREEN_FILL,
    CF_GREY_FILL, CF_RED_FILL, style_cell, col_letter, write_section_header,
    sized_comment,
)


# Excel quotes sheet names containing non-alphanumeric chars (e.g. hyphens).
_AH_QUOTED = f"'{SHEET_ACTUALS_HOURS}'"
EUR_FMT = "\u20ac#,##0"


def _forecast_leaf_refs(wp_id, month_idx, *, assignments, actuals_hours_layout):
    """Return list of ``'Actuals-Hours'!<ref>`` strings for a WP × month.

    One ref per person assigned to ``wp_id``. The column is the same for
    all refs (forecast_col_start + month_idx); the row differs per person
    (their WP row on Actuals-Hours). Persons assigned to the WP but
    without a matching WP row on Actuals-Hours (shouldn't happen in
    practice — they share the same assignment source) are skipped.
    """
    forecast_col_start = actuals_hours_layout["forecast_col_start"]
    person_groups = actuals_hours_layout["person_groups"]
    col = forecast_col_start + month_idx
    cl = col_letter(col)

    # Deterministic person order: walk assignments for this WP; dedup.
    seen: set[str] = set()
    refs: list[str] = []
    for a in assignments:
        if a["wp"] != wp_id:
            continue
        name = a["person"]
        if name in seen:
            continue
        seen.add(name)
        group = person_groups.get(name)
        if group is None:
            continue
        wp_row = group["wp_rows"].get(wp_id)
        if wp_row is None:
            continue
        refs.append(f"{_AH_QUOTED}!{cl}{wp_row}")
    return refs


def build_plan_forecast_sheet(
    ws,
    *,
    project_params,
    work_packages,
    wp_schedule,
    assignments,
    config_first_wp_row,
    actuals_hours_layout,
):
    """Build the Plan-Forecast sheet.

    Args:
        ws: Target worksheet.
        project_params: dict with ``title``, ``start``, ``months``, ``rate``.
        work_packages: list of WP dicts.
        wp_schedule: list of schedule entries — drives active-month greying.
        assignments: list of person-WP assignment dicts; determines which
            persons' forecast-EUR leaves flow into each WP cell.
        config_first_wp_row: Row number of the first WP on Config sheet
            (for ID/Name/Ceiling cross-refs).
        actuals_hours_layout: dict returned by ``build_actuals_hours_sheet``.
            Must contain ``forecast_col_start`` and ``person_groups``.

    Returns:
        Layout dict with WP row positions + total row.
    """
    if assignments is None:
        assignments = []

    months = project_params["months"]
    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    m_end = _month_col_end(months)

    _vcols = val_columns(months)
    val_col_sum = _vcols["sum"]
    val_col_ceil = _vcols["budget"]  # reused slot — "Ceiling" here, not "Budget FTE"
    val_col_diff = _vcols["diff"]

    ws.sheet_properties.tabColor = "00B050"  # green — forecast view

    cfg_wp_first = config_first_wp_row
    cfg_wp_last = cfg_wp_first + num_wps - 1

    # ── Row 1: compact header (Title / Start / Ceiling) ──
    # Title block
    style_cell(ws, 1, 1, "Title:", font=HEADER_FONT, border=False)
    style_cell(
        ws, 1, 2,
        font=Font(name="Arial", bold=True, size=14, color="008000"),
        border=False,
    ).value = f"={SHEET_CONFIG}!B2"

    # Start block
    style_cell(ws, 1, 4, "Start:", font=HEADER_FONT, border=False)
    style_cell(
        ws, 1, 5, font=GREEN_XREF, border=False, fmt="MMM YYYY",
    ).value = "=project_start"

    # Ceiling block — sum of WP ceilings (same source as Plan!G1)
    style_cell(ws, 1, 6, "Ceiling:", font=HEADER_FONT, border=False)
    style_cell(
        ws, 1, 7, font=GREEN_XREF, border=False, fmt=EUR_FMT,
    ).value = f"=SUM({SHEET_CONFIG}!E{cfg_wp_first}:E{cfg_wp_last})"

    # -- FORECAST EUR block (starts at row 3) --
    r = 3
    write_section_header(
        ws, r, "FORECAST EUR (from hours)", merge_end=val_col_diff,
    )
    r += 1

    year_row, month_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            (val_col_sum, "Sum"),
            (val_col_ceil, "Ceiling"),
            (val_col_diff, "Variance"),
        ],
    )
    del year_row, month_row  # positions captured in return dict below
    first_wp_row = r

    # ── WP rows ──
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = cfg_wp_first + i
        wp_id = wp["id"]

        # A: WP ID cross-ref  B: WP Name cross-ref
        style_cell(ws, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(ws, row, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        active = wp_active_months(wp_id, wp_schedule)
        for m in range(months):
            c = month_col_start + m
            month_num = m + 1
            if month_num in active:
                refs = _forecast_leaf_refs(
                    wp_id, m,
                    assignments=assignments,
                    actuals_hours_layout=actuals_hours_layout,
                )
                if refs:
                    formula = f"=SUM({', '.join(refs)})"
                else:
                    formula = "=0"
                style_cell(
                    ws, row, c, font=GREEN_XREF, fmt=EUR_FMT,
                ).value = formula
            else:
                style_cell(
                    ws, row, c, "", font=GREEN_XREF,
                    fill=GREY_FILL, fmt=EUR_FMT,
                )

        # Validation columns
        sum_range = (
            f"{col_letter(month_col_start)}{row}:{col_letter(m_end)}{row}"
        )
        style_cell(
            ws, row, val_col_sum, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = f"=SUM({sum_range})"
        style_cell(
            ws, row, val_col_ceil, font=GREEN_XREF, fmt=EUR_FMT,
        ).value = f"={SHEET_CONFIG}!E{config_row}"
        style_cell(
            ws, row, val_col_diff, font=BLACK_FORMULA, fmt=EUR_FMT,
        ).value = (
            f"={col_letter(val_col_ceil)}{row}"
            f"-{col_letter(val_col_sum)}{row}"
        )

    last_wp_row = first_wp_row + num_wps - 1
    total_row = last_wp_row + 1

    # ── TOTAL row ──
    style_cell(ws, total_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(
        ws, total_row, 2, "MONTHLY TOTAL",
        font=HEADER_FONT, fill=LIGHT_GREEN_FILL,
    )
    for m in range(months):
        c = month_col_start + m
        cl = col_letter(c)
        style_cell(
            ws, total_row, c, font=BLACK_FORMULA,
            fill=LIGHT_GREEN_FILL, fmt=EUR_FMT,
        ).value = f"=SUM({cl}{first_wp_row}:{cl}{last_wp_row})"

    # Validation-column totals
    for vcol in (val_col_sum, val_col_ceil, val_col_diff):
        cl = col_letter(vcol)
        style_cell(
            ws, total_row, vcol, font=BLACK_FORMULA,
            fill=LIGHT_GREEN_FILL, fmt=EUR_FMT,
        ).value = f"=SUM({cl}{first_wp_row}:{cl}{last_wp_row})"

    # Red CF on Variance when < 0 (Sum has exceeded Ceiling)
    diff_col_l = col_letter(val_col_diff)
    cf_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cf_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))

    # Cosmetic grey CF: forecast-EUR cells for months past
    # ``hours_actuals_through_month`` — no hours logged yet, so the forecast
    # aggregates zero contributions and should read as empty/unreliable
    # rather than a firm projection. Spans the WP block + TOTAL row.
    ms = col_letter(month_col_start)
    me = col_letter(m_end)
    forecast_body_range = f"{ms}{first_wp_row}:{me}{total_row}"
    ws.conditional_formatting.add(forecast_body_range, FormulaRule(
        formula=[
            f"COLUMN({ms}{first_wp_row})-COLUMN(${ms}$1)+1"
            f">hours_actuals_through_month"
        ],
        fill=CF_GREY_FILL,
    ))

    # Comment on the ceiling header — explain provenance
    ws.cell(row=1, column=7).comment = sized_comment("Sum of WP ceilings from Config. Same value as Plan!G1 and "
        "Actuals!G1 — verify all three agree.")

    # Freeze panes: lock WP ID/Name cols + header rows
    ws.freeze_panes = f"C{first_wp_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "config_first_wp_row": cfg_wp_first,
        "config_last_wp_row": cfg_wp_last,
        "forecast_header_row": 3,
        "forecast_first_wp_row": first_wp_row,
        "forecast_last_wp_row": last_wp_row,
        "forecast_total_row": total_row,
    }


# Re-export for convenience
__all__ = ["build_plan_forecast_sheet", "SHEET_PLAN_FORECAST"]
