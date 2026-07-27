"""Trends sheet builder — WP burn rate, cumulative chart, person load.

Dashboard — WP Burn Rate & Exhaustion table, cumulative EUR LineChart
(Planned / Actual / Forecast / Ceiling / RP End; Actual + Forecast series
clip via ``NA()`` past their respective through-month),
**Reconciliation: Forecast vs Actual EUR** section, per-WP Gap table, a
gap-over-time LineChart, and a Person Load table.

Role: read-only; ``current_month`` (B1) is the only user input.

**Non-obvious invariant:** the **Reconciled through** value
(``=MIN(eur_actuals_through_month, hours_actuals_through_month)``) is a
**local helper cell on Trends** at ``$B$<recon_section_row+4>`` — it is
**not a named range**; per-WP Gap SUMs and the gap chart both read it via
absolute same-sheet ref.

## Conditional formatting

- Gap %: **red** > |10%|, **amber** > |5%|.
- Variance to-date %: **red** > |20%|, **amber** > |10%|.
"""

from dateutil.relativedelta import relativedelta
from openpyxl.chart import LineChart
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_PLAN,
    SHEET_CONFIG,
    SHEET_ACTUALS,
    SHEET_CAPACITY,
    TREND_OK_BAND,
    TREND_OVER_THRESHOLD,
    TREND_WATCH_THRESHOLD,
    TREND_SLOW_THRESHOLD,
    COL_WP_START,
    COL_WP_END,
    month_col_end as _month_col_end,
    val_columns,
    wp_month_bounds,
)
from openpyxl.utils import get_column_letter
from budget_planner.styles import (
    BLACK_FORMULA,
    GREEN_XREF,
    HEADER_FONT,
    LIGHT_BLUE_FILL,
    LIGHT_GREEN_FILL,
    CENTER,
    style_cell,
    col_letter,
    write_section_header,
    auto_fit_columns,
    sized_comment,
)

# CF fills imported from styles — DXF context requires bgColor
from budget_planner.styles import CF_GREEN_FILL as _GREEN_CF_FILL
from budget_planner.styles import CF_AMBER_FILL as _AMBER_CF_FILL
from budget_planner.styles import CF_RED_FILL as _RED_CF_FILL


def _resolve_trends_layout(layout, actuals_layout, project_params, work_packages, team):
    """Resolve every column letter and row anchor for the Trends sheet
    into a single dict consumed by the helpers below.

    Pure function — no worksheet writes. All Trends row anchors are
    statically derivable from ``num_wps`` + ``len(team)``; lifting them
    out of ``build_trends_sheet`` keeps the orchestrator readable.
    """
    wp_start_col_letter = get_column_letter(COL_WP_START)
    wp_end_col_letter = get_column_letter(COL_WP_END)
    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    month_col_end = _month_col_end(project_params["months"])
    val_col_budget = val_columns(project_params["months"])["budget"]

    # Posted-cumulative helpers on Actuals (per Phase 2.4 rewire).
    # `posted_idx`/`posted_value`/`slope` replace the Stage A `cum_col` bandage:
    # Burn-rate Col D + Col G now read posted_value directly (guarded by
    # posted_idx=0 → "—"/"No data"); reconciliation + 2-anchor forecast use
    # the slope projection. See design §3 + §9.
    posted_idx_col_letter = col_letter(actuals_layout["posted_idx_col"])
    posted_value_col_letter = col_letter(actuals_layout["posted_value_col"])
    slope_col_letter = col_letter(actuals_layout["slope_col"])
    ceiling_col_letter = col_letter(val_col_budget)
    ms = col_letter(month_col_start)
    me = col_letter(month_col_end)

    # Section 1: WP Burn Rate & Exhaustion
    section1_row = 3
    header_row = 4
    first_wp_data_row = header_row + 1
    last_wp_data_row = first_wp_data_row + num_wps - 1
    total_row = last_wp_data_row + 1

    # Cumulative chart anchor (chart height 7.5 cm ≈ 14 rows)
    chart_anchor_row = total_row + 2

    # Section 2a: Reconciliation block
    recon_section_row = chart_anchor_row + 14
    recon_health_eur_through_row = recon_section_row + 2
    recon_health_hours_through_row = recon_section_row + 3
    recon_health_reconciled_row = recon_section_row + 4
    recon_table_header_row = recon_section_row + 6
    recon_first_wp_row = recon_table_header_row + 1
    recon_last_wp_row = recon_first_wp_row + num_wps - 1
    recon_total_row = recon_last_wp_row + 1

    # Section 2b: Gap chart sits two rows below the recon TOTAL row.
    recon_gap_chart_anchor_row = recon_total_row + 2

    # Section 3: Person Load (gap chart anchor + 14 rows chart + 1 spacer)
    person_section_row = recon_gap_chart_anchor_row + 15
    person_header_row = person_section_row + 1
    person_first_row = person_header_row + 1
    person_last_row = person_first_row + len(team) - 1

    # Hidden helper rows (chart data, 2-row gap below person load)
    helper_start_row = person_last_row + 2
    helper_row_planned = helper_start_row
    helper_row_actual = helper_start_row + 1
    helper_row_forecast = helper_start_row + 2
    helper_row_ceiling = helper_start_row + 3
    helper_row_labels = helper_row_ceiling + 1
    helper_row_rp = helper_row_labels + 1
    helper_row_gap = helper_row_rp + 1
    helper_row_zero = helper_row_gap + 1

    return {
        "wp_start_col_letter": wp_start_col_letter,
        "wp_end_col_letter": wp_end_col_letter,
        "month_col_start": month_col_start,
        "num_wps": num_wps,
        "month_col_end": month_col_end,
        "val_col_budget": val_col_budget,
        "posted_idx_col_letter": posted_idx_col_letter,
        "posted_value_col_letter": posted_value_col_letter,
        "slope_col_letter": slope_col_letter,
        "ceiling_col_letter": ceiling_col_letter,
        "ms": ms,
        "me": me,
        "section1_row": section1_row,
        "header_row": header_row,
        "first_wp_data_row": first_wp_data_row,
        "last_wp_data_row": last_wp_data_row,
        "total_row": total_row,
        "chart_anchor_row": chart_anchor_row,
        "recon_section_row": recon_section_row,
        "recon_health_eur_through_row": recon_health_eur_through_row,
        "recon_health_hours_through_row": recon_health_hours_through_row,
        "recon_health_reconciled_row": recon_health_reconciled_row,
        "recon_table_header_row": recon_table_header_row,
        "recon_first_wp_row": recon_first_wp_row,
        "recon_last_wp_row": recon_last_wp_row,
        "recon_total_row": recon_total_row,
        "recon_gap_chart_anchor_row": recon_gap_chart_anchor_row,
        "person_section_row": person_section_row,
        "person_header_row": person_header_row,
        "person_first_row": person_first_row,
        "person_last_row": person_last_row,
        "helper_row_planned": helper_row_planned,
        "helper_row_actual": helper_row_actual,
        "helper_row_forecast": helper_row_forecast,
        "helper_row_ceiling": helper_row_ceiling,
        "helper_row_labels": helper_row_labels,
        "helper_row_rp": helper_row_rp,
        "helper_row_gap": helper_row_gap,
        "helper_row_zero": helper_row_zero,
    }


def _write_burn_rate_table(ws, tlayout, layout, actuals_layout, work_packages):
    """Section 1: WP Burn Rate & Exhaustion — section header, column headers,
    and per-WP data rows (cols A–G)."""
    section1_row = tlayout["section1_row"]
    header_row = tlayout["header_row"]
    first_wp_data_row = tlayout["first_wp_data_row"]
    num_wps = tlayout["num_wps"]
    posted_idx_col_letter = tlayout["posted_idx_col_letter"]
    posted_value_col_letter = tlayout["posted_value_col_letter"]
    slope_col_letter = tlayout["slope_col_letter"]
    ceiling_col_letter = tlayout["ceiling_col_letter"]
    wp_start_col_letter = tlayout["wp_start_col_letter"]
    wp_end_col_letter = tlayout["wp_end_col_letter"]
    ms = tlayout["ms"]
    me = tlayout["me"]

    # ================================================================
    # Section 1: WP Burn Rate & Exhaustion Table
    # ================================================================
    write_section_header(ws, section1_row, "WP BURN RATE & EXHAUSTION", merge_end=7)

    # Column headers (row 4)
    headers = [
        "WP ID",
        "WP Name",
        "% Time Elapsed",
        "% Budget EUR",
        "Burn Delta",
        "Status",
        "Exhaustion Mo.",
    ]
    for ci, h in enumerate(headers, 1):
        style_cell(ws, header_row, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Cell comment on Exhaustion Mo. header
    ws.cell(row=header_row, column=7).comment = sized_comment(
        "Projection is linear based on quarterly actuals. Accuracy improves after Q2."
    )

    # WP data rows
    for i, wp in enumerate(work_packages):
        row = first_wp_data_row + i
        config_row = layout["config_first_wp_row"] + i
        actuals_row = actuals_layout["first_wp_row"] + i
        actuals_month_range = f"{SHEET_ACTUALS}!{ms}{actuals_row}:{me}{actuals_row}"

        # Col A: WP ID (green cross-sheet ref)
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"

        # Col B: WP Name (green cross-sheet ref)
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        # WP start/end: direct cell refs into WP BUDGET cols H/I under the
        # single-range schema (one row per WP, no MINIFS/MAXIFS indirection).
        wp_start = f"{SHEET_CONFIG}!{wp_start_col_letter}{config_row}"
        wp_end = f"{SHEET_CONFIG}!{wp_end_col_letter}{config_row}"

        # Col C: % Time Elapsed
        time_formula = f"MIN(1,(current_month-{wp_start}+1)/({wp_end}-{wp_start}+1))"
        style_cell(
            ws, row, 3, font=BLACK_FORMULA, fmt="0.0%"
        ).value = f'=IF(current_month<{wp_start},"\u2014",{time_formula})'

        # Col D: % Budget EUR
        # Stage A bandage: Posted -i now returns "" for "no postings yet" (design §2);
        # the +0/IFERROR wrapper coerces "" → 0 until Stage B rewires Trends to read
        # posted_value / posted_idx / slope cells directly.
        # Cross-sheet refs to the per-WP posted-cumulative helpers on
        # Actuals (Phase 2.4 rewire \u2014 replaces Stage A `cum_col` bandage).
        posted_idx_ref = f"{SHEET_ACTUALS}!{posted_idx_col_letter}{actuals_row}"
        posted_value_ref = f"{SHEET_ACTUALS}!{posted_value_col_letter}{actuals_row}"
        slope_ref = f"{SHEET_ACTUALS}!{slope_col_letter}{actuals_row}"

        # Col D: posted_value is "" when posted_idx=0 (design \u00a72
        # sentinel) -- outer guard on posted_idx=0 -> "\u2014" so
        # the division never sees the empty-string sentinel.
        budget_formula = f"{posted_value_ref}/{SHEET_ACTUALS}!{ceiling_col_letter}{actuals_row}"
        style_cell(
            ws, row, 4, font=BLACK_FORMULA, fmt="0.0%"
        ).value = f'=IF(current_month<{wp_start},"\u2014",IF({posted_idx_ref}=0,"\u2014",{budget_formula}))'

        # Col E: Burn Delta -- mirrors Col D's posted_idx=0 short-circuit
        # because the arithmetic D-C otherwise propagates "\u2014" - C as
        # #VALUE! once a WP starts but has no postings yet.
        style_cell(
            ws, row, 5, font=BLACK_FORMULA, fmt="0.0%"
        ).value = f'=IF(current_month<{wp_start},"\u2014",IF({posted_idx_ref}=0,"\u2014",D{row}-C{row}))'

        # Col F: Status -- same posted_idx=0 short-circuit (ABS(E)
        # propagates the upstream #VALUE! otherwise).
        status_formula = (
            f'IF(ABS(E{row})<={TREND_OK_BAND},"OK",'
            f'IF(E{row}>{TREND_OVER_THRESHOLD},"OVER",'
            f'IF(E{row}>{TREND_WATCH_THRESHOLD},"WATCH",'
            f'IF(E{row}<{TREND_SLOW_THRESHOLD},"SLOW","OK"))))'
        )
        style_cell(
            ws, row, 6, font=BLACK_FORMULA
        ).value = f'=IF(current_month<{wp_start},"\u2014",IF({posted_idx_ref}=0,"\u2014",{status_formula}))'

        # Col G: Exhaustion Mo. \u2014 projection = wp_start + ROUND(budget/slope)-1.
        # Per design \u00a73 the canonical burn rate is the per-project-month
        # `slope` (avg-burn = posted_value / posted_idx) -- replaces the
        # filled-cell-average heuristic (cum/COUNTIF) with the actuals-tab
        # source of truth. Guarded twice: posted_idx=0 -> "No data";
        # D=0 -> "\u2014". `actuals_month_range` is no longer consulted because
        # posted_idx encodes "have we posted yet" without a separate
        # COUNTIF count.
        del actuals_month_range
        exhaustion_formula = (
            f'IF({posted_idx_ref}=0,"No data",'
            f'IF(D{row}=0,"\u2014",'
            f"{wp_start}+ROUND("
            f"{SHEET_ACTUALS}!{ceiling_col_letter}{actuals_row}"
            f"/{slope_ref},0)-1))"
        )
        style_cell(
            ws, row, 7, font=BLACK_FORMULA, fmt="0"
        ).value = f'=IF(current_month<{wp_start},"\u2014",{exhaustion_formula})'

    _ = num_wps  # parity hint: helper width drives WP loop count


def _apply_burn_rate_cf(ws, tlayout, work_packages):
    """CF on Burn Rate columns E (Burn Delta), F (Status), and G (per-WP
    Exhaustion Month). Add-order is sacred — first added = highest priority.
    """
    first_wp_data_row = tlayout["first_wp_data_row"]
    last_wp_data_row = tlayout["last_wp_data_row"]
    num_wps = tlayout["num_wps"]

    # -- Conditional formatting: Burn Delta (col E, rows 5-12) --
    e_range = f"E{first_wp_data_row}:E{last_wp_data_row}"
    e_ref = f"E{first_wp_data_row}"

    # Green first (lowest priority) — ABS(delta) <= TREND_OK_BAND
    ws.conditional_formatting.add(
        e_range,
        FormulaRule(
            formula=[f"ABS({e_ref})<={TREND_OK_BAND}"],
            fill=_GREEN_CF_FILL,
        ),
    )
    # Amber: overspend (delta > TREND_WATCH_THRESHOLD AND delta <= TREND_OVER_THRESHOLD)
    ws.conditional_formatting.add(
        e_range,
        FormulaRule(
            formula=[f"AND({e_ref}>{TREND_WATCH_THRESHOLD},{e_ref}<={TREND_OVER_THRESHOLD})"],
            fill=_AMBER_CF_FILL,
        ),
    )
    # Amber: underspend (delta < TREND_SLOW_THRESHOLD)
    ws.conditional_formatting.add(
        e_range,
        FormulaRule(
            formula=[f"{e_ref}<{TREND_SLOW_THRESHOLD}"],
            fill=_AMBER_CF_FILL,
        ),
    )
    # Red: delta > TREND_OVER_THRESHOLD
    ws.conditional_formatting.add(
        e_range,
        FormulaRule(
            formula=[f"{e_ref}>{TREND_OVER_THRESHOLD}"],
            fill=_RED_CF_FILL,
        ),
    )

    # -- Conditional formatting: Status (col F, rows 5-12) --
    f_range = f"F{first_wp_data_row}:F{last_wp_data_row}"
    f_ref = f"F{first_wp_data_row}"

    # Green: "OK"
    ws.conditional_formatting.add(
        f_range,
        FormulaRule(
            formula=[f'{f_ref}="OK"'],
            fill=_GREEN_CF_FILL,
        ),
    )
    # Amber: "WATCH" or "SLOW"
    ws.conditional_formatting.add(
        f_range,
        FormulaRule(
            formula=[f'OR({f_ref}="WATCH",{f_ref}="SLOW")'],
            fill=_AMBER_CF_FILL,
        ),
    )
    # Red: "OVER"
    ws.conditional_formatting.add(
        f_range,
        FormulaRule(
            formula=[f'{f_ref}="OVER"'],
            fill=_RED_CF_FILL,
        ),
    )

    # -- Conditional formatting: Exhaustion Month (col G) — per WP row --
    # Each WP has a different end month; use literal value from wp_month_bounds()
    # since CF formulas can't easily use MINIFS/MAXIFS.
    for i in range(num_wps):
        row = first_wp_data_row + i
        wp_id = work_packages[i]["id"]
        _, wp_end_val = wp_month_bounds(wp_id, work_packages)
        g_cell = f"G{row}:G{row}"
        g_ref = f"G{row}"
        # Green first: exhaustion >= wp_end
        ws.conditional_formatting.add(
            g_cell,
            FormulaRule(
                formula=[f"AND(ISNUMBER({g_ref}),{g_ref}>={wp_end_val})"],
                fill=_GREEN_CF_FILL,
            ),
        )
        # Amber: exhaustion within 3 months of wp_end
        ws.conditional_formatting.add(
            g_cell,
            FormulaRule(
                formula=[f"AND(ISNUMBER({g_ref}),{g_ref}<{wp_end_val},{g_ref}>={wp_end_val}-3)"],
                fill=_AMBER_CF_FILL,
            ),
        )
        # Red: exhaustion < wp_end - 3
        ws.conditional_formatting.add(
            g_cell,
            FormulaRule(
                formula=[f"AND(ISNUMBER({g_ref}),{g_ref}<{wp_end_val}-3)"],
                fill=_RED_CF_FILL,
            ),
        )


def _write_burn_rate_total(ws, tlayout, actuals_layout):
    """TOTAL row of the Burn Rate table — weighted-average burn delta in col E."""
    first_wp_data_row = tlayout["first_wp_data_row"]
    num_wps = tlayout["num_wps"]
    total_row = tlayout["total_row"]
    ceiling_col_letter = tlayout["ceiling_col_letter"]
    act_first = actuals_layout["first_wp_row"]

    # -- TOTAL row --
    style_cell(ws, total_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 3, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 4, "", fill=LIGHT_GREEN_FILL)
    # Weighted average burn delta — individual IF() terms avoid text*number #VALUE!
    terms = []
    for i in range(num_wps):
        wr = first_wp_data_row + i
        ar = act_first + i
        terms.append(f"IF(ISNUMBER(E{wr}),E{wr}*{SHEET_ACTUALS}!{ceiling_col_letter}{ar},0)")
    style_cell(ws, total_row, 5, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="0.0%").value = (
        "=(" + "+".join(terms) + ")/partner_ceiling"
    )
    style_cell(ws, total_row, 6, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 7, "", fill=LIGHT_GREEN_FILL)


def _write_reconciliation_section(ws, tlayout, layout, actuals_layout):
    """Section 2a: Reconciliation block — section header, three health rows,
    per-WP table (header + WP rows + TOTAL row).

    Layout inside reconciliation block (offsets from recon_section_row):
      +0: section header
      +1: blank
      +2: Actuals EUR through
      +3: Hours through
      +4: Reconciled through
      +5: blank
      +6: table header
      +7..+6+num_wps: WP data rows
      +7+num_wps:     TOTAL row
      +8+num_wps:     blank spacer
    """
    num_wps = tlayout["num_wps"]
    ms = tlayout["ms"]
    me = tlayout["me"]
    recon_section_row = tlayout["recon_section_row"]
    recon_health_eur_through_row = tlayout["recon_health_eur_through_row"]
    recon_health_hours_through_row = tlayout["recon_health_hours_through_row"]
    recon_health_reconciled_row = tlayout["recon_health_reconciled_row"]
    recon_table_header_row = tlayout["recon_table_header_row"]
    recon_first_wp_row = tlayout["recon_first_wp_row"]
    recon_last_wp_row = tlayout["recon_last_wp_row"]
    recon_total_row = tlayout["recon_total_row"]
    posted_idx_col_letter = tlayout["posted_idx_col_letter"]
    posted_value_col_letter = tlayout["posted_value_col_letter"]
    slope_col_letter = tlayout["slope_col_letter"]

    # ================================================================
    # Section 2a: Reconciliation — Forecast vs Actual EUR
    # ================================================================
    # Chart is anchored 2 rows below total. Chart height is 7.5 cm (~14
    # rows at default row height), so the Reconciliation block starts at
    # chart_anchor_row + 14 (immediately after the chart).
    write_section_header(
        ws,
        recon_section_row,
        "RECONCILIATION: FORECAST vs ACTUAL EUR",
        merge_end=5,
    )

    # Health row 1: Actuals EUR through — green cross-sheet (named range → Config)
    style_cell(ws, recon_health_eur_through_row, 1, value="Actuals EUR through", font=HEADER_FONT)
    style_cell(ws, recon_health_eur_through_row, 2, font=GREEN_XREF, fmt="0").value = "=eur_actuals_through_month"

    # Health row 2: Hours through — green cross-sheet (named range → Config)
    style_cell(ws, recon_health_hours_through_row, 1, value="Hours through", font=HEADER_FONT)
    style_cell(ws, recon_health_hours_through_row, 2, font=GREEN_XREF, fmt="0").value = "=hours_actuals_through_month"

    # Health row 3: Reconciled through — same-sheet derived name (black)
    style_cell(ws, recon_health_reconciled_row, 1, value="Reconciled through", font=HEADER_FONT)
    style_cell(
        ws, recon_health_reconciled_row, 2, font=BLACK_FORMULA, fmt="0"
    ).value = "=MIN(eur_actuals_through_month,hours_actuals_through_month)"

    # Per-WP reconciliation table header
    recon_headers = ["WP", "Cum Forecast EUR", "Cum Actual EUR", "Gap EUR", "Gap %"]
    for ci, h in enumerate(recon_headers, 1):
        style_cell(ws, recon_table_header_row, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Data refs
    act_first_wp_row = actuals_layout["first_wp_row"]
    cfg_first_wp = layout["config_first_wp_row"]

    for i in range(num_wps):
        row = recon_first_wp_row + i
        config_row = cfg_first_wp + i
        actuals_wp_row = act_first_wp_row + i

        # Col A: WP ID (green cross-sheet ref to Config)
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"

        # Col B: Cumulative Forecast EUR through reconciled_through_month.
        # Per-WP slope projection on Actuals: linear extrapolation from the
        # WP's posted-cumulative anchor out to reconciled_through.
        # MAX(0, ...) clamps a negative projection (correction postings).
        # Empty-string sentinel handled by posted_idx=0 -> 0 (no projection).
        posted_idx_ref = f"{SHEET_ACTUALS}!{posted_idx_col_letter}{actuals_wp_row}"
        posted_value_ref = f"{SHEET_ACTUALS}!{posted_value_col_letter}{actuals_wp_row}"
        slope_ref = f"{SHEET_ACTUALS}!{slope_col_letter}{actuals_wp_row}"
        style_cell(ws, row, 2, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
            f"=IF({posted_idx_ref}=0,0,"
            f"MAX(0,{posted_value_ref}+{slope_ref}*"
            f"($B${recon_health_reconciled_row}-{posted_idx_ref})))"
        )

        # Col C: Cumulative Actual EUR through reconciled_through_month.
        # Under posted-cumulative semantics the Actuals row holds running-
        # total snapshots, not per-month deltas \u2014 SUMming them across the
        # window double-counts. Instead, INDEX into the latest filled
        # numeric cell at-or-before reconciled_through_month: the column
        # offset is the max of (1-based position * filled * within-window),
        # which already IS the posted-cumulative-to-date by definition.
        # "filled" is ({range}<>""), not ISNUMBER({range}): Excel gives
        # ISNUMBER(range) no array context outside CSE, collapsing the scan
        # to 0 (LibreOffice masks this). COLUMN(range) is unaffected and stays.
        # IFERROR guards INDEX(..., 0) when no numeric cells fall inside
        # the window.
        act_range_full = f"{SHEET_ACTUALS}!${ms}${actuals_wp_row}:${me}${actuals_wp_row}"
        act_start_ref = f"{SHEET_ACTUALS}!${ms}${actuals_wp_row}"
        style_cell(ws, row, 3, font=GREEN_XREF, fmt="\u20ac#,##0").value = (
            f"=IF($B${recon_health_reconciled_row}<=0,0,"
            f"IFERROR(INDEX({act_range_full},1,"
            f"SUMPRODUCT(MAX("
            f"(COLUMN({act_range_full})-COLUMN({act_start_ref})+1)"
            f'*({act_range_full}<>"")'
            f"*((COLUMN({act_range_full})-COLUMN({act_start_ref})+1)"
            f"<=$B${recon_health_reconciled_row})"
            f"))),0))"
        )

        # Col D: Gap EUR = Actual - Forecast (same-sheet derived → black)
        style_cell(ws, row, 4, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = f"=C{row}-B{row}"

        # Col E: Gap % = D / B, guarded against zero forecast.
        style_cell(ws, row, 5, font=BLACK_FORMULA, fmt="0.0%").value = f"=IFERROR(D{row}/B{row},0)"

    # TOTAL row immediately after last WP row
    style_cell(ws, recon_total_row, 1, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    style_cell(
        ws, recon_total_row, 2, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0"
    ).value = f"=SUM(B${recon_first_wp_row}:B${recon_last_wp_row})"
    style_cell(
        ws, recon_total_row, 3, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0"
    ).value = f"=SUM(C${recon_first_wp_row}:C${recon_last_wp_row})"
    style_cell(
        ws, recon_total_row, 4, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="\u20ac#,##0"
    ).value = f"=C{recon_total_row}-B{recon_total_row}"
    style_cell(
        ws, recon_total_row, 5, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="0.0%"
    ).value = f"=IFERROR(D{recon_total_row}/B{recon_total_row},0)"


def _apply_reconciliation_cf(ws, tlayout):
    """CF on the Reconciliation Gap % column (E, data rows + TOTAL row).
    Red added FIRST so it wins at |gap| > 10%.
    """
    recon_first_wp_row = tlayout["recon_first_wp_row"]
    recon_total_row = tlayout["recon_total_row"]

    # -- Conditional formatting: Gap % (col E), data rows + TOTAL row --
    # Red added first (highest priority) so red wins at |gap| > 10%.
    recon_e_range = f"E{recon_first_wp_row}:E{recon_total_row}"
    recon_e_ref = f"E{recon_first_wp_row}"
    # Red: |gap %| > 10%
    ws.conditional_formatting.add(
        recon_e_range,
        FormulaRule(
            formula=[f"ABS({recon_e_ref})>0.10"],
            fill=_RED_CF_FILL,
        ),
    )
    # Amber: |gap %| > 5%
    ws.conditional_formatting.add(
        recon_e_range,
        FormulaRule(
            formula=[f"ABS({recon_e_ref})>0.05"],
            fill=_AMBER_CF_FILL,
        ),
    )


def _write_person_load_table(ws, tlayout, capacity_layout, team_layout, team):
    """Section 3: Person Load — section header, column headers, per-person
    data rows (cols A–D)."""
    person_section_row = tlayout["person_section_row"]
    person_header_row = tlayout["person_header_row"]
    person_first_row = tlayout["person_first_row"]

    # ================================================================
    # Section 3: Person Load Table
    # ================================================================
    # Push Person Load below the gap chart (chart anchor + 14 rows chart
    # + 1 trailing spacer = + 15).
    write_section_header(ws, person_section_row, "PERSON LOAD", merge_end=4)

    # Column headers
    person_headers = [
        "Person",
        "Avg FTE to date",
        "Plan Avg FTE to date",
        "Variance to-date %",
    ]
    for ci, h in enumerate(person_headers, 1):
        style_cell(ws, person_header_row, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Person data rows
    person_groups = capacity_layout["person_groups"]
    person_rows_map = team_layout["person_rows"]
    # Capacity TO-DATE band column letters: Actual PM + Variance PM columns
    # hold the per-person windowed sums already computed by the Capacity
    # builder. Plan PM to date = Actual + Variance (same identity as
    # Capacity's to-date band).
    todate_actual_cl = col_letter(capacity_layout["todate_col_actual"])
    todate_var_cl = col_letter(capacity_layout["todate_col_var"])

    for pi, member in enumerate(team):
        name = member["name"]
        row = person_first_row + pi
        tr = person_groups[name]["total_row"]
        roster_row = person_rows_map[name]

        # Capacity refs for this person's to-date Actual / Variance PM.
        cap_actual = f"{SHEET_CAPACITY}!{todate_actual_cl}{tr}"
        cap_var = f"{SHEET_CAPACITY}!{todate_var_cl}{tr}"

        # Col A: Person name (green cross-sheet ref)
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{roster_row}"

        # Col B: Avg FTE to date — Actual PM / MAX(1, hours_actuals_through_month).
        # IFERROR guards DIV/0 when hours_actuals_through_month <= 0.
        style_cell(
            ws, row, 2, font=GREEN_XREF, fmt="0.00"
        ).value = f"=IFERROR({cap_actual}/MAX(1,hours_actuals_through_month),0)"

        # Col C: Plan Avg FTE to date — (Actual + Variance) / MAX(1, hours_actuals_through_month).
        # Plan = Actual + Variance identity lives on Capacity; reuse here.
        style_cell(
            ws, row, 3, font=GREEN_XREF, fmt="0.00"
        ).value = f"=IFERROR(({cap_actual}+{cap_var})/MAX(1,hours_actuals_through_month),0)"

        # Col D: Variance to-date % — Variance / Plan. Guard against zero plan.
        style_cell(ws, row, 4, font=GREEN_XREF, fmt="0.0%").value = f"=IFERROR({cap_var}/({cap_actual}+{cap_var}),0)"


def _apply_person_load_cf(ws, tlayout):
    """CF on Person Load Variance to-date % (col D) — across all person rows.
    Red added BEFORE amber so red wins at |variance| >= 20%.
    """
    person_first_row = tlayout["person_first_row"]
    person_last_row = tlayout["person_last_row"]

    # -- Conditional formatting: Variance to-date % (col D) — across all person rows --
    # Red BEFORE amber so red wins at >=20% drift (openpyxl: first rule added
    # gets highest priority).
    d_range = f"D{person_first_row}:D{person_last_row}"
    d_ref = f"D{person_first_row}"

    # Red: |variance| > 20%
    ws.conditional_formatting.add(
        d_range,
        FormulaRule(
            formula=[f"ABS({d_ref})>0.20"],
            fill=_RED_CF_FILL,
        ),
    )
    # Amber: |variance| > 10% (red wins the >=20% band via priority order)
    ws.conditional_formatting.add(
        d_range,
        FormulaRule(
            formula=[f"ABS({d_ref})>0.10"],
            fill=_AMBER_CF_FILL,
        ),
    )


def _write_chart_helper_rows(ws, tlayout, layout, actuals_layout, project_params):
    """Hidden helper rows for the cumulative chart: Planned, Actual,
    Forecast, Ceiling. Hides each helper row.
    """
    # ================================================================
    # Section 3: Cumulative LineChart + hidden helper rows
    # ================================================================
    # Helper rows placed below person load section (2-row gap).
    # Order matches series stacking on the chart for readability:
    # Planned, Actual, Forecast, Ceiling, (labels), RP markers.
    num_wps = tlayout["num_wps"]
    month_col_start = tlayout["month_col_start"]
    helper_row_planned = tlayout["helper_row_planned"]
    helper_row_actual = tlayout["helper_row_actual"]
    helper_row_forecast = tlayout["helper_row_forecast"]
    helper_row_ceiling = tlayout["helper_row_ceiling"]
    posted_idx_col_letter = tlayout["posted_idx_col_letter"]
    posted_value_col_letter = tlayout["posted_value_col_letter"]
    slope_col_letter = tlayout["slope_col_letter"]

    plan_first_wp = layout["planned_first_wp_row"]
    plan_last_wp = layout["planned_last_wp_row"]
    act_total = actuals_layout["total_row"]
    act_first_wp = actuals_layout["first_wp_row"]
    act_last_wp = actuals_layout["last_wp_row"]

    # Project-aggregated 2-anchor forecast (Phase 2.4 / design §9): two
    # cells of the forecast helper row hold real values; everything else is
    # NA() so LineChart draws a single straight segment between them.
    #
    #   now-anchor    @ eur_actuals_through_month  = SUM of WP posted_value
    #   future-anchor @ project_months             = SUM of MAX(0, posted_value
    #                                                  + slope * (project_months
    #                                                  - posted_idx)) per WP
    #
    # Per-WP MAX-clamp wrapped in IFERROR(...,0) so the empty-string sentinel
    # from posted_idx=0 (no postings yet) doesn't propagate #VALUE through
    # the SUM. Individual `+` terms (rather than SUMPRODUCT array form) per
    # the openpyxl rule on text*number propagation.
    sum_posted_terms = []
    sum_forecast_terms = []
    for wp_off in range(num_wps):
        ar = act_first_wp + wp_off
        sum_posted_terms.append(f"IFERROR({SHEET_ACTUALS}!{posted_value_col_letter}{ar}+0,0)")
        sum_forecast_terms.append(
            f"IFERROR(MAX(0,{SHEET_ACTUALS}!{posted_value_col_letter}{ar}"
            f"+{SHEET_ACTUALS}!{slope_col_letter}{ar}"
            f"*(project_months-{SHEET_ACTUALS}!{posted_idx_col_letter}{ar})),0)"
        )
    sum_posted_now_expr = "+".join(sum_posted_terms) if sum_posted_terms else "0"
    sum_forecast_end_expr = "+".join(sum_forecast_terms) if sum_forecast_terms else "0"
    # Silence "unused" hints — they're consumed inside the loop body via the
    # per-WP terms above; act_last_wp kept for symmetry with summary.py
    # debugging guards even though the per-WP loop bounds are sufficient.
    _ = act_last_wp

    # Per-WP rate range from Config column D (for SUMPRODUCT)
    cfg_wp_first = layout["config_first_wp_row"]
    rate_range = f"{SHEET_CONFIG}!$D${cfg_wp_first}:$D${cfg_wp_first + num_wps - 1}"

    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        cl = col_letter(c)

        # Helper Row 1: Cumulative planned EUR (SUMPRODUCT with per-WP rates)
        plan_range = f"{SHEET_PLAN}!{cl}{plan_first_wp}:{cl}{plan_last_wp}"
        if m_idx == 0:
            style_cell(
                ws, helper_row_planned, c, font=BLACK_FORMULA, fmt="\u20ac#,##0"
            ).value = f"=SUMPRODUCT({plan_range},{rate_range})"
        else:
            prev_cl = col_letter(c - 1)
            style_cell(
                ws, helper_row_planned, c, font=BLACK_FORMULA, fmt="\u20ac#,##0"
            ).value = f"={prev_cl}{helper_row_planned}+SUMPRODUCT({plan_range},{rate_range})"

        # Helper Row 2: Cumulative actual EUR — clipped past
        # `eur_actuals_through_month` so the chart line ends cleanly at the
        # through-month edge (NA() cells are not plotted by LineChart).
        # Month index in the formula is the literal integer m_idx + 1
        # (compile-time constant).
        month_idx_literal = m_idx + 1
        if m_idx == 0:
            style_cell(
                ws, helper_row_actual, c, font=BLACK_FORMULA, fmt="\u20ac#,##0"
            ).value = f"=IF({month_idx_literal}>eur_actuals_through_month,NA(),{SHEET_ACTUALS}!{cl}{act_total})"
        else:
            prev_cl = col_letter(c - 1)
            style_cell(ws, helper_row_actual, c, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = (
                f"=IF({month_idx_literal}>eur_actuals_through_month,NA(),"
                f"{prev_cl}{helper_row_actual}+{SHEET_ACTUALS}!{cl}{act_total})"
            )

        # Helper Row 3: Cumulative forecast EUR — 2-anchor pattern. The
        # now-anchor (m=eur_actuals_through_month) emits the current
        # posted-cumulative total; the project-end anchor
        # (m=project_months) emits the slope-projected end total. All
        # off-anchor cells emit NA() so the chart series draws a single
        # straight segment between the two anchors.
        forecast_formula = (
            f"=IF({month_idx_literal}=eur_actuals_through_month,"
            f"{sum_posted_now_expr},"
            f"IF({month_idx_literal}=project_months,"
            f"{sum_forecast_end_expr},"
            f"NA()))"
        )
        style_cell(ws, helper_row_forecast, c, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = forecast_formula

        # Helper Row 4: Ceiling line (flat)
        style_cell(ws, helper_row_ceiling, c, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = "=partner_ceiling"

    # Hide helper rows
    ws.row_dimensions[helper_row_planned].hidden = True
    ws.row_dimensions[helper_row_actual].hidden = True
    ws.row_dimensions[helper_row_forecast].hidden = True
    ws.row_dimensions[helper_row_ceiling].hidden = True


def _write_chart_label_and_rp_rows(ws, tlayout, project_params, reporting_periods):
    """Month-label row (X-axis categories) and RP boundary marker row for
    the cumulative chart. Both rows hidden.
    """
    month_col_start = tlayout["month_col_start"]
    month_col_end = tlayout["month_col_end"]
    helper_row_labels = tlayout["helper_row_labels"]
    helper_row_rp = tlayout["helper_row_rp"]

    # X-axis category labels (month-year, e.g. "Nov 25")
    project_start = project_params["start"]
    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        cal_date = project_start + relativedelta(months=m_idx)
        ws.cell(row=helper_row_labels, column=c, value=f"M{m_idx + 1}-{cal_date.strftime('%b%y')}")
    ws.row_dimensions[helper_row_labels].hidden = True

    # Helper Row 5: RP boundary markers (ceiling value at each RP end month)
    for rp in reporting_periods:
        boundary_col = month_col_start + rp["end"] - 1
        if boundary_col <= month_col_end:
            style_cell(
                ws, helper_row_rp, boundary_col, font=BLACK_FORMULA, fmt="\u20ac#,##0"
            ).value = "=partner_ceiling"
    ws.row_dimensions[helper_row_rp].hidden = True


def _write_gap_helper_rows(ws, tlayout, project_params):
    """Hidden helper rows for the gap-over-time chart: gap (actual-forecast)
    and zero reference line.
    """
    month_col_start = tlayout["month_col_start"]
    helper_row_actual = tlayout["helper_row_actual"]
    helper_row_forecast = tlayout["helper_row_forecast"]
    helper_row_gap = tlayout["helper_row_gap"]
    helper_row_zero = tlayout["helper_row_zero"]

    # Helper Rows 7 + 8: gap-over-time + zero reference line.
    # Gap = cumulative actual - cumulative forecast per month. Because the
    # actual and forecast helper rows already emit NA() past their
    # respective through-months, subtraction naturally propagates NA()
    # past reconciled_through_month (= MIN of the two).
    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        cl = col_letter(c)
        # Gap helper: actual - forecast. Black same-sheet derived.
        style_cell(
            ws, helper_row_gap, c, font=BLACK_FORMULA, fmt="\u20ac#,##0"
        ).value = f"={cl}{helper_row_actual}-{cl}{helper_row_forecast}"
        # Zero reference: formula =0 (Rule 1: formulas only).
        style_cell(ws, helper_row_zero, c, font=BLACK_FORMULA, fmt="\u20ac#,##0").value = "=0"
    ws.row_dimensions[helper_row_gap].hidden = True
    ws.row_dimensions[helper_row_zero].hidden = True


def _build_cumulative_chart(ws, tlayout):
    """Build the cumulative-EUR LineChart (Planned / Actual / Forecast /
    Ceiling / RP End) and anchor it at ``chart_anchor_row``.

    Series are constructed manually (NumDataSource/NumRef/AxDataSource)
    rather than via ``chart.add_data`` because LibreOffice's recalc macro
    strips val/tx/cat from add_data-created series.
    """
    ms = tlayout["ms"]
    me = tlayout["me"]
    chart_anchor_row = tlayout["chart_anchor_row"]
    helper_row_planned = tlayout["helper_row_planned"]
    helper_row_actual = tlayout["helper_row_actual"]
    helper_row_forecast = tlayout["helper_row_forecast"]
    helper_row_ceiling = tlayout["helper_row_ceiling"]
    helper_row_labels = tlayout["helper_row_labels"]
    helper_row_rp = tlayout["helper_row_rp"]

    # -- Build the LineChart --
    # NOTE: Series are constructed manually (not via chart.add_data) because
    # LibreOffice's recalc macro strips val/tx/cat from add_data-created series.
    chart = LineChart()
    chart.title = "Cumulative EUR: Planned vs Actual"
    chart.style = 10
    chart.visible_cells_only = False  # helper rows are hidden
    chart.y_axis.numFmt = "\u20ac#,##0"
    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.delete = False

    sheet_name = "Trends"
    range_str = f"'{sheet_name}'!${ms}${{}}:${me}${{}}"

    cat_ref = AxDataSource(strRef=StrRef(f=range_str.format(helper_row_labels, helper_row_labels)))

    # Series 1: Cumulative planned (blue solid)
    s_planned = Series(idx=0, order=0)
    s_planned.tx = SeriesLabel(v="Planned")
    s_planned.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_planned, helper_row_planned)))
    s_planned.cat = cat_ref
    s_planned.graphicalProperties.line.solidFill = "4472C4"
    chart.series.append(s_planned)

    # Series 2: Cumulative actual (red solid)
    s_actual = Series(idx=1, order=1)
    s_actual.tx = SeriesLabel(v="Actual")
    s_actual.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_actual, helper_row_actual)))
    s_actual.cat = cat_ref
    s_actual.graphicalProperties.line.solidFill = "FF0000"
    chart.series.append(s_actual)

    # Series 3: Cumulative forecast (amber dashed) — project-aggregated
    # 2-anchor series fed by `helper_row_forecast` (post Phase 2.4: NA() in
    # all months except eur_actuals_through_month and project_months, so
    # LineChart draws a single straight segment between the two anchors).
    # Inserted between Actual and Ceiling so legend ordering matches
    # semantic priority.
    s_forecast = Series(idx=2, order=2)
    s_forecast.tx = SeriesLabel(v="Cumulative Forecast")
    s_forecast.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_forecast, helper_row_forecast)))
    s_forecast.cat = cat_ref
    s_forecast.graphicalProperties.line.solidFill = "FFB300"
    s_forecast.graphicalProperties.line.dashStyle = "dash"
    chart.series.append(s_forecast)

    # Series 4: Ceiling (grey dashed)
    s_ceiling = Series(idx=3, order=3)
    s_ceiling.tx = SeriesLabel(v="Ceiling")
    s_ceiling.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_ceiling, helper_row_ceiling)))
    s_ceiling.cat = cat_ref
    s_ceiling.graphicalProperties.line.solidFill = "808080"
    s_ceiling.graphicalProperties.line.dashStyle = "dash"
    chart.series.append(s_ceiling)

    # Series 5: RP boundaries (green diamond markers at ceiling level)
    s_rp = Series(idx=4, order=4)
    s_rp.tx = SeriesLabel(v="RP End")
    s_rp.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_rp, helper_row_rp)))
    s_rp.cat = cat_ref
    s_rp.graphicalProperties.line.solidFill = "00B050"
    s_rp.graphicalProperties.line.dashStyle = "dot"
    rp_marker = Marker(symbol="diamond", size=10)
    assert rp_marker.graphicalProperties is not None
    rp_marker.graphicalProperties.solidFill = "00B050"
    s_rp.marker = rp_marker
    chart.series.append(s_rp)

    # Chart dimensions and placement
    chart.width = 30
    chart.height = 7.5
    ws.add_chart(chart, f"A{chart_anchor_row}")


def _build_gap_chart(ws, tlayout):
    """Build the gap-over-time LineChart (cumulative Actual EUR − cumulative
    Forecast EUR) and anchor it at ``recon_gap_chart_anchor_row``.

    Single-series line showing drift of actual vs forecast. Clips via NA()
    past reconciled_through_month (inherited from the actual / forecast
    helper rows). A second dashed-grey zero reference line makes sign of
    drift visually obvious.
    """
    ms = tlayout["ms"]
    me = tlayout["me"]
    recon_gap_chart_anchor_row = tlayout["recon_gap_chart_anchor_row"]
    helper_row_labels = tlayout["helper_row_labels"]
    helper_row_gap = tlayout["helper_row_gap"]
    helper_row_zero = tlayout["helper_row_zero"]

    # ================================================================
    # Gap-over-time chart (cumulative Actual EUR − cumulative Forecast EUR)
    # ================================================================
    # Single-series line showing drift of actual vs forecast. Clips via
    # NA() past reconciled_through_month (inherited from the actual /
    # forecast helper rows). A second dashed-grey zero reference line
    # makes sign of drift visually obvious.
    gap_chart = LineChart()
    gap_chart.title = "Cumulative EUR Gap (Actual \u2212 Forecast)"
    gap_chart.style = 10
    gap_chart.visible_cells_only = False  # helper rows are hidden
    gap_chart.y_axis.numFmt = "\u20ac#,##0"
    gap_chart.y_axis.tickLblPos = "nextTo"
    gap_chart.y_axis.delete = False
    gap_chart.x_axis.axPos = "b"
    gap_chart.x_axis.tickLblPos = "nextTo"
    gap_chart.x_axis.delete = False

    sheet_name = "Trends"
    range_str = f"'{sheet_name}'!${ms}${{}}:${me}${{}}"

    # Reuse the month-label helper row from the cumulative chart.
    gap_cat_ref = AxDataSource(strRef=StrRef(f=range_str.format(helper_row_labels, helper_row_labels)))

    # Series 1: Gap (solid purple)
    s_gap = Series(idx=0, order=0)
    s_gap.tx = SeriesLabel(v="Gap (Actual \u2212 Forecast)")
    s_gap.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_gap, helper_row_gap)))
    s_gap.cat = gap_cat_ref
    s_gap.graphicalProperties.line.solidFill = "7030A0"
    gap_chart.series.append(s_gap)

    # Series 2: Zero reference (grey dashed)
    s_zero = Series(idx=1, order=1)
    s_zero.tx = SeriesLabel(v="Zero")
    s_zero.val = NumDataSource(numRef=NumRef(f=range_str.format(helper_row_zero, helper_row_zero)))
    s_zero.cat = gap_cat_ref
    s_zero.graphicalProperties.line.solidFill = "808080"
    s_zero.graphicalProperties.line.dashStyle = "dash"
    gap_chart.series.append(s_zero)

    gap_chart.width = 30
    gap_chart.height = 7.5
    ws.add_chart(gap_chart, f"A{recon_gap_chart_anchor_row}")


def _apply_column_widths(ws):
    """Auto-fit columns then pin A/B widths for label readability."""
    # ================================================================
    # Column widths
    # ================================================================
    auto_fit_columns(ws)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 25


def build_trends_sheet(
    ws, layout, actuals_layout, capacity_layout, team_layout, *, project_params, work_packages, team, reporting_periods
):
    """Build the Trends sheet: WP burn rate, cumulative chart, person load.

    Args:
        ws: The Trends worksheet (already has current_month in B1).
        layout: Plan layout dict (config_first_wp_row, etc).
        actuals_layout: Actuals layout dict (first_wp_row, posted_idx_col,
            posted_value_col, slope_col, etc).
        capacity_layout: Capacity layout dict (person_groups with total_row per person).
        team_layout: Team layout dict (person_rows mapping name->row).
        project_params: Dict with keys title, start, months, rate.
        work_packages: List of WP dicts with single-range 'start'/'end'/'type'.
        team: List of team member dicts.
        reporting_periods: Reporting period entries.
    """
    tlayout = _resolve_trends_layout(
        layout,
        actuals_layout,
        project_params,
        work_packages,
        team,
    )

    # Section 1: WP Burn Rate & Exhaustion (table + CF + TOTAL row)
    _write_burn_rate_table(ws, tlayout, layout, actuals_layout, work_packages)
    _apply_burn_rate_cf(ws, tlayout, work_packages)
    _write_burn_rate_total(ws, tlayout, actuals_layout)

    # Section 2a: Reconciliation block (header + health rows + per-WP table + CF)
    _write_reconciliation_section(ws, tlayout, layout, actuals_layout)
    _apply_reconciliation_cf(ws, tlayout)

    # Section 3: Person Load (table + CF)
    _write_person_load_table(ws, tlayout, capacity_layout, team_layout, team)
    _apply_person_load_cf(ws, tlayout)

    # Hidden helper rows feeding both charts.
    _write_chart_helper_rows(ws, tlayout, layout, actuals_layout, project_params)
    _write_chart_label_and_rp_rows(ws, tlayout, project_params, reporting_periods)
    _write_gap_helper_rows(ws, tlayout, project_params)

    # Charts (order matters: cumulative first, then gap — preserves
    # ws._charts[0]/[1] index assertions in tests).
    _build_cumulative_chart(ws, tlayout)
    _build_gap_chart(ws, tlayout)

    _apply_column_widths(ws)

    # ================================================================
    # Return layout dict
    # ================================================================
    return {
        "section_header_row": tlayout["section1_row"],
        "burn_rate_header_row": tlayout["header_row"],
        "burn_rate_first_wp_row": tlayout["first_wp_data_row"],
        "burn_rate_last_wp_row": tlayout["last_wp_data_row"],
        "burn_rate_total_row": tlayout["total_row"],
        "chart_anchor_row": tlayout["chart_anchor_row"],
        "recon_section_row": tlayout["recon_section_row"],
        "recon_health_eur_through_row": tlayout["recon_health_eur_through_row"],
        "recon_health_hours_through_row": tlayout["recon_health_hours_through_row"],
        "recon_health_reconciled_row": tlayout["recon_health_reconciled_row"],
        "recon_table_header_row": tlayout["recon_table_header_row"],
        "recon_first_wp_row": tlayout["recon_first_wp_row"],
        "recon_last_wp_row": tlayout["recon_last_wp_row"],
        "recon_total_row": tlayout["recon_total_row"],
        "person_load_section_row": tlayout["person_section_row"],
        "person_load_header_row": tlayout["person_header_row"],
        "person_load_first_row": tlayout["person_first_row"],
        "person_load_last_row": tlayout["person_last_row"],
        "helper_row_planned": tlayout["helper_row_planned"],
        "helper_row_actual": tlayout["helper_row_actual"],
        "helper_row_forecast": tlayout["helper_row_forecast"],
        "helper_row_ceiling": tlayout["helper_row_ceiling"],
        "helper_row_labels": tlayout["helper_row_labels"],
        "helper_row_rp": tlayout["helper_row_rp"],
        "helper_row_gap": tlayout["helper_row_gap"],
        "helper_row_zero": tlayout["helper_row_zero"],
        "recon_gap_chart_anchor_row": tlayout["recon_gap_chart_anchor_row"],
    }
