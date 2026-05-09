"""Actuals sheet builder — posted-cumulative EUR per WP per month from accountant.

Role: monthly cells are an input surface for posted-cumulative EUR
(accountant figures — running total posted through that month). Cross-
sheet trust boundary (``eur_actuals_through_month`` named range at
``Actuals!$D$1``) and the through-month / two-cadence rules are
canonical in ``docs/spec.md`` § Two-cadence actuals. Posted-cumulative
redesign rationale:
``docs/plans/2026-05-09-actuals-posted-cumulative-design.md``.

## Per-WP columns (in display order)

-i side (collapsible band, hidden by default):
    monthly grid (cols C..month_col_end) — BLUE_INPUT, posted-cumulative
    Slope -i        — IF(posted_idx=0, "", posted_value/posted_idx)
    Budget -i       — cross-sheet ref to Config WP Budget -i
    Variance -i     — IF(posted_idx=0, "", Budget - Posted)
    Posted -i       — IFERROR(INDEX(grid, posted_idx), "")
    % Budget -i     — IF(posted_idx=0, "", Posted / Budget)
    Breach-In -i    — IF(posted_idx=0, "", IF(Posted=0, "—",
                          (Budget - Posted) / (Posted / MAX(1,
                          eur_actuals_through_month))))

+i mirror block (default-visible):
    Slope +i / Budget +i / Variance +i / Posted +i / % Budget +i /
    Breach-In +i — same shape; numeric leaves derive via
    ROUND(<-i> × 1.25, 0). Variance/%/Breach guard via the SAME -i
    posted_idx helper (single source of truth for "have we posted yet").

Hidden helpers (-i side, outline-grouped, BLACK_FORMULA):
    posted_idx     — SUMPRODUCT-MAX scan returning latest 1-based month
                     index with a numeric posting (banned-pattern-safe).
    posted_value   — IFERROR(INDEX(grid, posted_idx), "").

## Conditional formatting

- Negative variance (-i / +i): **red**.
- % of Budget (-i / +i mirror rule — scale-invariant): **amber** > 85%,
  **red** > 100%.
- Breach In (-i / +i): **amber** < 12 months, **red** < 6 months.
- **Reconciliation grey (T8):** Actual-EUR monthly cells past
  ``eur_actuals_through_month`` are greyed (not yet finance-closed).
"""

from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.layout import (
    MONTH_COL_START, SHEET_ACTUALS, SHEET_CONFIG,
    COL_WP_BUDGET_MINUS_I,
    ACTUALS_AMBER_PCT_CEIL, ACTUALS_AMBER_BREACH_MONTHS, ACTUALS_RED_BREACH_MONTHS,
    build_year_month_headers, cell_ref, month_range_ref,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
)
from budget_planner.styles import (
    BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, HEADER_FONT, SECTION_FONT,
    GREY_FILL, LIGHT_BLUE_FILL, LIGHT_GREEN_FILL,
    SECTION_FILL,
    WITH_INDIRECT_HEADER_FILL, WITHOUT_INDIRECT_HEADER_FILL,
    CF_RED_FILL, CF_AMBER_FILL, CF_GREY_FILL,
    CENTER,
    sized_comment, style_cell, col_letter,
)
from budget_planner import config


def _resolve_actuals_layout(project_params, work_packages, layout):
    """Resolve all column indices, row anchors, and band membership for the
    Actuals sheet into a single dict consumed by the helpers below.

    Pure function — no worksheet writes. Keeps the orchestrator readable
    by lifting the dense column-index block out of ``build_actuals_sheet``.
    """
    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    # -i rollups (live inside collapsible band — hidden by default).
    val_col_slope = _vcols["slope"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    val_col_posted = _vcols["posted"]
    val_col_pct_ceil = _vcols["pct_ceil"]
    val_col_breach = _vcols["breach"]
    # +i mirrors (default-visible outside outline band).
    val_col_slope_pi = _vcols["slope_plus_i"]
    val_col_budget_pi = _vcols["budget_plus_i"]
    val_col_diff_pi = _vcols["diff_plus_i"]
    val_col_posted_pi = _vcols["posted_plus_i"]
    val_col_pct_ceil_pi = _vcols["pct_ceil_plus_i"]
    val_col_breach_pi = _vcols["breach_plus_i"]
    # Hidden helpers (-i side; outline-grouped, BLACK_FORMULA). +i side
    # does NOT need helper mirrors — Posted/Slope +i = ROUND(-i × 1.25, 0)
    # reading the -i visible cells directly, and Variance/%/Breach +i
    # guards reuse the same -i posted_idx (single source of truth).
    val_col_posted_idx = _vcols["posted_idx"]
    val_col_posted_value = _vcols["posted_value"]

    # Columns that belong inside the collapsible -i outline band: the
    # monthly grid (cols 3..month_col_end), the -i visible rollups
    # (offsets 1..6), and the hidden helpers (offsets 13..14). +i mirrors
    # at offsets 7..12 stay outside (default-visible).
    _minus_i_band_cols = (
        *range(month_col_start, month_col_end + 1),
        val_col_slope, val_col_budget, val_col_diff,
        val_col_posted, val_col_pct_ceil, val_col_breach,
        val_col_posted_idx, val_col_posted_value,
    )

    config_first_wp_row = layout["config_first_wp_row"]

    return {
        "month_col_start": month_col_start,
        "month_col_end": month_col_end,
        "num_wps": num_wps,
        "val_col_slope": val_col_slope,
        "val_col_budget": val_col_budget,
        "val_col_diff": val_col_diff,
        "val_col_posted": val_col_posted,
        "val_col_pct_ceil": val_col_pct_ceil,
        "val_col_breach": val_col_breach,
        "val_col_slope_pi": val_col_slope_pi,
        "val_col_budget_pi": val_col_budget_pi,
        "val_col_diff_pi": val_col_diff_pi,
        "val_col_posted_pi": val_col_posted_pi,
        "val_col_pct_ceil_pi": val_col_pct_ceil_pi,
        "val_col_breach_pi": val_col_breach_pi,
        "val_col_posted_idx": val_col_posted_idx,
        "val_col_posted_value": val_col_posted_value,
        "minus_i_band_cols": _minus_i_band_cols,
        "config_first_wp_row": config_first_wp_row,
    }


def _write_asof_header_row(ws, r, alayout):
    """Write the "Posted through:" header row (row 2 — shifted down by the
    row 1 through-month strip).

    Single label/date pair: B/C cells above the per-WP block. The
    through-month cell holds an integer month index;
    EDATE(project_start, N-1) turns it into a calendar date so the
    header reads "Posted through: Apr 2026" not "1900-01". The legacy
    duplicate above the forecast monitoring block was dropped in the
    posted-cumulative redesign (Phase 2.2 — that block no longer exists).
    """
    del alayout  # signature kept consistent with peer helpers
    _ASOF_DATE_FMT = "MMM YYYY"
    style_cell(ws, r, 2, value="Posted through:", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, border=False)
    style_cell(ws, r, 3,
               value="=EDATE(project_start, eur_actuals_through_month - 1)",
               font=GREEN_XREF, fill=LIGHT_BLUE_FILL,
               fmt=_ASOF_DATE_FMT, border=False)


def _write_section_headers(ws, r, alayout):
    """Write the merged ACTUAL EUR section banners.

    Splits into two merged regions: "-i" over the collapsible rollup
    block (cols 1..breach) and "+i" over the default-visible mirror
    block (slope_pi..breach_pi). Forecast-monitoring banners were
    dropped in the posted-cumulative redesign — per-WP slope/forecast
    are inline columns, not a separate band.
    """
    val_col_breach = alayout["val_col_breach"]
    val_col_slope_pi = alayout["val_col_slope_pi"]
    val_col_breach_pi = alayout["val_col_breach_pi"]

    def _section_header(first_col, last_col, label):
        ws.merge_cells(start_row=r, start_column=first_col,
                       end_row=r, end_column=last_col)
        style_cell(ws, r, first_col, label, font=SECTION_FONT,
                   fill=SECTION_FILL, alignment=CENTER)
        for c2 in range(first_col + 1, last_col + 1):
            style_cell(ws, r, c2, fill=SECTION_FILL)

    _section_header(1, val_col_breach, "ACTUAL EUR -i")
    _section_header(val_col_slope_pi, val_col_breach_pi, "ACTUAL EUR +i")


def _apply_header_fill_overrides(ws, year_row, month_row, alayout):
    """Override header fills for the -i / +i blocks (default fill from
    build_year_month_headers is LIGHT_BLUE_FILL — replace with the
    T1 indirect-view fills so +i (green) and -i (orange) blocks are
    visually distinguishable at the month-header row).
    """
    _minus_i_header_cols = (
        alayout["val_col_slope"], alayout["val_col_budget"], alayout["val_col_diff"],
        alayout["val_col_posted"], alayout["val_col_pct_ceil"], alayout["val_col_breach"],
    )
    _plus_i_header_cols = (
        alayout["val_col_slope_pi"], alayout["val_col_budget_pi"], alayout["val_col_diff_pi"],
        alayout["val_col_posted_pi"], alayout["val_col_pct_ceil_pi"], alayout["val_col_breach_pi"],
    )
    for c in _minus_i_header_cols:
        ws.cell(year_row, c).fill = WITHOUT_INDIRECT_HEADER_FILL
        ws.cell(month_row, c).fill = WITHOUT_INDIRECT_HEADER_FILL
    for c in _plus_i_header_cols:
        ws.cell(year_row, c).fill = WITH_INDIRECT_HEADER_FILL
        ws.cell(month_row, c).fill = WITH_INDIRECT_HEADER_FILL


def _append_minus_i_year_anchor_suffix(ws, year_row, month_col_start):
    """Append "-i" suffix to the first calendar-year cell.

    Month-section label (-i suffix): the year-row cells above the
    monthly grid carry the calendar-year labels; append "-i" to the
    first year-row cell so the section reads as without-indirect per
    D8. Uses the first calendar-year cell (month_col_start) as the
    anchor — subsequent year cells are unsuffixed.
    """
    _year_anchor = ws.cell(year_row, month_col_start)
    if _year_anchor.value:
        _year_anchor.value = f"{_year_anchor.value} -i"
    else:
        _year_anchor.value = "-i"


def _write_wp_rows(ws, first_wp_row, alayout, work_packages, preserved,
                   project_params):
    """Write one row per WP — blue-input posted-cumulative EUR cells +
    -i visible rollups + +i mirrors + hidden helpers (posted_idx /
    posted_value).

    Returns the last WP row index (``last_wp_row``).

    Posted-cumulative semantics: each monthly cell is the running EUR
    posted through that month (blank = no posting recorded). Visible
    Posted -i / Slope -i / Variance / % / Breach -i columns are derived
    from the two hidden helper cells (posted_idx, posted_value); see
    ``docs/plans/2026-05-09-actuals-posted-cumulative-design.md`` § 2.
    """
    month_col_start = alayout["month_col_start"]
    month_col_end = alayout["month_col_end"]
    config_first_wp_row = alayout["config_first_wp_row"]
    val_col_slope = alayout["val_col_slope"]
    val_col_budget = alayout["val_col_budget"]
    val_col_diff = alayout["val_col_diff"]
    val_col_posted = alayout["val_col_posted"]
    val_col_pct_ceil = alayout["val_col_pct_ceil"]
    val_col_breach = alayout["val_col_breach"]
    val_col_slope_pi = alayout["val_col_slope_pi"]
    val_col_budget_pi = alayout["val_col_budget_pi"]
    val_col_diff_pi = alayout["val_col_diff_pi"]
    val_col_posted_pi = alayout["val_col_posted_pi"]
    val_col_pct_ceil_pi = alayout["val_col_pct_ceil_pi"]
    val_col_breach_pi = alayout["val_col_breach_pi"]
    val_col_posted_idx = alayout["val_col_posted_idx"]
    val_col_posted_value = alayout["val_col_posted_value"]

    ms = col_letter(month_col_start)

    r = first_wp_row
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = config_first_wp_row + i
        # Cross-sheet refs to Config WP table
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"={SHEET_CONFIG}!A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"={SHEET_CONFIG}!B{config_row}"

        # Get preserved monthly values if available
        preserved_values = None
        if preserved and wp["id"] in preserved:
            preserved_values = preserved[wp["id"]]

        active = wp_active_months(wp["id"], work_packages)
        for m_idx in range(project_params["months"]):
            c = month_col_start + m_idx
            month_num = m_idx + 1
            if month_num in active:
                if preserved_values is not None and m_idx < len(preserved_values) and preserved_values[m_idx] is not None:
                    eur_value: int | float = preserved_values[m_idx]
                    style_cell(ws, row, c, eur_value, font=BLUE_INPUT, fmt="€#,##0")
                else:
                    style_cell(ws, row, c, font=BLUE_INPUT, fmt="€#,##0")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL,
                           fmt="€#,##0")

        # ── Hidden helpers (outline-grouped, BLACK_FORMULA). Written
        #     BEFORE the visible -i columns so the visible-column
        #     formulas can reference these cells by row coord. ──
        # posted_idx: latest 1-based month index with a numeric posting
        # at-or-before current_month. SUMPRODUCT-MAX-array form is the
        # banned-pattern-safe replacement for MATCH/LOOKUP under
        # openpyxl + LO recalc (see openpyxl.md § Core patterns).
        wp_range = month_range_ref(row, month_col_start, month_col_end)
        idx_l = col_letter(val_col_posted_idx)
        val_l = col_letter(val_col_posted_value)
        style_cell(ws, row, val_col_posted_idx, font=BLACK_FORMULA,
                   fmt="0").value = (
            f"=SUMPRODUCT(MAX("
            f"(COLUMN({wp_range})-COLUMN(${ms}$1)+1<=current_month)"
            f"*(COLUMN({wp_range})-COLUMN(${ms}$1)+1)"
            f"*ISNUMBER({wp_range})"
            f"))"
        )
        # posted_value: INDEX into the WP row at posted_idx; "" when
        # posted_idx=0 (no postings yet) so downstream consumers can
        # treat the empty-string sentinel as "no data."
        style_cell(ws, row, val_col_posted_value, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f"=IFERROR(INDEX({wp_range},{idx_l}{row}),\"\")"
        )

        # ── -i visible rollup block (inside collapsible band) ──
        # Slope -i: avg-burn = posted_value / posted_idx (per design §3).
        # Guard against /0 when no postings; "" sentinel surfaces as
        # blank in display.
        style_cell(ws, row, val_col_slope, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f'=IF({idx_l}{row}=0,"",{val_l}{row}/{idx_l}{row})'
        )
        # Budget -i (from Config WP Budget -i column — the black-formula
        # cell that derives = Budget +i / 1.25, so the numeric value here
        # matches the legacy Ceiling value)
        style_cell(ws, row, val_col_budget, font=GREEN_XREF,
                   fmt="€#,##0").value = (
            f"={cell_ref(SHEET_CONFIG, config_row, COL_WP_BUDGET_MINUS_I, absolute=False)}"
        )
        # Posted -i: latest filled cumulative through current_month.
        # Identical formula to the posted_value helper — cell-level
        # redundancy is intentional (user clicks the cell, sees the
        # calc; the helper is hidden).
        style_cell(ws, row, val_col_posted, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f"=IFERROR(INDEX({wp_range},{idx_l}{row}),\"\")"
        )
        budget_l = col_letter(val_col_budget)
        posted_l = col_letter(val_col_posted)
        # Variance -i = Budget -i - Posted -i. Guarded by posted_idx so
        # "" (no postings) surfaces as blank, not a misleading
        # Budget-minus-empty.
        style_cell(ws, row, val_col_diff, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f'=IF({idx_l}{row}=0,"",{budget_l}{row}-{posted_l}{row})'
        )
        # % of Budget -i = Posted -i / Budget -i. Same blank-sentinel
        # guard as Variance.
        style_cell(ws, row, val_col_pct_ceil, font=BLACK_FORMULA,
                   fmt="0.0%").value = (
            f'=IF({idx_l}{row}=0,"",{posted_l}{row}/{budget_l}{row})'
        )
        # Breach-In -i = months until Budget -i exhausted at posted-to-
        # date rate. Denominator inlines MAX(1,
        # eur_actuals_through_month) so the slope reflects spend per
        # closed finance-month, not the raw project-month index.
        # Outer guard: posted_idx=0 → ""; inner guard preserves the
        # legacy Posted=0 → "—" branch (genuine zero spend, distinct
        # from "no postings yet").
        style_cell(ws, row, val_col_breach, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({idx_l}{row}=0,"",'
            f'IF({posted_l}{row}=0,"—",'
            f"({budget_l}{row}-{posted_l}{row})"
            f"/({posted_l}{row}/MAX(1,eur_actuals_through_month))))"
        )

        # ── +i mirror block (default-visible) ──
        # Slope +i = ROUND(Slope -i * 1.25, 0).
        slope_l = col_letter(val_col_slope)
        style_cell(ws, row, val_col_slope_pi, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f'=IF({idx_l}{row}=0,"",ROUND({slope_l}{row}*1.25,0))'
        )
        # Budget +i = INDEX(wp_budget_with_indirect, i+1). This targets
        # the Config Budget +i column directly and is already
        # with-indirect, so NO ×1.25 applied.
        style_cell(ws, row, val_col_budget_pi, font=GREEN_XREF,
                   fmt="€#,##0").value = (
            f"=INDEX(wp_budget_with_indirect,{i + 1})"
        )
        # Posted +i = ROUND(Posted -i * 1.25, 0). Same blank-sentinel
        # guard as Slope.
        style_cell(ws, row, val_col_posted_pi, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f'=IF({idx_l}{row}=0,"",ROUND({posted_l}{row}*1.25,0))'
        )
        budget_pi_l = col_letter(val_col_budget_pi)
        posted_pi_l = col_letter(val_col_posted_pi)
        # Variance +i = Budget +i - Posted +i (guarded).
        style_cell(ws, row, val_col_diff_pi, font=BLACK_FORMULA,
                   fmt="€#,##0").value = (
            f'=IF({idx_l}{row}=0,"",{budget_pi_l}{row}-{posted_pi_l}{row})'
        )
        # % of Budget +i = Posted +i / Budget +i (guarded).
        style_cell(ws, row, val_col_pct_ceil_pi, font=BLACK_FORMULA,
                   fmt="0.0%").value = (
            f'=IF({idx_l}{row}=0,"",{posted_pi_l}{row}/{budget_pi_l}{row})'
        )
        # Breach-In +i — same shape as -i, against the +i cells.
        style_cell(ws, row, val_col_breach_pi, font=BLACK_FORMULA,
                   fmt="0").value = (
            f'=IF({idx_l}{row}=0,"",'
            f'IF({posted_pi_l}{row}=0,"—",'
            f"({budget_pi_l}{row}-{posted_pi_l}{row})"
            f"/({posted_pi_l}{row}/MAX(1,eur_actuals_through_month))))"
        )

    last_wp_row = first_wp_row + alayout["num_wps"] - 1
    return last_wp_row


def _write_total_row_monthly_grid(ws, total_row, first_wp_row, last_wp_row,
                                  alayout, project_params):
    """Write the TOTAL row label + the monthly-grid SUM cells."""
    month_col_start = alayout["month_col_start"]
    style_cell(ws, total_row, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, total_row, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for m_idx in range(project_params["months"]):
        c = month_col_start + m_idx
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, total_row, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="€#,##0").value = f"=SUM({col_range})"


def _write_total_row_minus_i(ws, total_row, first_wp_row, last_wp_row, alayout):
    """Write TOTAL row -i rollups (inside collapsible band).

    Posted/Slope are aggregated as ``=SUM`` over the per-WP -i columns —
    Excel/LO treat the IFERROR-empty-string sentinel ("") as 0 in SUM,
    so WPs with no postings yet contribute 0 cleanly. Variance / % /
    Breach derive from the aggregated TOTAL Posted (without a per-WP
    posted_idx — the trust boundary at sheet level is
    `eur_actuals_through_month`).
    """
    r = total_row
    val_col_slope = alayout["val_col_slope"]
    val_col_budget = alayout["val_col_budget"]
    val_col_diff = alayout["val_col_diff"]
    val_col_posted = alayout["val_col_posted"]
    val_col_pct_ceil = alayout["val_col_pct_ceil"]
    val_col_breach = alayout["val_col_breach"]

    # Slope -i = SUM over WP -i Slope column (avg-burn aggregated).
    slope_range = (f"{col_letter(val_col_slope)}{first_wp_row}:"
                   f"{col_letter(val_col_slope)}{last_wp_row}")
    style_cell(ws, r, val_col_slope, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = f"=SUM({slope_range})"
    # Budget -i = partner_ceiling (without-indirect).
    style_cell(ws, r, val_col_budget, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = "=partner_ceiling"
    # Posted -i = SUM over WP -i Posted column.
    posted_range = (f"{col_letter(val_col_posted)}{first_wp_row}:"
                    f"{col_letter(val_col_posted)}{last_wp_row}")
    style_cell(ws, r, val_col_posted, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = f"=SUM({posted_range})"
    # Variance -i = Budget -i - Posted -i
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = (
        f"={col_letter(val_col_budget)}{r}"
        f"-{col_letter(val_col_posted)}{r}"
    )
    # % of Budget -i = Posted -i / partner_ceiling
    posted_l = col_letter(val_col_posted)
    style_cell(ws, r, val_col_pct_ceil, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.0%").value = f"={posted_l}{r}/partner_ceiling"
    # Breach-In -i — inlines MAX(1, eur_actuals_through_month). At the
    # TOTAL level we treat aggregate Posted=0 as "no spend yet" → "—".
    style_cell(ws, r, val_col_breach, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({posted_l}{r}=0,"—",'
        f"(partner_ceiling-{posted_l}{r})"
        f"/({posted_l}{r}/MAX(1,eur_actuals_through_month)))"
    )


def _write_total_row_plus_i(ws, total_row, first_wp_row, last_wp_row, alayout):
    """Write TOTAL row +i mirrors (default-visible)."""
    r = total_row
    val_col_slope_pi = alayout["val_col_slope_pi"]
    val_col_budget_pi = alayout["val_col_budget_pi"]
    val_col_diff_pi = alayout["val_col_diff_pi"]
    val_col_posted_pi = alayout["val_col_posted_pi"]
    val_col_pct_ceil_pi = alayout["val_col_pct_ceil_pi"]
    val_col_breach_pi = alayout["val_col_breach_pi"]

    # Slope +i = SUM over WP +i Slope column.
    slope_pi_range = (f"{col_letter(val_col_slope_pi)}{first_wp_row}:"
                      f"{col_letter(val_col_slope_pi)}{last_wp_row}")
    style_cell(ws, r, val_col_slope_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = f"=SUM({slope_pi_range})"
    # Budget +i = partner_ceiling * 1.25. (The partner_ceiling named range
    # is -i; consumers apply ×1.25 — no separate +i named range.)
    style_cell(ws, r, val_col_budget_pi, font=GREEN_XREF, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = "=ROUND(partner_ceiling*1.25,0)"
    # Posted +i = SUM over WP +i Posted column.
    posted_pi_range = (f"{col_letter(val_col_posted_pi)}{first_wp_row}:"
                       f"{col_letter(val_col_posted_pi)}{last_wp_row}")
    style_cell(ws, r, val_col_posted_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = f"=SUM({posted_pi_range})"
    # Variance +i = Budget +i - Posted +i
    style_cell(ws, r, val_col_diff_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="€#,##0").value = (
        f"={col_letter(val_col_budget_pi)}{r}"
        f"-{col_letter(val_col_posted_pi)}{r}"
    )
    # % of Budget +i = Posted +i / (partner_ceiling * 1.25)
    posted_pi_l = col_letter(val_col_posted_pi)
    style_cell(ws, r, val_col_pct_ceil_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.0%").value = f"={posted_pi_l}{r}/(partner_ceiling*1.25)"
    # Breach-In +i — same shape as -i, against +i cells.
    style_cell(ws, r, val_col_breach_pi, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0").value = (
        f'=IF({posted_pi_l}{r}=0,"—",'
        f"((partner_ceiling*1.25)-{posted_pi_l}{r})"
        f"/({posted_pi_l}{r}/MAX(1,eur_actuals_through_month)))"
    )


def _apply_conditional_formatting(ws, first_wp_row, last_wp_row, total_row,
                                  alayout):
    """Apply all conditional-formatting rules.

    CF priority is order-sensitive (first added = highest priority). The
    blocks below preserve the original add-order verbatim: -i variance,
    -i % ceiling, -i breach, then +i mirrors (variance, % ceiling,
    breach), finally the reconciliation grey rule on the monthly grid.
    """
    val_col_diff = alayout["val_col_diff"]
    val_col_pct_ceil = alayout["val_col_pct_ceil"]
    val_col_breach = alayout["val_col_breach"]
    val_col_diff_pi = alayout["val_col_diff_pi"]
    val_col_pct_ceil_pi = alayout["val_col_pct_ceil_pi"]
    val_col_breach_pi = alayout["val_col_breach_pi"]
    month_col_start = alayout["month_col_start"]
    month_col_end = alayout["month_col_end"]

    # -- Conditional formatting: red on negative variance --
    diff_col_l = col_letter(val_col_diff)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))

    # -- Conditional formatting: % Ceiling column (amber >ACTUALS_AMBER_PCT_CEIL, red >100%) --
    # Red first = highest priority (lower number wins in Excel CF)
    pct_col_l = col_letter(val_col_pct_ceil)
    pct_range = f"{pct_col_l}{first_wp_row}:{pct_col_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_range, FormulaRule(
        formula=[f"{pct_col_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"], fill=CF_AMBER_FILL,
    ))

    # -- Conditional formatting: Breach In column (amber <ACTUALS_AMBER_BREACH_MONTHS, red <ACTUALS_RED_BREACH_MONTHS) --
    # Red first = highest priority
    breach_col_l = col_letter(val_col_breach)
    breach_range = f"{breach_col_l}{first_wp_row}:{breach_col_l}{last_wp_row}"
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(breach_range, FormulaRule(
        formula=[f"{breach_col_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"], fill=CF_AMBER_FILL,
    ))

    # ── Conditional formatting: +i mirror block. Every -i CF rule above
    # has a +i sibling targeting the mirror columns. ──
    # Variance +i: red if < 0
    diff_pi_col_l = col_letter(val_col_diff_pi)
    diff_pi_range = f"{diff_pi_col_l}{first_wp_row}:{diff_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(diff_pi_range, FormulaRule(
        formula=[f"{diff_pi_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))
    # % of Budget +i: red >100%, amber >=ACTUALS_AMBER_PCT_CEIL
    pct_pi_col_l = col_letter(val_col_pct_ceil_pi)
    pct_pi_range = f"{pct_pi_col_l}{first_wp_row}:{pct_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(pct_pi_range, FormulaRule(
        formula=[f"{pct_pi_col_l}{first_wp_row}>1"], fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(pct_pi_range, FormulaRule(
        formula=[f"{pct_pi_col_l}{first_wp_row}>{ACTUALS_AMBER_PCT_CEIL}"],
        fill=CF_AMBER_FILL,
    ))
    # Breach-In +i: red/amber bands
    breach_pi_col_l = col_letter(val_col_breach_pi)
    breach_pi_range = f"{breach_pi_col_l}{first_wp_row}:{breach_pi_col_l}{last_wp_row}"
    ws.conditional_formatting.add(breach_pi_range, FormulaRule(
        formula=[f"{breach_pi_col_l}{first_wp_row}<{ACTUALS_RED_BREACH_MONTHS}"],
        fill=CF_RED_FILL,
    ))
    ws.conditional_formatting.add(breach_pi_range, FormulaRule(
        formula=[f"{breach_pi_col_l}{first_wp_row}<{ACTUALS_AMBER_BREACH_MONTHS}"],
        fill=CF_AMBER_FILL,
    ))

    # ── Reconciliation visual: grey out monthly Actual EUR cells whose
    # month index is past the user-controlled `eur_actuals_through_month`
    # threshold. Those months haven't been closed by finance yet, so any
    # value (including blanks) should read as stale/incoming rather than
    # authoritative. Mirror cell on Actuals-Hours forecast band gets the
    # opposite treatment (grey when month <= threshold).
    ms = col_letter(month_col_start)
    me = col_letter(month_col_end)
    actual_eur_range = f"{ms}{first_wp_row}:{me}{total_row}"
    ws.conditional_formatting.add(actual_eur_range, FormulaRule(
        formula=[
            f"COLUMN({ms}{first_wp_row})-COLUMN(${ms}$1)+1"
            f">eur_actuals_through_month"
        ],
        fill=CF_GREY_FILL,
    ))


def _apply_outline_group(ws, alayout):
    """Apply outline grouping to the collapsible -i band.

    Outline group: collapsible -i band (T3 indirect-view restructure —
    2026-04-22). Bands together: monthly -i grid (cols C..me), -i
    visible rollups (Slope/Budget/Variance/Posted/%/Breach), and the
    hidden helpers (posted_idx, posted_value). +i mirrors stay outside
    (default-visible). ``hidden=True`` collapses the band by default.
    Summary arrow points right so the group toggle sits next to the +i
    mirrors.
    """
    for col in alayout["minus_i_band_cols"]:
        dim = ws.column_dimensions[col_letter(col)]
        dim.outline_level = 1
        dim.hidden = True
    if ws.sheet_properties.outlinePr is None:
        from openpyxl.worksheet.properties import Outline
        ws.sheet_properties.outlinePr = Outline(summaryRight=True)
    else:
        ws.sheet_properties.outlinePr.summaryRight = True


def _write_through_month_strip(ws, first_wp_row, last_wp_row, alayout,
                               eur_through_override):
    """Write the row 1 EUR through-month control strip.

    A: label, B: override (blue; blank by default), C: auto-derived scan
    of the WP data block (excludes TOTAL), D: effective = override-if-
    numeric-else-auto. ``eur_actuals_through_month`` now resolves to the
    effective cell on this sheet rather than Config!B7.
    """
    month_col_start = alayout["month_col_start"]
    month_col_end = alayout["month_col_end"]
    style_cell(ws, 1, 1, value="EUR Actuals Through Month:",
               font=HEADER_FONT, border=False)
    ws.cell(row=1, column=1).comment = sized_comment(
        "Trust boundary for EUR-side cumulative figures. The effective "
        "value in column D is what downstream sheets consume via the "
        "named range `eur_actuals_through_month`. Override (B) beats auto-"
        "derived (C) only when it holds a number."
    )
    style_cell(ws, 1, 2, value=eur_through_override, font=BLUE_INPUT, fmt="0",
               border=False)
    ws.cell(row=1, column=2).comment = sized_comment(
        "Optional override — enter an integer month index (1-based) to "
        "force a tighter EUR trust cutoff than the auto-derived value. "
        "Leave blank to defer to the auto scan in column C. Blue = user "
        "input."
    )
    ms_strip = col_letter(month_col_start)
    me_strip = col_letter(month_col_end)
    wp_scan_range = (
        f"${ms_strip}${first_wp_row}:${me_strip}${last_wp_row}"
    )
    style_cell(ws, 1, 3, font=BLACK_FORMULA, fmt="0", border=False).value = (
        f"=SUMPRODUCT(MAX((COLUMN({wp_scan_range})"
        f"-COLUMN(${ms_strip}$1)+1)"
        f"*ISNUMBER({wp_scan_range})))"
    )
    ws.cell(row=1, column=3).comment = sized_comment(
        "Auto-derived: largest 1-based month index with any numeric "
        "value in the WP data block (TOTAL row excluded). Returns 0 "
        "when no actuals have been logged yet. Do not edit."
    )
    style_cell(ws, 1, 4, font=BLACK_FORMULA, fmt="0", border=False).value = (
        "=IF(ISNUMBER(B1),B1,C1)"
    )
    ws.cell(row=1, column=4).comment = sized_comment(
        "Effective EUR trust cutoff consumed by Actuals, Summary, and "
        "Trends via the named range `eur_actuals_through_month`. Override "
        "in B wins when numeric; otherwise falls through to the auto scan "
        "in C. Do not edit."
    )


def _register_named_ranges(wb):
    """Register the ``eur_actuals_through_month`` named range pointing at
    the row 1 effective cell (Actuals!$D$1)."""
    wb.defined_names.add(DefinedName(
        "eur_actuals_through_month",
        attr_text=f"{SHEET_ACTUALS}!$D$1",
    ))


def build_actuals_sheet(wb, layout, *, project_params=None, work_packages=None,
                        preserved=None,
                        eur_through_override: int | None = None):
    """Build the Actuals sheet for posted-cumulative EUR per WP per month.

    Monthly EUR cells are always left empty (blue input) for manual entry
    or demo seeding. When ``preserved`` is provided (rebuild), those values
    are restored. Per-cell convention: running EUR posted through that
    month (blanks allowed inside the active window per design §1).

    Args:
        wb: The parent Workbook.
        layout: Layout dict from build_plan_sheet(), needs config_first_wp_row.

    Returns:
        Layout dict with keys: asof_row, year_row, month_row,
        first_wp_row, last_wp_row, total_row + per-WP column refs.
    """
    if project_params is None or work_packages is None:
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES

    alayout = _resolve_actuals_layout(project_params, work_packages, layout)

    ws = wb.create_sheet(SHEET_ACTUALS)
    ws.sheet_properties.tabColor = "0000FF"

    # Row 1 holds the EUR through-month control strip (label / override /
    # auto-derived / effective). The strip formulas depend on
    # ``first_wp_row``/``last_wp_row``, so we write the placeholder cells
    # here to claim the row geometry and back-patch formulas + comments
    # after the WP block has been written. Trust-boundary invariant lives
    # in ``docs/spec.md`` § Two-cadence actuals.
    r = 2

    # -- "Posted through:" header row (shifted to row 2 by the through-
    #     month strip) --
    _write_asof_header_row(ws, r, alayout)
    r += 1

    # -- Section headers --
    _write_section_headers(ws, r, alayout)
    r += 1

    year_row, month_row, r = build_year_month_headers(
        ws, r, project_params, alayout["month_col_start"],
        val_headers=[
            # -i visible rollup block (collapsible band)
            (alayout["val_col_slope"], "Slope -i"),
            (alayout["val_col_budget"], "Budget -i"),
            (alayout["val_col_diff"], "Variance -i"),
            (alayout["val_col_posted"], "Posted -i"),
            (alayout["val_col_pct_ceil"], "% of Budget -i"),
            (alayout["val_col_breach"], "Breach-In -i"),
            # +i mirror block (default-visible)
            (alayout["val_col_slope_pi"], "Slope +i"),
            (alayout["val_col_budget_pi"], "Budget +i"),
            (alayout["val_col_diff_pi"], "Variance +i"),
            (alayout["val_col_posted_pi"], "Posted +i"),
            (alayout["val_col_pct_ceil_pi"], "% of Budget +i"),
            (alayout["val_col_breach_pi"], "Breach-In +i"),
            # Hidden helpers (-i side; outline-grouped)
            (alayout["val_col_posted_idx"], "posted_idx"),
            (alayout["val_col_posted_value"], "posted_value"),
        ],
    )
    first_wp_row = r

    _apply_header_fill_overrides(ws, year_row, month_row, alayout)
    _append_minus_i_year_anchor_suffix(ws, year_row, alayout["month_col_start"])

    # -- WP rows (blue input EUR cells) --
    last_wp_row = _write_wp_rows(
        ws, first_wp_row, alayout, work_packages, preserved, project_params,
    )

    # -- TOTAL row --
    total_row = last_wp_row + 1
    _write_total_row_monthly_grid(
        ws, total_row, first_wp_row, last_wp_row, alayout, project_params,
    )
    _write_total_row_minus_i(ws, total_row, first_wp_row, last_wp_row, alayout)
    _write_total_row_plus_i(ws, total_row, first_wp_row, last_wp_row, alayout)

    _apply_conditional_formatting(
        ws, first_wp_row, last_wp_row, total_row, alayout,
    )
    _apply_outline_group(ws, alayout)

    # ── Row 1: EUR through-month control strip ──
    _write_through_month_strip(
        ws, first_wp_row, last_wp_row, alayout, eur_through_override,
    )
    _register_named_ranges(wb)

    # Freeze panes
    ws.freeze_panes = f"C{first_wp_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "asof_row": 1,       # "Posted through:" header row (T7 addition; section header shifted to row 2)
        "year_row": year_row,
        "month_row": month_row,
        "first_wp_row": first_wp_row,
        "last_wp_row": last_wp_row,
        "total_row": total_row,
        # Posted-cumulative redesign keys (Stage A):
        "posted_col": alayout["val_col_posted"],
        "posted_col_plus_i": alayout["val_col_posted_pi"],
        "slope_col": alayout["val_col_slope"],
        "slope_col_plus_i": alayout["val_col_slope_pi"],
        "posted_idx_col": alayout["val_col_posted_idx"],
        "posted_value_col": alayout["val_col_posted_value"],
        "pct_ceiling_col": alayout["val_col_pct_ceil"],
        "breach_col": alayout["val_col_breach"],
        # Transitional alias (retire after Stage B Trends rewire):
        # `trends.py:69` reads `actuals_layout["cum_col"]`. Same column
        # slot as `posted_col`.
        "cum_col": alayout["val_col_posted"],
    }
