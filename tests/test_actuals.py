"""Tests for the Actuals sheet builder."""

import pytest
import openpyxl

from budget_planner.config import (
    WORK_PACKAGES,
    NUM_WPS,
    MONTH_COL_START,
    VAL_COL_SLOPE,
    VAL_COL_BUDGET,
    VAL_COL_DIFF,
    VAL_COL_POSTED,
    VAL_COL_PCT_CEIL,
    VAL_COL_BREACH,
)
from budget_planner.layout import wp_month_bounds
from budget_planner.styles import col_letter


@pytest.fixture(scope="module")
def actuals_workbook(demo_wb_path):
    return openpyxl.load_workbook(demo_wb_path, data_only=False)


def test_actuals_wp_count(actuals_workbook):
    """Actuals has one row per WP between header and TOTAL."""
    ws = actuals_workbook["Actuals"]
    # Find TOTAL row
    total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None
    # Find first WP row (row after month headers -- look for "WP ID" in col A)
    month_header_row = None
    for r in range(1, total_row):
        if ws.cell(r, 1).value == "WP ID":
            month_header_row = r
            break
    assert month_header_row is not None
    first_wp_row = month_header_row + 1
    wp_count = total_row - first_wp_row
    assert wp_count == NUM_WPS


def test_actuals_demo_values_seeded(actuals_workbook):
    """Fresh demo workbook has non-zero EUR in active months up to DEMO_CURRENT_MONTH."""
    from budget_planner.config import DEMO_CURRENT_MONTH

    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    wp = WORK_PACKAGES[0]  # WP1: active from M1, multiplier 1.0
    wp_start, wp_end = wp_month_bounds(wp["id"], WORK_PACKAGES)
    row = first_wp_row
    # Active month within demo range should have a positive value
    c = MONTH_COL_START + wp_start - 1
    assert ws.cell(row, c).value > 0
    # Month after DEMO_CURRENT_MONTH should be empty (not seeded)
    if DEMO_CURRENT_MONTH < wp_end:
        c_after = MONTH_COL_START + DEMO_CURRENT_MONTH
        val = ws.cell(row, c_after).value
        assert val is None or val == 0, f"Expected empty/0 after demo range, got {val}"


def test_actuals_inactive_months_empty(actuals_workbook):
    """Inactive month cells are empty with grey fill."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    # Find any WP with start > 1 -- months before start should be inactive
    late_idx = next(i for i, w in enumerate(WORK_PACKAGES) if wp_month_bounds(w["id"], WORK_PACKAGES)[0] > 1)
    late_start = wp_month_bounds(WORK_PACKAGES[late_idx]["id"], WORK_PACKAGES)[0]
    row = first_wp_row + late_idx
    for m in range(late_start - 1):
        c = MONTH_COL_START + m
        cell = ws.cell(row, c)
        assert cell.value is None or cell.value == "", f"M{m + 1} should be inactive, got {cell.value}"
        assert cell.fill.fgColor and cell.fill.fgColor.rgb and "D9D9D9" in str(cell.fill.fgColor.rgb), (
            f"M{m + 1} should have grey fill"
        )


def test_actuals_slope_formula(actuals_workbook):
    """Slope -i is the avg-burn formula posted_value/posted_idx, guarded
    by IF(posted_idx=0, "", ...) per posted-cumulative design §3."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_SLOPE)
    formula = str(cell.value)
    assert formula.startswith("=IF("), formula
    # /0 guard: posted_idx=0 → ""
    assert '=0,""' in formula, formula


def test_actuals_ceiling_refs_plan(actuals_workbook):
    """Ceiling validation column references Plan config table."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_BUDGET)
    assert cell.value and "Config!E" in str(cell.value)


def test_actuals_variance_formula(actuals_workbook):
    """Variance -i = Budget -i - Posted -i (guarded by IF(posted_idx=0,
    \"\", ...))."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_DIFF)
    formula = str(cell.value)
    budget_col = col_letter(VAL_COL_BUDGET)
    posted_col = col_letter(VAL_COL_POSTED)
    assert budget_col in formula and posted_col in formula


def test_actuals_total_ceiling_uses_named_range(actuals_workbook):
    """TOTAL row ceiling cell uses partner_ceiling named range."""
    ws = actuals_workbook["Actuals"]
    total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None
    cell = ws.cell(total_row, VAL_COL_BUDGET)
    assert "partner_ceiling" in str(cell.value)


# ── Cumulative columns ──


def _find_first_wp_row(ws):
    """Find the first WP data row (row after 'WP ID' header)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r + 1
    return None


def _find_total_row(ws):
    """Find the TOTAL row."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            return r
    return None


def _find_year_row(ws):
    """Find the year header row (row before the month header row)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r - 1
    return None


def test_actuals_posted_eur_formula(actuals_workbook):
    """First WP row at VAL_COL_POSTED is IFERROR(INDEX(grid, posted_idx),
    \"\") — the latest filled cumulative through current_month per
    posted-cumulative design §2."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_POSTED)
    formula = str(cell.value)
    assert formula.startswith("=IFERROR(INDEX("), formula
    assert ',"")' in formula, formula


def test_actuals_pct_ceiling_formula(actuals_workbook):
    """First WP row at VAL_COL_PCT_CEIL divides Posted by Budget (guarded
    by IF(posted_idx=0, \"\", ...))."""
    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_PCT_CEIL)
    formula = str(cell.value)
    posted_l = col_letter(VAL_COL_POSTED)
    ceil_l = col_letter(VAL_COL_BUDGET)
    assert posted_l in formula and ceil_l in formula


def test_actuals_breach_formula(actuals_workbook):
    """First WP row at VAL_COL_BREACH has IF formula with cum-to-date slope.

    Denominator inlines ``MAX(1, eur_actuals_through_month)`` (post-2026-04-18
    cleanup; the ``months_elapsed_eur`` derived named range was dropped).
    See ``test_actuals_breach_todate.py`` for the full invariant.
    """
    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    cell = ws.cell(first_wp_row, VAL_COL_BREACH)
    formula = str(cell.value)
    assert "IF" in formula
    assert "MAX(1,eur_actuals_through_month)" in formula


def test_actuals_posted_headers(actuals_workbook):
    """Year row has -i Posted / % of Budget / Breach-In headers at the
    -i positions. Posted-cumulative redesign (2026-05-09) renamed
    ``Cum EUR -i`` → ``Posted -i`` (matching cell convention shift to
    posted-cumulative); ``% of Budget -i`` / ``Breach-In -i`` headers
    are unchanged. +i mirrors at offsets 10/11/12 covered by a sibling
    assert below.
    """
    ws = actuals_workbook["Actuals"]
    year_row = _find_year_row(ws)
    assert year_row is not None
    assert ws.cell(year_row, VAL_COL_POSTED).value == "Posted -i"
    assert ws.cell(year_row, VAL_COL_PCT_CEIL).value == "% of Budget -i"
    assert ws.cell(year_row, VAL_COL_BREACH).value == "Breach-In -i"


# ── +i mirror block (T3 indirect-view restructure 2026-04-22) ──


def test_actuals_plus_i_headers(actuals_workbook):
    """+i mirror block headers carry the '+i' suffix in the year row.

    Posted-cumulative redesign renamed Sum/Cum → Slope/Posted on both
    -i and +i sides; the four other +i headers (Budget / Variance /
    % of Budget / Breach-In) are unchanged.
    """
    from budget_planner.layout import val_columns
    from budget_planner.config import PROJECT_MONTHS

    ws = actuals_workbook["Actuals"]
    year_row = _find_year_row(ws)
    assert year_row is not None
    vcols = val_columns(PROJECT_MONTHS)
    assert ws.cell(year_row, vcols["slope_plus_i"]).value == "Slope +i"
    assert ws.cell(year_row, vcols["budget_plus_i"]).value == "Budget +i"
    assert ws.cell(year_row, vcols["diff_plus_i"]).value == "Variance +i"
    assert ws.cell(year_row, vcols["posted_plus_i"]).value == "Posted +i"
    assert ws.cell(year_row, vcols["pct_ceil_plus_i"]).value == "% of Budget +i"
    assert ws.cell(year_row, vcols["breach_plus_i"]).value == "Breach-In +i"


def test_actuals_plus_i_header_fills(actuals_workbook):
    """+i headers use WITH_INDIRECT_HEADER_FILL (E2EFDA); -i headers use
    WITHOUT_INDIRECT_HEADER_FILL (FBE5D6). Overrides the default
    LIGHT_BLUE_FILL that build_year_month_headers writes."""
    from budget_planner.layout import val_columns
    from budget_planner.config import PROJECT_MONTHS

    ws = actuals_workbook["Actuals"]
    year_row = _find_year_row(ws)
    assert year_row is not None
    vcols = val_columns(PROJECT_MONTHS)

    def _fg(cell):
        return str(cell.fill.fgColor.rgb).upper() if cell.fill.fgColor else ""

    # +i mirror cols use light-green fill
    for key in ("slope_plus_i", "budget_plus_i", "diff_plus_i", "posted_plus_i", "pct_ceil_plus_i", "breach_plus_i"):
        assert _fg(ws.cell(year_row, vcols[key])).endswith("E2EFDA"), (
            f"col {key} expected WITH_INDIRECT_HEADER_FILL (E2EFDA), got {_fg(ws.cell(year_row, vcols[key]))}"
        )
    # -i rollup cols use pale-orange fill
    for key in ("slope", "budget", "diff", "posted", "pct_ceil", "breach"):
        assert _fg(ws.cell(year_row, vcols[key])).endswith("FBE5D6"), (
            f"col {key} expected WITHOUT_INDIRECT_HEADER_FILL (FBE5D6), got {_fg(ws.cell(year_row, vcols[key]))}"
        )


def test_actuals_slope_plus_i_formula_multiplies_by_1_25(actuals_workbook):
    """Slope +i = ROUND(Slope -i * 1.25, 0) per D3 + openpyxl.md EUR
    rule, guarded by IF(posted_idx=0, \"\", ...) so empty (no postings)
    surfaces as blank in the +i mirror."""
    from budget_planner.layout import val_columns
    from budget_planner.config import PROJECT_MONTHS

    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    vcols = val_columns(PROJECT_MONTHS)
    formula = str(ws.cell(first_wp_row, vcols["slope_plus_i"]).value)
    slope_minus_l = col_letter(vcols["slope"])
    # Format: =IF(<idx_l>{r}=0,"",ROUND(<slope_l>{r}*1.25,0))
    assert f"ROUND({slope_minus_l}{first_wp_row}*1.25,0)" in formula, formula
    assert formula.startswith("=IF("), formula


def test_actuals_budget_plus_i_uses_named_range(actuals_workbook):
    """Budget +i resolves via INDEX(wp_budget_with_indirect, i+1). No
    ×1.25 — Config Budget +i is already with-indirect by D4."""
    from budget_planner.layout import val_columns
    from budget_planner.config import PROJECT_MONTHS

    ws = actuals_workbook["Actuals"]
    first_wp_row = _find_first_wp_row(ws)
    assert first_wp_row is not None
    vcols = val_columns(PROJECT_MONTHS)
    formula = str(ws.cell(first_wp_row, vcols["budget_plus_i"]).value)
    assert "wp_budget_with_indirect" in formula
    assert "INDEX" in formula
    # No ×1.25 applied since Config Budget +i is already +i.
    assert "1.25" not in formula, formula


def _col_outline_level(ws, col_idx):
    """Return outline_level for `col_idx` (1-based) by scanning
    ColumnDimension ranges — openpyxl stores contiguous runs on a single
    entry (e.g. cols 6..16 on key 'F'), so direct
    ``ws.column_dimensions[letter]`` misses cols inside a range. Returns
    0 if no covering range sets an outline level."""
    for dim in ws.column_dimensions.values():
        if dim.min <= col_idx <= dim.max and dim.outline_level:
            return dim.outline_level
    return 0


def test_actuals_minus_i_band_outline_level(actuals_workbook):
    """Monthly -i grid + -i visible rollups + hidden helpers sit in the
    outline band at level 1 (collapsed by default via hidden=True).

    Posted-cumulative redesign (2026-05-09) replaced the forecast
    monitoring block with two hidden helpers (posted_idx,
    posted_value) that join the band; +i mirrors stay outside.
    """
    from budget_planner.layout import (
        MONTH_COL_START,
        val_columns,
        month_col_end as _mce,
    )
    from budget_planner.config import PROJECT_MONTHS

    ws = actuals_workbook["Actuals"]
    vcols = val_columns(PROJECT_MONTHS)
    mce = _mce(PROJECT_MONTHS)

    inside = (
        *range(MONTH_COL_START, mce + 1),
        vcols["slope"],
        vcols["budget"],
        vcols["diff"],
        vcols["posted"],
        vcols["pct_ceil"],
        vcols["breach"],
        vcols["posted_idx"],
        vcols["posted_value"],
    )
    for col in inside:
        lvl = _col_outline_level(ws, col)
        assert lvl == 1, f"col {col_letter(col)} (idx {col}) should be in outline band, got outline_level={lvl}"

    # +i mirrors stay outside the outline band
    outside = (
        vcols["slope_plus_i"],
        vcols["budget_plus_i"],
        vcols["diff_plus_i"],
        vcols["posted_plus_i"],
        vcols["pct_ceil_plus_i"],
        vcols["breach_plus_i"],
    )
    for col in outside:
        lvl = _col_outline_level(ws, col)
        assert lvl == 0, (
            f"col {col_letter(col)} (+i mirror idx {col}) should NOT be in outline band, got outline_level={lvl}"
        )


def test_actuals_total_posted_formula(actuals_workbook):
    """TOTAL row at VAL_COL_POSTED is =SUM over the per-WP Posted column.

    Posted-cumulative redesign (2026-05-09) replaced the per-WP
    SUMPRODUCT-current_month-mask with the visible Posted -i column
    (latest filled cumulative through current_month). At the TOTAL
    level, aggregation is a simple SUM; the IFERROR-empty-string
    sentinel from per-WP cells coerces to 0 inside SUM.
    """
    ws = actuals_workbook["Actuals"]
    total_row = _find_total_row(ws)
    assert total_row is not None
    cell = ws.cell(total_row, VAL_COL_POSTED)
    formula = str(cell.value)
    assert formula.startswith("=SUM("), formula


def test_actuals_total_pct_ceiling_uses_partner_ceiling(actuals_workbook):
    """TOTAL row at VAL_COL_PCT_CEIL references partner_ceiling."""
    ws = actuals_workbook["Actuals"]
    total_row = _find_total_row(ws)
    assert total_row is not None
    cell = ws.cell(total_row, VAL_COL_PCT_CEIL)
    formula = str(cell.value)
    assert "partner_ceiling" in formula
