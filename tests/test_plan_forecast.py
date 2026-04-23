"""Tests for the Plan-Forecast sheet builder.

Plan-Forecast aggregates forecast-EUR leaves from Actuals-Hours into a
WP × month grid that mirrors the Plan sheet shape.

T4 (2026-04-22 indirect-view restructure) flipped the default-visible
rollup surface to with-indirect (+i) and wrapped the monthly grid + -i
rollups in a collapsible column outline group.
"""
import re

import pytest
from openpyxl import load_workbook
from openpyxl.formatting.rule import Rule

from budget_planner.config import (
    WORK_PACKAGES,
    PROJECT_MONTHS,
)
from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_PLAN_FORECAST,
    val_columns,
    month_col_end,
    wp_active_months,
)
from budget_planner.styles import (
    CF_GREY_FILL, GREY_FILL,
    WITH_INDIRECT_HEADER_FILL, WITHOUT_INDIRECT_HEADER_FILL,
    col_letter,
)


# ── Fixtures ──

@pytest.fixture(scope="module")
def assigned_wb(assigned_wb_path):
    return load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def empty_wb(fresh_wb_path):
    return load_workbook(fresh_wb_path, data_only=False)


@pytest.fixture(scope="module")
def pf(assigned_wb):
    return assigned_wb[SHEET_PLAN_FORECAST]


# ── Existence ──
# Adjacency to Actuals-Hours was a pre-v2 invariant; Phase 4 moved
# Actuals-Hours to position 3, so Plan-Forecast no longer sits next to it.


def test_plan_forecast_sheet_exists(assigned_wb):
    assert SHEET_PLAN_FORECAST in assigned_wb.sheetnames


# ── Shape ──

def test_plan_forecast_wp_count(pf):
    """One WP row per WORK_PACKAGES entry.

    Derives WP-row count by scanning col A for green cross-refs to Config
    (``=Config!A<n>``). The WP block sits below the compact header + section
    header + year/month rows; scanning catches all of them without
    hard-coding row positions.
    """
    wp_id_refs = 0
    for r in range(1, pf.max_row + 1):
        v = pf.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^=Config!A\d+$", v):
            wp_id_refs += 1
    assert wp_id_refs == len(WORK_PACKAGES)


# ── Formula content ──

def _find_wp_row(ws, wp_idx):
    """Return the Plan-Forecast row for WORK_PACKAGES[wp_idx].

    Walks col-A cross-refs in order; the nth one corresponds to the nth WP.
    """
    count = 0
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^=Config!A\d+$", v):
            if count == wp_idx:
                return r
            count += 1
    raise AssertionError(f"No WP row for idx {wp_idx}")


def test_plan_forecast_cell_sums_actuals_hours_leaves(pf):
    """WP3 has Alice + Dave assigned (2 persons).

    For an active month, the cell formula must be:
        =SUM('Actuals-Hours'!<cellA>, 'Actuals-Hours'!<cellB>)

    with correct row positions (one per assignee WP3 row on Actuals-Hours)
    and identical columns (both at forecast_col_start + month_idx).
    """
    wp3_idx = 2
    wp_id = WORK_PACKAGES[wp3_idx]["id"]
    assert wp_id == "WP3"

    # Pick first active month in WP3
    active = sorted(wp_active_months(wp_id, WORK_PACKAGES))
    month_num = active[0]  # M1
    m_idx = month_num - 1
    col = MONTH_COL_START + m_idx

    row = _find_wp_row(pf, wp3_idx)
    formula = pf.cell(row, col).value
    assert isinstance(formula, str), f"Expected formula string, got {formula!r}"
    assert formula.startswith("=SUM("), formula
    assert "Actuals-Hours" in formula, formula

    # Extract the referenced row numbers — they must be distinct (one per
    # assignee) and the column must be the same across refs.
    refs = re.findall(r"Actuals-Hours'?!([A-Z]+)(\d+)", formula)
    # 2 assignees for WP3 in demo: Alice + Dave
    assert len(refs) == 2, f"Expected 2 leaf refs, got {refs} in {formula!r}"
    cols = {r[0] for r in refs}
    rows = {r[1] for r in refs}
    assert len(cols) == 1, f"Leaves must share a column; got {cols}"
    assert len(rows) == 2, f"Leaves must point at distinct rows; got {rows}"


def test_plan_forecast_wp_with_no_assignees_is_zero(empty_wb):
    """With no assignments, every WP row monthly cell (active months) is =0.

    Inactive months remain empty/grey (not =0).
    """
    pf = empty_wb[SHEET_PLAN_FORECAST]
    wp_idx = 0  # WP1
    wp_id = WORK_PACKAGES[wp_idx]["id"]
    active = wp_active_months(wp_id, WORK_PACKAGES)
    row = _find_wp_row(pf, wp_idx)

    # Pick an active month — formula must be =0 (explicit, not SUM of nothing)
    month_num = min(active)
    col = MONTH_COL_START + (month_num - 1)
    assert pf.cell(row, col).value == "=0"


def test_plan_forecast_inactive_months_grey(pf):
    """WP4 starts M2 (demo schedule has WP4 at M2, WP2 is M1-M6).

    Pick a WP whose schedule start > 1 so there is an inactive month at M1.
    WP4 (idx=3) starts at M2 per WORK_PACKAGES.
    """
    wp_idx = 3  # WP4 — schedule starts M2
    wp_id = WORK_PACKAGES[wp_idx]["id"]
    assert wp_id == "WP4"
    active = wp_active_months(wp_id, WORK_PACKAGES)
    assert 1 not in active  # M1 is inactive for WP4

    row = _find_wp_row(pf, wp_idx)
    cell = pf.cell(row, MONTH_COL_START)  # M1 col
    assert cell.value in (None, "")
    fill = cell.fill
    assert fill is not None
    assert str(getattr(fill.fgColor, "rgb", "")).upper().endswith(
        str(GREY_FILL.fgColor.rgb).upper()
    ), f"Inactive month fill should be GREY_FILL, got {fill.fgColor.rgb!r}"


def test_plan_forecast_variance_red_cf(pf):
    """At least one CF rule with red fill applies on the Variance -i column.

    The -i variance CF was the historical behaviour; T4 added a +i sibling
    (asserted separately below). This test pins the -i rule still exists.
    """
    vcols = val_columns(PROJECT_MONTHS)
    diff_col = col_letter(vcols["diff"])

    red_rules = 0
    for cf_range in pf.conditional_formatting:
        rng_str = str(cf_range)
        if diff_col not in rng_str:
            continue
        rules = pf.conditional_formatting[cf_range]
        for rule in rules:
            assert isinstance(rule, Rule)
            fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
            if fill is None:
                continue
            bg = getattr(fill.bgColor, "rgb", None)
            if bg and "C7CE" in str(bg).upper():  # CF_RED_FILL signature
                red_rules += 1
    assert red_rules >= 1, (
        f"Expected ≥1 red CF rule on variance col {diff_col}"
    )


def test_plan_forecast_variance_plus_i_red_cf(pf):
    """Red CF rule mirrors onto the Variance +i column (T4 review-fix)."""
    vcols = val_columns(PROJECT_MONTHS)
    diff_pi_col = col_letter(vcols["diff_plus_i"])

    red_rules = 0
    for cf_range in pf.conditional_formatting:
        rng_str = str(cf_range)
        if diff_pi_col not in rng_str:
            continue
        rules = pf.conditional_formatting[cf_range]
        for rule in rules:
            assert isinstance(rule, Rule)
            fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
            if fill is None:
                continue
            bg = getattr(fill.bgColor, "rgb", None)
            if bg and "C7CE" in str(bg).upper():
                red_rules += 1
    assert red_rules >= 1, (
        f"Expected ≥1 red CF rule on variance +i col {diff_pi_col}"
    )


def test_plan_forecast_total_row_sums_wp_rows(pf):
    """TOTAL row per-month formula sums the WP rows above it."""
    # Locate WP rows + total row by scanning col A.
    wp_rows = []
    for r in range(1, pf.max_row + 1):
        v = pf.cell(r, 1).value
        if isinstance(v, str) and re.match(r"^=Config!A\d+$", v):
            wp_rows.append(r)
    assert len(wp_rows) == len(WORK_PACKAGES)
    first_wp, last_wp = wp_rows[0], wp_rows[-1]
    total_row = last_wp + 1

    # Month-1 column
    col = MONTH_COL_START
    formula = pf.cell(total_row, col).value
    assert isinstance(formula, str) and formula.startswith("=SUM(")
    # Should reference the WP block: col{first_wp}:col{last_wp}
    cl = col_letter(col)
    assert f"{cl}{first_wp}:{cl}{last_wp}" in formula, formula


# ── Validation columns ──

def test_plan_forecast_budget_minus_i_column_refs_config(pf):
    """Budget -i column on each WP row cross-refs Config!E<wp_row> (col 5
    = Budget -i, numerically identical to the legacy Ceiling value)."""
    vcols = val_columns(PROJECT_MONTHS)
    budget_col = vcols["budget"]

    # Take first WP row
    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, budget_col).value
    assert isinstance(val, str) and val.startswith("=Config!E"), val


def test_plan_forecast_grey_cf_past_hours_actuals_threshold(pf):
    """A cosmetic grey CF shades forecast-EUR cells whose month index is
    past ``hours_actuals_through_month`` — no hours yet, so the forecast
    aggregates zero and should read as empty/unreliable rather than zero."""
    target_grey = str(CF_GREY_FILL.bgColor.rgb).upper()
    # MONTH_COL_START = col C; the month band spans C..<end>.
    month_start_l = col_letter(MONTH_COL_START)

    found_formula = None
    for cf_range in pf.conditional_formatting:
        rng_str = str(cf_range)
        if month_start_l not in rng_str:
            continue
        for rule in pf.conditional_formatting[cf_range]:
            fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
            if fill is None:
                continue
            bg = str(getattr(fill.bgColor, "rgb", "") or "").upper()
            if not bg.endswith(target_grey):
                continue
            for f in (rule.formula or []):
                if "hours_actuals_through_month" in f:
                    found_formula = f
                    break

    assert found_formula is not None, (
        "Missing grey CF keyed on hours_actuals_through_month on the "
        "Plan-Forecast monthly range"
    )
    assert ">hours_actuals_through_month" in found_formula, found_formula


def test_plan_forecast_sum_column_totals_months(pf):
    """Sum -i column = SUM(<month range>) for each WP row."""
    vcols = val_columns(PROJECT_MONTHS)
    sum_col = vcols["sum"]
    end_col = month_col_end(PROJECT_MONTHS)

    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, sum_col).value
    assert isinstance(val, str) and val.startswith("=SUM("), val
    assert f"{col_letter(MONTH_COL_START)}{wp_row}:{col_letter(end_col)}{wp_row}" in val


# ── T4 (2026-04-22): indirect-view +i mirror block ──

def _header_text_at(ws, col, *, scan_rows=8):
    """Return the rollup-column month-row header text for ``col``.

    Walks rows 1..scan_rows looking for a short (<20 chars) string in the
    column that contains a ``+i`` / ``-i`` suffix. The length filter skips
    the merged section-header row ("FORECAST EUR (from hours) +i") which
    is long prose, and surfaces the month-row rollup label
    ("Sum +i", "Budget -i", etc.).
    """
    for r in range(1, scan_rows + 1):
        v = ws.cell(r, col).value
        if (
            isinstance(v, str)
            and ("+i" in v or "-i" in v)
            and len(v) < 20
        ):
            return v
    return None


def test_plan_forecast_plus_i_rollup_headers(pf):
    """The five +i mirror columns carry headers with the '+i' suffix and
    WITH_INDIRECT_HEADER_FILL on the month-header row."""
    vcols = val_columns(PROJECT_MONTHS)
    expected = {
        vcols["sum_plus_i"]: "Sum +i",
        vcols["budget_plus_i"]: "Budget +i",
        vcols["diff_plus_i"]: "Variance +i",
        vcols["cum_plus_i"]: "Cum EUR +i",
        vcols["pct_ceil_plus_i"]: "% of Budget +i",
    }
    target_fill = str(WITH_INDIRECT_HEADER_FILL.fgColor.rgb).upper()
    for col, want_label in expected.items():
        got = _header_text_at(pf, col)
        assert got == want_label, (
            f"Col {col_letter(col)}: expected header {want_label!r}, got {got!r}"
        )
        # Month-header row = where the label is written. Find it and
        # confirm the fill matches WITH_INDIRECT_HEADER_FILL.
        for r in range(1, 10):
            if pf.cell(r, col).value == want_label:
                fill = pf.cell(r, col).fill
                bg = str(getattr(fill.fgColor, "rgb", "") or "").upper()
                assert bg.endswith(target_fill), (
                    f"Col {col_letter(col)}: expected WITH_INDIRECT fill, got {bg!r}"
                )
                break


def test_plan_forecast_minus_i_rollup_headers(pf):
    """The five -i sibling columns carry headers with the '-i' suffix and
    WITHOUT_INDIRECT_HEADER_FILL on the month-header row."""
    vcols = val_columns(PROJECT_MONTHS)
    expected = {
        vcols["sum"]: "Sum -i",
        vcols["budget"]: "Budget -i",
        vcols["diff"]: "Variance -i",
        vcols["cum"]: "Cum EUR -i",
        vcols["pct_ceil"]: "% of Budget -i",
    }
    target_fill = str(WITHOUT_INDIRECT_HEADER_FILL.fgColor.rgb).upper()
    for col, want_label in expected.items():
        got = _header_text_at(pf, col)
        assert got == want_label, (
            f"Col {col_letter(col)}: expected header {want_label!r}, got {got!r}"
        )
        for r in range(1, 10):
            if pf.cell(r, col).value == want_label:
                fill = pf.cell(r, col).fill
                bg = str(getattr(fill.fgColor, "rgb", "") or "").upper()
                assert bg.endswith(target_fill), (
                    f"Col {col_letter(col)}: expected WITHOUT_INDIRECT fill, got {bg!r}"
                )
                break


def test_plan_forecast_sum_plus_i_formula(pf):
    """Sum +i = ROUND(Sum -i * 1.25, 0) per openpyxl.md EUR rule."""
    vcols = val_columns(PROJECT_MONTHS)
    sum_pi_col = vcols["sum_plus_i"]
    sum_col_l = col_letter(vcols["sum"])

    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, sum_pi_col).value
    assert isinstance(val, str), val
    assert val == f"=ROUND({sum_col_l}{wp_row}*1.25,0)", val


def test_plan_forecast_budget_plus_i_xrefs_named_range(pf):
    """Budget +i = INDEX(wp_budget_with_indirect, i+1). No ×1.25 applied —
    the named range already targets the Config Budget +i column."""
    vcols = val_columns(PROJECT_MONTHS)
    budget_pi_col = vcols["budget_plus_i"]

    # First WP row — i=0 → INDEX(..,1)
    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, budget_pi_col).value
    assert isinstance(val, str), val
    assert val == "=INDEX(wp_budget_with_indirect,1)", val
    assert "*1.25" not in val, (
        "Budget +i xrefs an already-with-indirect source; must not apply ×1.25"
    )

    # Second WP row — i=1 → INDEX(..,2)
    wp_row_2 = _find_wp_row(pf, 1)
    val_2 = pf.cell(wp_row_2, budget_pi_col).value
    assert val_2 == "=INDEX(wp_budget_with_indirect,2)", val_2


def test_plan_forecast_cum_plus_i_formula(pf):
    """Cum EUR +i = ROUND(Cum EUR -i * 1.25, 0)."""
    vcols = val_columns(PROJECT_MONTHS)
    cum_pi_col = vcols["cum_plus_i"]
    cum_col_l = col_letter(vcols["cum"])

    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, cum_pi_col).value
    assert isinstance(val, str), val
    assert val == f"=ROUND({cum_col_l}{wp_row}*1.25,0)", val


def test_plan_forecast_variance_plus_i_formula(pf):
    """Variance +i = Budget +i - Sum +i."""
    vcols = val_columns(PROJECT_MONTHS)
    diff_pi_col = vcols["diff_plus_i"]
    sum_pi_l = col_letter(vcols["sum_plus_i"])
    budget_pi_l = col_letter(vcols["budget_plus_i"])

    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, diff_pi_col).value
    assert isinstance(val, str), val
    assert val == f"={budget_pi_l}{wp_row}-{sum_pi_l}{wp_row}", val


def test_plan_forecast_pct_plus_i_formula(pf):
    """% of Budget +i = Cum +i / Budget +i."""
    vcols = val_columns(PROJECT_MONTHS)
    pct_pi_col = vcols["pct_ceil_plus_i"]
    cum_pi_l = col_letter(vcols["cum_plus_i"])
    budget_pi_l = col_letter(vcols["budget_plus_i"])

    wp_row = _find_wp_row(pf, 0)
    val = pf.cell(wp_row, pct_pi_col).value
    assert isinstance(val, str), val
    assert val == f"={cum_pi_l}{wp_row}/{budget_pi_l}{wp_row}", val


def test_plan_forecast_monthly_grid_in_outline_group(pf):
    """Monthly forecast grid cols are wrapped in a column outline group
    (outline_level=1, hidden=True) per T4's collapsible -i band.

    The -i rollup cols are also in the band; the +i mirrors are NOT."""
    vcols = val_columns(PROJECT_MONTHS)
    end_col = month_col_end(PROJECT_MONTHS)

    # Spot-check the monthly grid (a few columns, not all 48).
    for col in (MONTH_COL_START, MONTH_COL_START + 1, end_col):
        dim = pf.column_dimensions[col_letter(col)]
        assert dim.outline_level == 1, (
            f"Monthly-grid col {col_letter(col)} should be outline_level=1"
        )
        assert dim.hidden is True, (
            f"Monthly-grid col {col_letter(col)} should be hidden by default"
        )

    # -i rollup cols must also sit inside the band.
    for key in ("sum", "budget", "diff", "cum", "pct_ceil"):
        dim = pf.column_dimensions[col_letter(vcols[key])]
        assert dim.outline_level == 1, (
            f"-i rollup col {col_letter(vcols[key])} should be in outline band"
        )
        assert dim.hidden is True

    # +i mirrors must NOT be in the collapsible band.
    for key in ("sum_plus_i", "budget_plus_i", "diff_plus_i",
                "cum_plus_i", "pct_ceil_plus_i"):
        dim = pf.column_dimensions[col_letter(vcols[key])]
        assert (dim.outline_level or 0) == 0, (
            f"+i mirror col {col_letter(vcols[key])} must stay outside "
            f"the collapsible band"
        )


def test_plan_forecast_no_ceiling_residue(pf):
    """Header cells and the compact row-1 block no longer contain the
    word 'Ceiling' — T4 renamed all such surfaces to 'Budget'."""
    # Row-1 compact label cell (col 6 holds the "Budget:" label).
    assert pf.cell(1, 6).value == "Budget:"
    # Scan rollup-column month-header row for any "Ceiling" text.
    for r in range(1, 10):
        for c in range(1, pf.max_column + 1):
            v = pf.cell(r, c).value
            if isinstance(v, str):
                assert "Ceiling" not in v and "ceiling" not in v, (
                    f"Ceiling/ceiling residue at {col_letter(c)}{r}: {v!r}"
                )
