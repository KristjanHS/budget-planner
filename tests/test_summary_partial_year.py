"""Summary calendar-year WP block — per-column partial-year shading (Task 8).

The five per-year columns (Plan FTE, Plan €, Act €, Fcst €, Var €) shade grey
when the year is not fully covered by actuals:

- Act €  shades by ``eur_actuals_through_month``
- Fcst € shades by ``hours_actuals_through_month``
- Var €  shades by the stricter MIN of the two

``year_coverage_pct = months_of_year_covered / 12`` is the design metric; the
grey CF fires when the year-end project-month exceeds the through-month value.
The partial-year gate also tightens the existing Fcst < 0.8 * Plan red CF so
it only fires on fully-covered years.
"""

import pytest
import openpyxl

from budget_planner.layout import _calendar_year_groups
from budget_planner.config import (
    PROJECT_MONTHS, PROJECT_START,
)
from budget_planner.styles import col_letter


_YEAR_GROUP_START = 3
_COLS_PER_YEAR = 5  # Plan FTE, Plan €, Act €, Fcst €, Var €
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
_TOTAL_GROUP_START = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR
_GREY_HEX = "FFD9D9D9"
_RED_HEX = "FFFFC7CE"


@pytest.fixture(scope="module")
def summary_wb(assigned_wb_path):
    return openpyxl.load_workbook(assigned_wb_path, data_only=False)


def _collect_cf_rules(ws):
    """Return list of (sqref, [formulas], fill_bg_hex)."""
    results = []
    for cf in ws.conditional_formatting:
        sqref = str(cf.sqref)
        for rule in ws.conditional_formatting[cf]:
            formulas = list(rule.formula) if rule.formula else []
            fg = None
            if rule.dxf and rule.dxf.fill:
                fill = rule.dxf.fill
                if fill.bgColor and str(fill.bgColor.rgb) != "00000000":
                    fg = str(fill.bgColor.rgb)
                elif fill.fgColor and str(fill.fgColor.rgb) != "00000000":
                    fg = str(fill.fgColor.rgb)
            results.append((sqref, formulas, fg))
    return results


# ── Helpers ──

def _grey_rules(rules):
    return [r for r in rules if r[2] is not None and r[2].endswith(_GREY_HEX)]


def _red_rules(rules):
    return [r for r in rules if r[2] is not None and r[2].endswith(_RED_HEX)]


# ── Grey partial-year rules present per year ──

@pytest.mark.parametrize("y", range(_NUM_YEARS))
def test_act_eur_has_grey_cf_per_year(summary_wb, y):
    """Act € column has a grey CF gated on eur_actuals_through_month."""
    ws = summary_wb["Summary"]
    rules = _grey_rules(_collect_cf_rules(ws))
    group_start, span = _CAL_GROUPS[y]
    year_end_pm = group_start + span
    sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
    act_col_l = col_letter(sc + 2)
    expected_formula = f"eur_actuals_through_month<{year_end_pm}"
    matches = [
        r for r in rules
        if r[0].startswith(f"{act_col_l}")
        and any(expected_formula == f for f in r[1])
    ]
    assert len(matches) == 1, (
        f"Year {y} Act € grey CF: expected 1 rule with formula "
        f"{expected_formula!r}, got {len(matches)}"
    )


@pytest.mark.parametrize("y", range(_NUM_YEARS))
def test_fcst_eur_has_grey_cf_per_year(summary_wb, y):
    """Fcst € column has a grey CF gated on hours_actuals_through_month."""
    ws = summary_wb["Summary"]
    rules = _grey_rules(_collect_cf_rules(ws))
    group_start, span = _CAL_GROUPS[y]
    year_end_pm = group_start + span
    sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
    fcst_col_l = col_letter(sc + 3)
    expected_formula = f"hours_actuals_through_month<{year_end_pm}"
    matches = [
        r for r in rules
        if r[0].startswith(f"{fcst_col_l}")
        and any(expected_formula == f for f in r[1])
    ]
    assert len(matches) == 1, (
        f"Year {y} Fcst € grey CF: expected 1 rule with formula "
        f"{expected_formula!r}, got {len(matches)}"
    )


@pytest.mark.parametrize("y", range(_NUM_YEARS))
def test_var_eur_has_grey_cf_per_year_with_min(summary_wb, y):
    """Var € column has a grey CF using the stricter MIN of both through-months."""
    ws = summary_wb["Summary"]
    rules = _grey_rules(_collect_cf_rules(ws))
    group_start, span = _CAL_GROUPS[y]
    year_end_pm = group_start + span
    sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
    var_col_l = col_letter(sc + 4)
    matches = [
        r for r in rules
        if r[0].startswith(f"{var_col_l}")
        and any(
            "MIN(eur_actuals_through_month,hours_actuals_through_month)" in f
            and f"<{year_end_pm}" in f
            for f in r[1]
        )
    ]
    assert len(matches) == 1, (
        f"Year {y} Var € grey CF (MIN): expected 1 rule, got {len(matches)}"
    )


# ── Existing Fcst<0.8*Plan CF is now gated ──

def test_fcst_red_cf_gated_on_fully_covered_year(summary_wb):
    """Red Fcst < 0.8*Plan CF formula now wraps AND(hours_actuals_through_month>=N,...)."""
    ws = summary_wb["Summary"]
    rules = _red_rules(_collect_cf_rules(ws))
    # Per-year red Fcst rules (range sqref, "0.8*" in formula).
    per_year_fcst_red = [
        r for r in rules
        if ":" in r[0]
        and any("0.8*" in f for f in r[1])
    ]
    # num_years per-year rules + 1 project-total rule
    assert len(per_year_fcst_red) == _NUM_YEARS + 1, (
        f"Expected {_NUM_YEARS + 1} red Fcst CFs, got {len(per_year_fcst_red)}"
    )

    # Walk each per-year rule's formula and confirm it's gated.
    gated = [
        r for r in per_year_fcst_red
        if any(
            f.startswith("AND(hours_actuals_through_month>=")
            for f in r[1]
        )
    ]
    # The project-total Fcst red CF is intentionally NOT gated (spans all years),
    # so we expect exactly num_years gated rules.
    assert len(gated) == _NUM_YEARS, (
        f"Expected {_NUM_YEARS} gated per-year Fcst red CFs, got {len(gated)}"
    )

    # Confirm each year's threshold N matches group_start + span.
    for y, (group_start, span) in enumerate(_CAL_GROUPS):
        year_end_pm = group_start + span
        token = f"AND(hours_actuals_through_month>={year_end_pm},"
        hits = [r for r in gated if any(token in f for f in r[1])]
        assert len(hits) == 1, (
            f"Year {y}: expected gated Fcst red CF with threshold "
            f"{year_end_pm}, got {len(hits)}"
        )


# ── PROJECT TOTAL block has no grey CF ──

def test_project_total_block_has_no_grey_cf(summary_wb):
    """PROJECT TOTAL columns (across-years) must have no grey partial-year CF."""
    ws = summary_wb["Summary"]
    rules = _grey_rules(_collect_cf_rules(ws))
    total_col_letters = {
        col_letter(_TOTAL_GROUP_START + i) for i in range(_COLS_PER_YEAR)
    }
    for sqref, formulas, _fg in rules:
        # sqref like "H5:H9" — take column prefix
        col_prefix = sqref.split(":", 1)[0].rstrip("0123456789")
        assert col_prefix not in total_col_letters, (
            f"Grey CF unexpectedly applied to PROJECT TOTAL col {col_prefix} "
            f"(sqref={sqref}, formula={formulas})"
        )


# ── Negative-variance CFs unchanged ──

def test_negative_variance_red_cfs_unchanged(summary_wb):
    """Negative-Variance red CFs remain simple `<colLetter><row><0` formulas.

    Task 8 must not have touched these — they still fire on partial years
    because a genuine overspend against plan matters even with partial data.
    """
    ws = summary_wb["Summary"]
    rules = _red_rules(_collect_cf_rules(ws))
    # Variance red rules: range sqref, "<0" but not "0.8*" and not AND(...)
    var_red = [
        r for r in rules
        if ":" in r[0]
        and any("<0" in f for f in r[1])
        and not any("0.8*" in f for f in r[1])
        and not any(f.startswith("AND(") for f in r[1])
    ]
    # Expect: num_years per-year Var € + 1 project-total Var €
    # plus the RP-summary block (num_rps + 1). Here we only assert the
    # calendar-year block count as a lower bound — that's what Task 8 scope
    # covers.
    assert len(var_red) >= _NUM_YEARS + 1, (
        f"Expected >= {_NUM_YEARS + 1} simple <0 red variance CFs, "
        f"got {len(var_red)}"
    )

    # For each per-year Var € column, confirm the formula is still the bare
    # single-cell "<0" comparison.
    for y in range(_NUM_YEARS):
        var_col = _YEAR_GROUP_START + y * _COLS_PER_YEAR + 4
        var_col_l = col_letter(var_col)
        matches = [
            r for r in var_red
            if r[0].startswith(f"{var_col_l}")
            and any(
                f.startswith(var_col_l) and f.endswith("<0")
                for f in r[1]
            )
        ]
        assert len(matches) == 1, (
            f"Year {y} Var € red <0 CF: expected 1 unchanged rule on col "
            f"{var_col_l}, got {len(matches)}"
        )
