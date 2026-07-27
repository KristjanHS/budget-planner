"""Summary calendar-year WP block — per-column partial-year shading.

Phase 2.3 (2026-05-09) retired the per-year Forecast EUR slot. The four
per-year columns (Plan FTE, Plan €, Act €, Var €) shade grey when the
year is not fully covered by actuals:

- Act €  shades by ``eur_actuals_through_month``
- Var €  shades by the stricter MIN of EUR/hours coverage

``year_coverage_pct = months_of_year_covered / 12`` is the design metric;
the grey CF fires when the year-end project-month exceeds the through-month
value.

The previous "Fcst < 0.8 * Plan red CF, gated on full coverage" rule retired
along with the per-year Fcst slot — replaced by a single project-end
Forecast EUR column with its own (un-gated) red CF guarded by ``<>""``.
"""

import pytest
import openpyxl

from budget_planner.layout import _calendar_year_groups
from budget_planner.config import (
    PROJECT_MONTHS,
    PROJECT_START,
)
from budget_planner.styles import col_letter


_YEAR_GROUP_START = 3
_COLS_PER_YEAR = 4  # Plan FTE, Plan €, Act €, Var €
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
# Project-end Forecast EUR column sits between year band and PROJECT TOTAL.
_FORECAST_COL = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR
_TOTAL_GROUP_START = _FORECAST_COL + 1
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


# Helpers


def _grey_rules(rules):
    return [r for r in rules if r[2] is not None and r[2].endswith(_GREY_HEX)]


def _red_rules(rules):
    return [r for r in rules if r[2] is not None and r[2].endswith(_RED_HEX)]


# Grey partial-year rules present per year


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
    matches = [r for r in rules if r[0].startswith(f"{act_col_l}") and any(expected_formula == f for f in r[1])]
    assert len(matches) == 1, (
        f"Year {y} Act € grey CF: expected 1 rule with formula {expected_formula!r}, got {len(matches)}"
    )


@pytest.mark.parametrize("y", range(_NUM_YEARS))
def test_var_eur_has_grey_cf_per_year_with_min(summary_wb, y):
    """Var € column has a grey CF using the stricter MIN of both through-months.

    Phase 2.3: Var is at offset +3 (was +4 before the per-year Forecast slot
    retired).
    """
    ws = summary_wb["Summary"]
    rules = _grey_rules(_collect_cf_rules(ws))
    group_start, span = _CAL_GROUPS[y]
    year_end_pm = group_start + span
    sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
    var_col_l = col_letter(sc + 3)
    matches = [
        r
        for r in rules
        if r[0].startswith(f"{var_col_l}")
        and any(
            "MIN(eur_actuals_through_month,hours_actuals_through_month)" in f and f"<{year_end_pm}" in f for f in r[1]
        )
    ]
    assert len(matches) == 1, f"Year {y} Var € grey CF (MIN): expected 1 rule, got {len(matches)}"


# Project-end Forecast EUR red CF (replacement for the retired per-year rule)


def test_project_end_forecast_red_cf_present(summary_wb):
    """A single red CF on the project-end Forecast EUR column compares
    Forecast < 0.8 * Plan EUR project total, with an `<>""` empty-sentinel
    guard."""
    ws = summary_wb["Summary"]
    rules = _red_rules(_collect_cf_rules(ws))
    fc_l = col_letter(_FORECAST_COL)
    plan_total_l = col_letter(_TOTAL_GROUP_START + 1)
    matches = [
        r
        for r in rules
        if r[0].startswith(f"{fc_l}") and any("0.8*" in f and '<>""' in f and plan_total_l in f for f in r[1])
    ]
    assert len(matches) == 1, (
        f"Project-end Forecast red CF: expected 1 rule on col {fc_l} with "
        f"'0.8*', '<>\"\"', and Plan total ref {plan_total_l}, got "
        f"{len(matches)}. Rules: {rules}"
    )

    # Bystander: the adjacent last-year Var EUR column must NOT carry a
    # 0.8* red rule. Catches a regression where the helper accidentally
    # writes the rule to `forecast_col - 1` or duplicates onto the year band.
    adjacent_var_l = col_letter(_FORECAST_COL - 1)
    adjacent_hits = [r for r in rules if r[0].startswith(f"{adjacent_var_l}") and any("0.8*" in f for f in r[1])]
    assert len(adjacent_hits) == 0, f"0.8* red CF unexpectedly on adjacent col {adjacent_var_l}: {adjacent_hits}"


# PROJECT TOTAL block has no grey CF


def test_project_total_block_has_no_grey_cf(summary_wb):
    """PROJECT TOTAL columns (across-years) must have no grey partial-year CF."""
    ws = summary_wb["Summary"]
    rules = _grey_rules(_collect_cf_rules(ws))
    total_col_letters = {col_letter(_TOTAL_GROUP_START + i) for i in range(_COLS_PER_YEAR)}
    for sqref, formulas, _fg in rules:
        # sqref like "H5:H9" — take column prefix
        col_prefix = sqref.split(":", 1)[0].rstrip("0123456789")
        assert col_prefix not in total_col_letters, (
            f"Grey CF unexpectedly applied to PROJECT TOTAL col {col_prefix} (sqref={sqref}, formula={formulas})"
        )


# Negative-variance CFs unchanged


def test_negative_variance_red_cfs_unchanged(summary_wb):
    """Negative-Variance red CFs remain simple ``<colLetter><row><0`` formulas.

    Phase 2.3 shifted Var from offset +4 to +3 but the formula shape is
    unchanged — they still fire on partial years because a genuine overspend
    against plan matters even with partial data.
    """
    ws = summary_wb["Summary"]
    rules = _red_rules(_collect_cf_rules(ws))
    # Variance red rules: range sqref, "<0" but not "0.8*" and not AND(...)
    var_red = [
        r
        for r in rules
        if ":" in r[0]
        and any("<0" in f for f in r[1])
        and not any("0.8*" in f for f in r[1])
        and not any(f.startswith("AND(") for f in r[1])
    ]
    # Expect: num_years per-year Var € + 1 project-total Var €
    # plus the RP-summary block (num_rps + 1). Here we only assert the
    # calendar-year block count as a lower bound — that's what this scope
    # covers.
    assert len(var_red) >= _NUM_YEARS + 1, (
        f"Expected >= {_NUM_YEARS + 1} simple <0 red variance CFs, got {len(var_red)}"
    )

    # Locate the calendar-year first WP row dynamically — sqrefs from the RP
    # block share col-letters with the year-band block at offset +3 (e.g. col
    # F for year 0 also hosts the first RP's Var €), so we filter by row range
    # to disambiguate.
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("=Config!A"):
            first_wp_row = r
            break
    assert first_wp_row is not None

    # For each per-year Var € column, confirm the formula is still the bare
    # single-cell "<0" comparison at the calendar-year block's first WP row.
    for y in range(_NUM_YEARS):
        var_col = _YEAR_GROUP_START + y * _COLS_PER_YEAR + 3
        var_col_l = col_letter(var_col)
        matches = [
            r
            for r in var_red
            if r[0].startswith(f"{var_col_l}{first_wp_row}")
            and any(f.startswith(var_col_l) and f.endswith("<0") for f in r[1])
        ]
        assert len(matches) == 1, (
            f"Year {y} Var € red <0 CF: expected 1 unchanged rule on col "
            f"{var_col_l} starting row {first_wp_row}, got {len(matches)}"
        )
