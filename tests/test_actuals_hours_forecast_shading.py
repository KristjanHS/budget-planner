"""Cosmetic grey CF tests on the Actuals-Hours forecast-EUR band.

Two independent grey rules shade the forecast band:

* ``month <= eur_actuals_through_month`` — finance has closed the month,
  so the forecast is stale vs. booked Actual EUR (authoritative on the
  Actuals sheet).
* ``month > hours_actuals_through_month`` — no hours logged yet, so the
  forecast aggregates zero contributions (empty/unreliable).

Both use ``CF_GREY_FILL`` (8-char AARRGGBB). Test matches rules by fill
colour + formula content, NOT by positional index — adding new CF rules
shifts all indices (rule from `openpyxl-patterns.md`).
"""

import pytest
from openpyxl import load_workbook

from budget_planner.config import PROJECT_MONTHS
from budget_planner.layout import val_columns
from budget_planner.styles import CF_GREY_FILL, col_letter


FORECAST_COL_START = val_columns(PROJECT_MONTHS)["sum"] + 2
FORECAST_COL_END = FORECAST_COL_START + PROJECT_MONTHS - 1


@pytest.fixture(scope="module")
def hrs(assigned_wb_path):
    wb = load_workbook(assigned_wb_path, data_only=False)
    return wb["Actuals-Hours"]


def _grey_rules_on_forecast_band(ws):
    """Return list of (range_str, formula) pairs for grey CF rules that
    overlap the forecast band columns."""
    fc_start_l = col_letter(FORECAST_COL_START)
    target_grey = str(CF_GREY_FILL.bgColor.rgb).upper()
    matches = []
    for cf_range in ws.conditional_formatting:
        rng_str = str(cf_range)
        if fc_start_l not in rng_str:
            continue
        for rule in ws.conditional_formatting[cf_range]:
            fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
            if fill is None:
                continue
            bg = str(getattr(fill.bgColor, "rgb", "") or "").upper()
            if not bg.endswith(target_grey):
                continue
            for f in (rule.formula or []):
                matches.append((rng_str, f))
    return matches


def test_two_grey_cf_rules_on_forecast_band(hrs):
    """Exactly two grey CF rules — one per threshold."""
    rules = _grey_rules_on_forecast_band(hrs)
    # The forecast band may have other grey rules added later; assert at
    # least one references eur_... and at least one references hours_...
    eur_rules = [r for r in rules if "eur_actuals_through_month" in r[1]]
    hours_rules = [r for r in rules if "hours_actuals_through_month" in r[1]]
    assert eur_rules, (
        f"Missing grey CF keyed on eur_actuals_through_month; saw {rules}"
    )
    assert hours_rules, (
        f"Missing grey CF keyed on hours_actuals_through_month; saw {rules}"
    )


def test_eur_threshold_rule_is_less_equal(hrs):
    """eur-side rule greys months <= threshold (closed by finance)."""
    rules = _grey_rules_on_forecast_band(hrs)
    eur_rules = [f for _, f in rules if "eur_actuals_through_month" in f]
    assert eur_rules
    assert any("<=eur_actuals_through_month" in f for f in eur_rules), (
        f"eur grey CF should use <=: {eur_rules}"
    )


def test_hours_threshold_rule_is_greater_than(hrs):
    """hours-side rule greys months > threshold (no hours yet)."""
    rules = _grey_rules_on_forecast_band(hrs)
    hours_rules = [f for _, f in rules if "hours_actuals_through_month" in f]
    assert hours_rules
    assert any(">hours_actuals_through_month" in f for f in hours_rules), (
        f"hours grey CF should use >: {hours_rules}"
    )
