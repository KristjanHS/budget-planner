"""Tests for the FORECAST MONITORING block on the Actuals sheet (Task 6).

The forecast block parallels the Actual-EUR validation columns, sourcing
data from ``Plan-Forecast`` so reviewers can compare "what hours-log
forecast implies" against "what finance has booked" on the same row.

Columns (per WP row, in order after the Breach-In column):
    1. Cumulative Forecast EUR (SUMPRODUCT from Plan-Forecast)
    2. Forecast Variance vs Ceiling
    3. Forecast % Ceiling
    4. Forecast Breach In

All four columns share ``outline_level = 1`` so the band is collapsible.
CF rules mirror the Actual-EUR shape (red on >100% / variance <0 /
breach <6, amber on >=85% / breach <12).
"""
from __future__ import annotations

import pytest
from openpyxl import load_workbook

from budget_planner.workbook import init_workbook
from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS, WORK_PACKAGES
from budget_planner.layout import (
    SHEET_ACTUALS,
    forecast_val_columns,
    val_columns,
)
from budget_planner.config import PROJECT_MONTHS
from budget_planner.styles import col_letter


# ── Fixtures ──

@pytest.fixture(scope="module")
def workbook(tmp_path_factory):
    out = tmp_path_factory.mktemp("actuals_forecast") / "wb.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return load_workbook(str(out), data_only=False)


@pytest.fixture(scope="module")
def actuals_ws(workbook):
    return workbook[SHEET_ACTUALS]


# ── Helpers ──

def _find_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r + 1
    raise AssertionError("No 'WP ID' header row on Actuals")


def _collect_cf_rules(ws):
    out = []
    for cf in ws.conditional_formatting:
        sqref = str(cf.sqref)
        for rule in cf.rules:
            formulas = list(rule.formula) if rule.formula else []
            fg = None
            if rule.dxf and rule.dxf.fill:
                fill = rule.dxf.fill
                if fill.bgColor and str(fill.bgColor.rgb) != "00000000":
                    fg = str(fill.bgColor.rgb)
                elif fill.fgColor and str(fill.fgColor.rgb) != "00000000":
                    fg = str(fill.fgColor.rgb)
            out.append((sqref, formulas, fg))
    return out


# ── Shape ──

def test_forecast_block_columns_are_after_actual_breach(actuals_ws):
    """Forecast columns live strictly after the Actual-EUR breach column."""
    breach = val_columns(PROJECT_MONTHS)["breach"]
    fcols = forecast_val_columns(PROJECT_MONTHS)
    assert fcols["cum_forecast"] > breach
    assert fcols["var_forecast"] == fcols["cum_forecast"] + 1
    assert fcols["pct_ceil_forecast"] == fcols["cum_forecast"] + 2
    assert fcols["breach_forecast"] == fcols["cum_forecast"] + 3


def test_forecast_block_header_exists(actuals_ws):
    """Row above the WP block contains 'FORECAST MONITORING' over the block."""
    fcols = forecast_val_columns(PROJECT_MONTHS)
    # Header lives on row 1 (section header row); scan rows 1-3 defensively
    for r in range(1, 4):
        for c in (
            fcols["cum_forecast"],
            fcols["var_forecast"],
            fcols["pct_ceil_forecast"],
            fcols["breach_forecast"],
        ):
            v = actuals_ws.cell(r, c).value
            if isinstance(v, str) and "FORECAST MONITORING" in v:
                return
    raise AssertionError("FORECAST MONITORING header not found on Actuals")


def test_forecast_block_outline_level(actuals_ws):
    """All 4 forecast columns have outline_level=1 (collapsible band)."""
    fcols = forecast_val_columns(PROJECT_MONTHS)
    for key in (
        "cum_forecast", "var_forecast", "pct_ceil_forecast", "breach_forecast",
    ):
        letter = col_letter(fcols[key])
        dim = actuals_ws.column_dimensions[letter]
        assert dim.outline_level == 1, (
            f"Column {letter} ({key}) missing outline_level=1 "
            f"(got {dim.outline_level})"
        )


def test_forecast_block_summary_right(actuals_ws):
    """Outline summary sits to the right of the grouped columns."""
    outline_pr = actuals_ws.sheet_properties.outlinePr
    assert outline_pr is not None
    assert outline_pr.summaryRight is True


# ── Formula content ──

def test_cum_forecast_formula_structure(actuals_ws):
    """First WP row cum-forecast cell is a SUMPRODUCT against Plan-Forecast."""
    first = _find_first_wp_row(actuals_ws)
    fcols = forecast_val_columns(PROJECT_MONTHS)
    formula = actuals_ws.cell(first, fcols["cum_forecast"]).value
    assert isinstance(formula, str), formula
    assert formula.startswith("=SUMPRODUCT"), formula
    assert "'Plan-Forecast'!" in formula, formula
    assert "current_month" in formula, formula


def test_forecast_variance_formula(actuals_ws):
    """Variance = Ceiling cell - CumForecast cell."""
    first = _find_first_wp_row(actuals_ws)
    fcols = forecast_val_columns(PROJECT_MONTHS)
    vcols = val_columns(PROJECT_MONTHS)
    formula = actuals_ws.cell(first, fcols["var_forecast"]).value
    assert isinstance(formula, str), formula
    ceil_l = col_letter(vcols["budget"])
    cum_l = col_letter(fcols["cum_forecast"])
    assert formula == f"={ceil_l}{first}-{cum_l}{first}", formula


def test_forecast_pct_ceiling_formula(actuals_ws):
    """% Ceiling = CumForecast / Ceiling."""
    first = _find_first_wp_row(actuals_ws)
    fcols = forecast_val_columns(PROJECT_MONTHS)
    vcols = val_columns(PROJECT_MONTHS)
    formula = actuals_ws.cell(first, fcols["pct_ceil_forecast"]).value
    assert isinstance(formula, str), formula
    ceil_l = col_letter(vcols["budget"])
    cum_l = col_letter(fcols["cum_forecast"])
    assert formula == f"={cum_l}{first}/{ceil_l}{first}", formula


def test_forecast_breach_formula(actuals_ws):
    """Breach In uses the same IF shape as Actual Breach, sourced from CumForecast."""
    first = _find_first_wp_row(actuals_ws)
    fcols = forecast_val_columns(PROJECT_MONTHS)
    formula = actuals_ws.cell(first, fcols["breach_forecast"]).value
    assert isinstance(formula, str), formula
    assert formula.startswith("=IF("), formula
    assert "current_month" in formula, formula


def test_forecast_block_spans_every_wp_row(actuals_ws):
    """Every WP row has non-empty forecast formulas (not just the first)."""
    first = _find_first_wp_row(actuals_ws)
    fcols = forecast_val_columns(PROJECT_MONTHS)
    for i in range(len(WORK_PACKAGES)):
        row = first + i
        for key in (
            "cum_forecast", "var_forecast", "pct_ceil_forecast",
            "breach_forecast",
        ):
            v = actuals_ws.cell(row, fcols[key]).value
            assert v is not None and str(v).startswith("="), (
                f"Row {row} col {key} missing formula: {v!r}"
            )


# ── Conditional formatting ──

def test_forecast_pct_ceil_cf_has_red_and_amber(actuals_ws):
    """Forecast %Ceiling column has at least 1 red (>1) and 1 amber (>=0.85) rule."""
    fcols = forecast_val_columns(PROJECT_MONTHS)
    pct_col_l = col_letter(fcols["pct_ceil_forecast"])
    rules = _collect_cf_rules(actuals_ws)
    pct_rules = [r for r in rules if r[0].startswith(pct_col_l)]
    reds = [r for r in pct_rules
            if r[2] and r[2].endswith("FFC7CE")
            and any(">1" in f for f in r[1])]
    ambers = [r for r in pct_rules
              if r[2] and r[2].endswith("FFD966")
              and any(">0.85" in f for f in r[1])]
    assert len(reds) >= 1, f"Expected red CF on forecast %ceiling, got {pct_rules}"
    assert len(ambers) >= 1, (
        f"Expected amber CF on forecast %ceiling, got {pct_rules}"
    )


def test_forecast_breach_cf_has_red_and_amber(actuals_ws):
    """Forecast Breach In column has at least 1 red (<6) and 1 amber (<12) rule."""
    fcols = forecast_val_columns(PROJECT_MONTHS)
    breach_col_l = col_letter(fcols["breach_forecast"])
    rules = _collect_cf_rules(actuals_ws)
    breach_rules = [r for r in rules if r[0].startswith(breach_col_l)]
    reds = [r for r in breach_rules
            if r[2] and r[2].endswith("FFC7CE")
            and any("<6" in f for f in r[1])]
    ambers = [r for r in breach_rules
              if r[2] and r[2].endswith("FFD966")
              and any("<12" in f for f in r[1])]
    assert len(reds) >= 1, f"Expected red CF on forecast breach, got {breach_rules}"
    assert len(ambers) >= 1, (
        f"Expected amber CF on forecast breach, got {breach_rules}"
    )


def test_forecast_variance_cf_has_red(actuals_ws):
    """Forecast Variance column has at least 1 red (<0) rule."""
    fcols = forecast_val_columns(PROJECT_MONTHS)
    var_col_l = col_letter(fcols["var_forecast"])
    rules = _collect_cf_rules(actuals_ws)
    var_rules = [r for r in rules if r[0].startswith(var_col_l)]
    reds = [r for r in var_rules
            if r[2] and r[2].endswith("FFC7CE")
            and any("<0" in f for f in r[1])]
    assert len(reds) >= 1, f"Expected red CF on forecast variance, got {var_rules}"


# ── T8 reconciliation visual ──

def test_actuals_eur_actuals_through_month_cf_exists(actuals_ws):
    """T8: monthly Actual EUR cells have a grey CF gated by
    `eur_actuals_through_month` (greys months past the user's threshold)."""
    rules = _collect_cf_rules(actuals_ws)
    matches = [
        r for r in rules
        if any("eur_actuals_through_month" in f for f in r[1])
    ]
    assert len(matches) >= 1, (
        "Expected at least one CF rule on Actuals referencing "
        f"eur_actuals_through_month, got: {rules}"
    )
    # At least one should be a grey fill (D9D9D9 — see CF_GREY_FILL).
    greys = [r for r in matches if r[2] and r[2].upper().endswith("D9D9D9")]
    assert len(greys) >= 1, (
        f"Expected grey-fill CF rule on Actuals, got: {matches}"
    )
