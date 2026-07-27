"""Tests for Task 9 — Capacity twin view (Actual FTE, Variance, Utilisation)."""

import pytest
import openpyxl

from budget_planner.config import (
    TEAM,
    EXTERNAL_COMMITMENTS,
    DEMO_CAPACITY_ASSIGNMENTS,
    MONTH_COL_START,
    PROJECT_MONTHS,
)
from budget_planner.layout import (
    SHEET_ACTUALS_HOURS,
    CAPACITY_VARIANCE_FTE_RED,
    CAPACITY_AMBER_LOAD,
    val_columns,
)
from budget_planner.styles import col_letter


# ── Fixtures ──


@pytest.fixture(scope="module")
def wb_and_cap(assigned_wb_path):
    wb = openpyxl.load_workbook(assigned_wb_path, data_only=False)
    return wb, wb["Capacity"]


@pytest.fixture(scope="module")
def cap_ws(wb_and_cap):
    return wb_and_cap[1]


@pytest.fixture(scope="module")
def band_cols():
    """Expected actual / variance band start columns."""
    vcols = val_columns(PROJECT_MONTHS)
    val_diff = vcols["diff"]
    actual_start = val_diff + 2
    actual_end = actual_start + PROJECT_MONTHS - 1
    variance_start = actual_end + 2
    variance_end = variance_start + PROJECT_MONTHS - 1
    return {
        "actual_start": actual_start,
        "actual_end": actual_end,
        "variance_start": variance_start,
        "variance_end": variance_end,
    }


def _collect_cf_rules(ws):
    results = []
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
            results.append((sqref, formulas, fg))
    return results


# ── Positional tests ──


def test_actual_band_columns_after_last_planned_month(cap_ws, band_cols):
    """The Actual band starts 2 columns after the planned Variance column
    (1-col gutter), and spans PROJECT_MONTHS columns."""
    vcols = val_columns(PROJECT_MONTHS)
    assert band_cols["actual_start"] == vcols["diff"] + 2
    assert band_cols["actual_end"] - band_cols["actual_start"] + 1 == PROJECT_MONTHS


def test_variance_band_columns_follow_actual_band(band_cols):
    """Variance band starts 2 columns after Actual band end (1-col gutter)."""
    assert band_cols["variance_start"] == band_cols["actual_end"] + 2
    assert band_cols["variance_end"] - band_cols["variance_start"] + 1 == PROJECT_MONTHS


def test_band_columns_have_outline_level_1(cap_ws, band_cols):
    """All Actual + Variance band columns are outline_level=1 so the user
    can collapse them via Excel's outline (summaryRight)."""
    for c in range(band_cols["actual_start"], band_cols["actual_end"] + 1):
        dim = cap_ws.column_dimensions[col_letter(c)]
        assert dim.outline_level == 1, f"Actual col {col_letter(c)} outline_level={dim.outline_level}"
    for c in range(band_cols["variance_start"], band_cols["variance_end"] + 1):
        dim = cap_ws.column_dimensions[col_letter(c)]
        assert dim.outline_level == 1, f"Variance col {col_letter(c)} outline_level={dim.outline_level}"


def test_capacity_sheet_summary_right_set(cap_ws):
    """Capacity outline pr must have summaryRight=True for the right-side band."""
    outline_pr = cap_ws.sheet_properties.outlinePr
    assert outline_pr is not None and outline_pr.summaryRight is True


# ── Formula content tests (rely on cap_ws + a layout probe) ──


def _find_any_wp_row(cap_ws):
    """Return (row, asgn) for the first demo assignment on the Capacity sheet."""
    # Find a row whose Col A resolves to the WP id of the first assignment.
    asgn = DEMO_CAPACITY_ASSIGNMENTS[0]
    # The planned-FTE cell in an active month must be a number for a
    # distributed assignment, so scan rows 1..200 of col A for a non-empty
    # Config-ref formula matching any WP id.
    for r in range(1, 400):
        a_val = cap_ws.cell(row=r, column=1).value
        b_val = cap_ws.cell(row=r, column=2).value
        # Header/section rows have string values; WP rows are formulas.
        if (
            isinstance(a_val, str)
            and a_val.startswith("=Config!A")
            and isinstance(b_val, str)
            and b_val.startswith("=Config!B")
        ):
            # Match by loading the target Config row — we only need ANY WP row
            # to verify the formula pattern.
            return r, asgn
    raise AssertionError("No WP row found on Capacity sheet")


def test_actual_fte_formula_references_actuals_hours(cap_ws, band_cols):
    """Actual FTE cell formula refs Actuals-Hours and divides by hours_per_fte_month."""
    wp_row, _ = _find_any_wp_row(cap_ws)
    # Find a non-grey Actual cell (active month) on this row.
    found = False
    for c in range(band_cols["actual_start"], band_cols["actual_end"] + 1):
        cell = cap_ws.cell(row=wp_row, column=c)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            formula = cell.value
            assert SHEET_ACTUALS_HOURS in formula or "Actuals-Hours" in formula, (
                f"Actual FTE formula missing sheet ref: {formula}"
            )
            assert "hours_per_fte_month" in formula, f"Actual FTE formula missing divisor: {formula}"
            found = True
            break
    assert found, f"No Actual FTE formula cell found on row {wp_row}"


def test_variance_formula_is_planned_minus_actual(cap_ws, band_cols):
    """Variance = planned - actual on the same row."""
    wp_row, _ = _find_any_wp_row(cap_ws)
    found = False
    for mi in range(PROJECT_MONTHS):
        vc = band_cols["variance_start"] + mi
        cell = cap_ws.cell(row=wp_row, column=vc)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            planned_c = MONTH_COL_START + mi
            actual_c = band_cols["actual_start"] + mi
            planned_ref = f"{col_letter(planned_c)}{wp_row}"
            actual_ref = f"{col_letter(actual_c)}{wp_row}"
            assert planned_ref in cell.value and actual_ref in cell.value, (
                f"Variance formula should contain {planned_ref} and {actual_ref}: {cell.value}"
            )
            assert f"{planned_ref}-{actual_ref}" in cell.value, (
                f"Variance formula must compute planned − actual: {cell.value}"
            )
            found = True
            break
    assert found, "No variance formula cell found"


def test_utilisation_cf_references_max_fte_on_total_row(cap_ws, band_cols):
    """Utilisation red CF on Actual TOTAL row references Max FTE for the person."""
    rules = _collect_cf_rules(cap_ws)
    a_start = band_cols["actual_start"]
    a_end = band_cols["actual_end"]
    # Find utilisation rules — sqref in Actual-band columns, formula with
    # Config!$B$<row> (the roster Max FTE cell) and without a SUM( prefix.
    util_rules = [
        r
        for r in rules
        if a_start <= _sqref_first_col_index(r[0]) <= a_end and any("Config!$B$" in f and "SUM(" not in f for f in r[1])
    ]
    assert len(util_rules) >= 2 * len(TEAM), (
        f"Expected ≥ {2 * len(TEAM)} utilisation CF rules (amber+red per person), got {len(util_rules)}"
    )
    # At least one is red, at least one is amber.
    reds = [r for r in util_rules if r[2] and r[2].endswith("FFC7CE")]
    ambers = [r for r in util_rules if r[2] and r[2].endswith("FFD966")]
    assert len(reds) >= len(TEAM), f"Expected ≥ {len(TEAM)} red util rules"
    assert len(ambers) >= len(TEAM), f"Expected ≥ {len(TEAM)} amber util rules"
    # Red formula must reference Max FTE (Config!$B$<roster_row>) and NOT use
    # the amber 0.8 multiplier.
    for _sqref, formulas, _fg in reds:
        assert any("Config!$B$" in f and f"{CAPACITY_AMBER_LOAD}*" not in f for f in formulas)


def test_variance_cf_red_on_abs_threshold(cap_ws, band_cols):
    """Variance band has red CF when ABS(variance) > CAPACITY_VARIANCE_FTE_RED."""
    rules = _collect_cf_rules(cap_ws)
    v_start = band_cols["variance_start"]
    v_end = band_cols["variance_end"]
    variance_rules = [
        r for r in rules if v_start <= _sqref_first_col_index(r[0]) <= v_end and any("ABS(" in f for f in r[1])
    ]
    assert len(variance_rules) >= 1, "No variance ABS red CF rules found"
    for _sqref, formulas, fg in variance_rules:
        assert fg and fg.endswith("FFC7CE"), f"Variance ABS rule should be red, got {fg}"
        assert any(f">{CAPACITY_VARIANCE_FTE_RED}" in f for f in formulas), (
            f"Variance formula missing >{CAPACITY_VARIANCE_FTE_RED}: {formulas}"
        )


def _sqref_first_col_index(sqref: str) -> int:
    """Parse first column letter from a CF sqref like 'AX12:BB12' → column index."""
    import re

    m = re.match(r"([A-Z]+)", sqref)
    if not m:
        return -1
    # Convert letters to 1-based column index.
    letters = m.group(1)
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


# ── Font color tests ──


def test_actual_cells_use_green_xref_font(cap_ws, band_cols):
    """Actual FTE formula cells use green cross-sheet font (006080 family)."""
    wp_row, _ = _find_any_wp_row(cap_ws)
    for c in range(band_cols["actual_start"], band_cols["actual_end"] + 1):
        cell = cap_ws.cell(row=wp_row, column=c)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            font = cell.font
            assert font.color and font.color.rgb, f"Actual cell ({wp_row},{c}) has no font color"
            rgb = str(font.color.rgb).upper()
            assert rgb.endswith("008000"), f"Actual cell font expected green (008000), got {rgb}"
            return
    pytest.fail(f"No Actual FTE formula cell found on row {wp_row}")


def test_variance_cells_use_black_formula_font(cap_ws, band_cols):
    """Variance FTE cells use black formula font (000000)."""
    wp_row, _ = _find_any_wp_row(cap_ws)
    for c in range(band_cols["variance_start"], band_cols["variance_end"] + 1):
        cell = cap_ws.cell(row=wp_row, column=c)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            font = cell.font
            assert font.color and font.color.rgb, f"Variance cell ({wp_row},{c}) has no font color"
            rgb = str(font.color.rgb).upper()
            assert rgb.endswith("000000"), f"Variance cell font expected black (000000), got {rgb}"
            return
    pytest.fail(f"No variance formula cell found on row {wp_row}")


def test_actual_total_row_cells_use_black_formula_font(cap_ws, band_cols):
    """Actual TOTAL row cells (SUM of actuals) use black formula font."""
    # Find a person's TOTAL row — col B == "TOTAL".
    total_row = None
    for r in range(1, 400):
        if cap_ws.cell(row=r, column=2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None, "No TOTAL row found"
    cell = cap_ws.cell(row=total_row, column=band_cols["actual_start"])
    assert isinstance(cell.value, str) and cell.value.startswith("=")
    rgb = str(cell.font.color.rgb).upper()
    assert rgb.endswith("000000"), f"Actual TOTAL cell expected black font, got {rgb}"


# ── Inactive month greying on band cells ──


def test_band_inactive_months_are_grey(cap_ws, band_cols):
    """Inactive months on the band must be empty + grey, mirroring planned side."""
    # Find a WP row whose assignment does NOT cover month 1.
    target = next(
        (a for a in DEMO_CAPACITY_ASSIGNMENTS if a["start"] > 1),
        None,
    )
    if target is None:
        pytest.skip("No demo assignment starts after M1 — nothing to grey-test")
    # Find that WP row on Capacity by matching the first WP row of the person.
    # Simpler: scan all wp rows and locate one whose planned M1 cell is grey
    # (proves inactive) and assert the band M1 cell is also grey.
    found = False
    for r in range(1, 400):
        a_val = cap_ws.cell(row=r, column=1).value
        b_val = cap_ws.cell(row=r, column=2).value
        if not (
            isinstance(a_val, str)
            and a_val.startswith("=Config!A")
            and isinstance(b_val, str)
            and b_val.startswith("=Config!B")
        ):
            continue
        planned_m1_fill = cap_ws.cell(row=r, column=MONTH_COL_START).fill
        planned_m1_rgb = str(planned_m1_fill.fgColor.rgb) if planned_m1_fill.fgColor else ""
        if not planned_m1_rgb.endswith("D9D9D9"):
            continue  # not an inactive-M1 row
        # Planned M1 is grey → both band M1 cells must be grey too.
        actual_m1 = cap_ws.cell(row=r, column=band_cols["actual_start"])
        var_m1 = cap_ws.cell(row=r, column=band_cols["variance_start"])
        assert str(actual_m1.fill.fgColor.rgb).endswith("D9D9D9") if actual_m1.fill.fgColor else False, (
            f"Band Actual M1 should be grey on inactive row {r}"
        )
        assert str(var_m1.fill.fgColor.rgb).endswith("D9D9D9") if var_m1.fill.fgColor else False, (
            f"Band Variance M1 should be grey on inactive row {r}"
        )
        found = True
        break
    assert found, "No WP row with inactive M1 found — greying test skipped-by-design"


# ── External rows included in variance CF ──


def test_external_rows_get_variance_cf(cap_ws, band_cols):
    """External commitment rows must also get the variance-red CF rule."""
    rules = _collect_cf_rules(cap_ws)
    v_start = band_cols["variance_start"]
    v_end = band_cols["variance_end"]
    variance_rules_count = sum(
        1 for r in rules if v_start <= _sqref_first_col_index(r[0]) <= v_end and any("ABS(" in f for f in r[1])
    )
    # One rule per WP row + one per external row.
    expected_min = len(DEMO_CAPACITY_ASSIGNMENTS) + len(EXTERNAL_COMMITMENTS)
    assert variance_rules_count >= expected_min, (
        f"Expected ≥ {expected_min} variance CF rules (WPs + externals), got {variance_rules_count}"
    )
