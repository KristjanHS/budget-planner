"""Tests for the Trends reconciliation section (v2 Task 12)."""
from typing import TYPE_CHECKING

import pytest
import openpyxl

if TYPE_CHECKING:
    pass

from budget_planner.config import NUM_WPS


# ── Fixtures ──

@pytest.fixture(scope="module")
def recon_workbook(demo_wb_path):
    return openpyxl.load_workbook(demo_wb_path, data_only=False)


@pytest.fixture(scope="module")
def recon_ws(recon_workbook):
    return recon_workbook["Trends"]


def _find_row_containing(ws, col, substring):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v is not None and substring in str(v):
            return r
    return None


def _recon_section_row(ws):
    r = _find_row_containing(ws, 1, "RECONCILIATION: FORECAST vs ACTUAL EUR")
    assert r is not None, "RECONCILIATION section header not found"
    return r


# ================================================================
# Section header + placement
# ================================================================

def test_reconciliation_section_header(recon_ws):
    """Recon section header exists and sits between chart and Person Load."""
    recon_row = _recon_section_row(recon_ws)
    person_row = _find_row_containing(recon_ws, 1, "PERSON LOAD")
    assert person_row is not None
    assert recon_row < person_row, (
        f"Recon header row {recon_row} must precede Person Load row {person_row}"
    )


# ================================================================
# Three health-cell rows
# ================================================================

def test_health_row_labels_and_formulas(recon_ws):
    """Three health rows: labels in col A, formulas in col B."""
    recon_row = _recon_section_row(recon_ws)

    # +2: Actuals EUR through
    assert recon_ws.cell(recon_row + 2, 1).value == "Actuals EUR through"
    assert recon_ws.cell(recon_row + 2, 2).value == "=eur_actuals_through_month"

    # +3: Hours through
    assert recon_ws.cell(recon_row + 3, 1).value == "Hours through"
    assert recon_ws.cell(recon_row + 3, 2).value == "=hours_actuals_through_month"

    # +4: Reconciled through — inlined MIN() local helper (derived name dropped 2026-04-18)
    assert recon_ws.cell(recon_row + 4, 1).value == "Reconciled through"
    assert recon_ws.cell(recon_row + 4, 2).value == (
        "=MIN(eur_actuals_through_month,hours_actuals_through_month)"
    )


def test_health_cell_fonts(recon_ws):
    """Green for eur / hours (cross-sheet named refs), black for reconciled."""
    recon_row = _recon_section_row(recon_ws)

    def _rgb(cell):
        return str(getattr(cell.font.color, "rgb", "")).upper()

    # Green for the two through-month inputs (cross-sheet via named range)
    assert _rgb(recon_ws.cell(recon_row + 2, 2)).endswith("008000"), (
        f"Actuals EUR through value should be GREEN_XREF (008000), "
        f"got {_rgb(recon_ws.cell(recon_row + 2, 2))}"
    )
    assert _rgb(recon_ws.cell(recon_row + 3, 2)).endswith("008000"), (
        f"Hours through value should be GREEN_XREF (008000), "
        f"got {_rgb(recon_ws.cell(recon_row + 3, 2))}"
    )
    # Black for the derived reconciled_through_month
    reconciled_rgb = _rgb(recon_ws.cell(recon_row + 4, 2))
    assert reconciled_rgb.endswith("000000") or reconciled_rgb == "00000000", (
        f"Reconciled through value should be BLACK_FORMULA, got {reconciled_rgb}"
    )


def test_health_cell_integer_format(recon_ws):
    """All three health value cells use fmt='0' (integer month index)."""
    recon_row = _recon_section_row(recon_ws)
    for offset in (2, 3, 4):
        cell = recon_ws.cell(recon_row + offset, 2)
        assert cell.number_format == "0", (
            f"Health row offset {offset}: expected fmt '0', got {cell.number_format!r}"
        )


# ================================================================
# Per-WP summary table
# ================================================================

def test_recon_table_headers(recon_ws):
    """Table header at +6 has WP / Cum Forecast EUR / Cum Actual EUR / Gap EUR / Gap %."""
    recon_row = _recon_section_row(recon_ws)
    header_row = recon_row + 6
    expected = ["WP", "Cum Forecast EUR", "Cum Actual EUR", "Gap EUR", "Gap %"]
    for i, h in enumerate(expected, 1):
        got = recon_ws.cell(header_row, i).value
        assert got == h, f"Col {i} header: expected {h!r}, got {got!r}"


def test_recon_wp_row_count_plus_total(recon_ws):
    """NUM_WPS data rows + one TOTAL row."""
    recon_row = _recon_section_row(recon_ws)
    first_wp_row = recon_row + 7

    # Data rows: col A must be a formula referencing Config!A
    for i in range(NUM_WPS):
        val = recon_ws.cell(first_wp_row + i, 1).value
        assert isinstance(val, str) and val.startswith("=Config!A"), (
            f"Recon WP row {i} col A should reference Config!A, got {val!r}"
        )

    total_row = first_wp_row + NUM_WPS
    assert recon_ws.cell(total_row, 1).value == "TOTAL", (
        f"Expected TOTAL label at row {total_row}, "
        f"got {recon_ws.cell(total_row, 1).value!r}"
    )


def test_recon_col_b_forecast_formula(recon_ws):
    """Col B: per-WP slope projection on Actuals, guarded by the local
    ``Reconciled through`` helper cell ($B$<recon_row+4>) and by
    posted_idx=0 -> 0. Linear extrapolation from the WP's posted-cumulative
    anchor out to reconciled_through; MAX(0, ...) clamps a negative
    projection.
    """
    recon_row = _recon_section_row(recon_ws)
    reconciled_ref = f"$B${recon_row + 4}"
    first_wp_row = recon_row + 7
    formula = str(recon_ws.cell(first_wp_row, 2).value)
    # Defensive: Col B must not reference Plan-Forecast.
    assert "Plan-Forecast" not in formula, (
        f"Col B must not reference Plan-Forecast: {formula!r}"
    )
    assert "Actuals!" in formula, (
        f"Col B must reference Actuals! (per-WP posted/slope cells): {formula!r}"
    )
    assert reconciled_ref in formula, (
        f"Col B must use {reconciled_ref} (local Reconciled helper) in the "
        f"slope projection: {formula!r}"
    )
    # Slope-projection shape: posted_idx=0 sentinel + MAX(0, ...) clamp.
    assert "=0" in formula, (
        f"Col B must guard posted_idx=0 with the empty-projection branch: "
        f"{formula!r}"
    )
    assert "MAX(0," in formula, (
        f"Col B must clamp projection at 0 via MAX(0, ...): {formula!r}"
    )


def test_recon_col_c_actual_formula(recon_ws):
    """Col C: INDEX-window SUM on Actuals, guarded by the local
    ``Reconciled through`` helper cell ($B$<recon_row+4>)."""
    recon_row = _recon_section_row(recon_ws)
    reconciled_ref = f"$B${recon_row + 4}"
    first_wp_row = recon_row + 7
    formula = str(recon_ws.cell(first_wp_row, 3).value)
    assert "Actuals!" in formula, f"Col C must reference Actuals!: {formula!r}"
    assert reconciled_ref in formula, (
        f"Col C must guard on {reconciled_ref} (local Reconciled helper): {formula!r}"
    )
    assert "INDEX" in formula
    assert "SUM" in formula


def test_recon_col_d_gap_formula(recon_ws):
    """Col D = C - B on same row."""
    recon_row = _recon_section_row(recon_ws)
    first_wp_row = recon_row + 7
    formula = str(recon_ws.cell(first_wp_row, 4).value)
    assert formula == f"=C{first_wp_row}-B{first_wp_row}", (
        f"Col D expected =C{first_wp_row}-B{first_wp_row}, got {formula!r}"
    )


def test_recon_col_e_gap_pct_formula(recon_ws):
    """Col E = IFERROR(D/B, 0)."""
    recon_row = _recon_section_row(recon_ws)
    first_wp_row = recon_row + 7
    formula = str(recon_ws.cell(first_wp_row, 5).value)
    assert "IFERROR" in formula
    assert f"D{first_wp_row}/B{first_wp_row}" in formula, (
        f"Col E must divide D/B on same row: {formula!r}"
    )


def test_recon_total_row_formulas(recon_ws):
    """TOTAL row: B, C are SUMs; D = C-B; E = IFERROR(D/B, 0)."""
    recon_row = _recon_section_row(recon_ws)
    first_wp_row = recon_row + 7
    last_wp_row = first_wp_row + NUM_WPS - 1
    total_row = first_wp_row + NUM_WPS

    b_formula = str(recon_ws.cell(total_row, 2).value)
    c_formula = str(recon_ws.cell(total_row, 3).value)
    d_formula = str(recon_ws.cell(total_row, 4).value)
    e_formula = str(recon_ws.cell(total_row, 5).value)

    assert b_formula == f"=SUM(B${first_wp_row}:B${last_wp_row})", (
        f"TOTAL col B expected SUM over WP data rows, got {b_formula!r}"
    )
    assert c_formula == f"=SUM(C${first_wp_row}:C${last_wp_row})", (
        f"TOTAL col C expected SUM over WP data rows, got {c_formula!r}"
    )
    assert d_formula == f"=C{total_row}-B{total_row}", (
        f"TOTAL col D expected =C-B, got {d_formula!r}"
    )
    assert "IFERROR" in e_formula and f"D{total_row}/B{total_row}" in e_formula, (
        f"TOTAL col E expected IFERROR(D/B, 0), got {e_formula!r}"
    )


# ================================================================
# Conditional formatting on Gap % (col E)
# ================================================================

def test_recon_gap_pct_cf_red_before_amber(recon_ws):
    """Col E across data rows + TOTAL row: red (ABS > 0.10) priority above amber (ABS > 0.05)."""
    recon_row = _recon_section_row(recon_ws)
    first_wp_row = recon_row + 7
    total_row = first_wp_row + NUM_WPS
    target_range = f"E{first_wp_row}:E{total_row}"

    rules = []
    for cf_range in recon_ws.conditional_formatting:
        if target_range not in str(cf_range):
            continue
        for rule in recon_ws.conditional_formatting[cf_range]:
            rules.append(rule)

    red_rules = [
        r for r in rules
        if r.formula and "ABS" in r.formula[0] and ">0.1" in r.formula[0]
    ]
    amber_rules = [
        r for r in rules
        if r.formula and "ABS" in r.formula[0] and ">0.05" in r.formula[0]
    ]

    assert len(red_rules) == 1, (
        f"Expected one red (ABS > 0.10) rule on {target_range}, "
        f"got rules={[r.formula for r in rules]}"
    )
    assert len(amber_rules) == 1, (
        f"Expected one amber (ABS > 0.05) rule on {target_range}, "
        f"got rules={[r.formula for r in rules]}"
    )

    red_priority = red_rules[0].priority
    amber_priority = amber_rules[0].priority
    assert red_priority is not None and amber_priority is not None
    assert red_priority < amber_priority, (
        f"Red priority must be lower (higher precedence) than amber. "
        f"red={red_priority} amber={amber_priority}"
    )


# ================================================================
# Gap-over-time chart (Task 13)
# ================================================================

def _gap_chart_anchor_row(ws):
    """Return the expected gap-chart anchor row (1-indexed).

    = recon total row + 2. Derived from the workbook rather than
    hardcoded so seed-data changes don't invalidate the test.
    """
    recon_row = _recon_section_row(ws)
    total_row = recon_row + 7 + NUM_WPS
    return total_row + 2


def test_gap_chart_exists_at_reconciliation_anchor(recon_ws):
    """A second LineChart is anchored at the reconciliation gap anchor row.

    openpyxl anchors use 0-indexed rows; the gap anchor row is 1-indexed.
    """
    charts = recon_ws._charts  # type: ignore[attr-defined]
    assert len(charts) >= 2, (
        f"Expected at least 2 charts on Trends, got {len(charts)}"
    )
    expected_row_0idx = _gap_chart_anchor_row(recon_ws) - 1
    gap_chart = charts[1]
    actual_row_0idx = gap_chart.anchor._from.row
    assert actual_row_0idx == expected_row_0idx, (
        f"Gap chart anchor row: expected {expected_row_0idx} (0-indexed), "
        f"got {actual_row_0idx}"
    )


def test_gap_chart_series_count(recon_ws):
    """Gap chart has exactly 2 series: Gap + Zero reference."""
    charts = recon_ws._charts  # type: ignore[attr-defined]
    gap_chart = charts[1]
    assert len(gap_chart.series) == 2, (
        f"Expected 2 series (gap + zero), got {len(gap_chart.series)}"
    )


def test_gap_chart_title_mentions_gap(recon_ws):
    """Gap chart title contains 'Gap'."""
    charts = recon_ws._charts  # type: ignore[attr-defined]
    gap_chart = charts[1]
    # chart.title can be Title object or string; extract text robustly.
    title = gap_chart.title
    if title is None:
        title_text = ""
    elif isinstance(title, str):
        title_text = title
    else:
        # Title with rich-text runs
        try:
            title_text = " ".join(
                run.t for p in title.tx.rich.p for run in p.r  # type: ignore[attr-defined]
            )
        except AttributeError:
            title_text = str(title)
    assert "Gap" in title_text, (
        f"Gap chart title should mention 'Gap', got {title_text!r}"
    )


def test_gap_chart_has_zero_reference_line(recon_ws):
    """Second series on the gap chart is the dashed zero reference."""
    charts = recon_ws._charts  # type: ignore[attr-defined]
    gap_chart = charts[1]
    zero_series = gap_chart.series[1]
    line = zero_series.graphicalProperties.line
    dash = line.dashStyle or getattr(line, "prstDash", None)
    assert dash == "dash", (
        f"Expected dashed line on zero-reference series, got dashStyle={dash!r}"
    )


def _find_hidden_row_with_first_month_formula(ws, predicate):
    """Return first hidden row whose col C (first month) formula matches."""
    from budget_planner.layout import MONTH_COL_START
    for r in range(1, ws.max_row + 1):
        if not ws.row_dimensions[r].hidden:
            continue
        v = ws.cell(r, MONTH_COL_START).value
        if v is not None and predicate(str(v)):
            return r
    return None


def test_gap_helper_row_formula(recon_ws):
    """Gap helper row's first-month formula is =<cl><actual>-<cl><forecast>.

    Discovers the helper rows dynamically via openpyxl — matches the
    pattern used elsewhere in the trends test suite. The forecast-row
    probe matches the 2-anchor (project_months, eur_actuals_through_month,
    NA()) shape.
    """
    from budget_planner.layout import MONTH_COL_START
    from budget_planner.styles import col_letter

    # Find the actual helper row (the hidden row with an Actuals! ref in
    # its first-month cell). Disambiguate from the new 2-anchor forecast
    # row, which ALSO references Actuals + eur_actuals_through_month, by
    # requiring the cumulative shape (no NA() / project_months on the
    # actual row).
    actual_row = _find_hidden_row_with_first_month_formula(
        recon_ws,
        lambda f: (
            "Actuals!" in f
            and "eur_actuals_through_month" in f
            and "project_months" not in f
            and "NA()" in f  # actual row clips with NA(); kept post Phase 2.4
        ),
    )
    forecast_row = _find_hidden_row_with_first_month_formula(
        recon_ws,
        lambda f: (
            "project_months" in f
            and "eur_actuals_through_month" in f
            and "NA()" in f
            and "Actuals!" in f
        ),
    )
    assert actual_row is not None, "Could not locate actual helper row"
    assert forecast_row is not None, (
        "Could not locate forecast helper row (2-anchor pattern)"
    )

    cl = col_letter(MONTH_COL_START)
    expected = f"={cl}{actual_row}-{cl}{forecast_row}"

    gap_row = _find_hidden_row_with_first_month_formula(
        recon_ws, lambda f: f == expected
    )
    assert gap_row is not None, (
        f"Could not find gap helper row with formula {expected!r}"
    )


def test_gap_helper_row_hidden(recon_ws):
    """Both gap helper rows (gap + zero) are hidden. Asserted together so
    the pair-check is localised."""
    # Discover the gap row by its identifying formula structure.
    gap_row = _find_hidden_row_with_first_month_formula(
        recon_ws,
        lambda f: f.startswith("=C") and "-C" in f and "Actuals" not in f,
    )
    assert gap_row is not None, "Gap helper row not found"
    assert recon_ws.row_dimensions[gap_row].hidden is True


def test_zero_helper_row_hidden(recon_ws):
    """Zero reference helper row contains =0 formula and is hidden."""
    zero_row = _find_hidden_row_with_first_month_formula(
        recon_ws, lambda f: f == "=0"
    )
    assert zero_row is not None, (
        "Zero reference helper row (formula '=0') not found"
    )
    assert recon_ws.row_dimensions[zero_row].hidden is True
