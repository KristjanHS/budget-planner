"""Posted-cumulative helper / slope tests on the Actuals sheet.

Posted-cumulative redesign (2026-05-09) — design doc
``docs/plans/archived/2026-05-09-actuals-posted-cumulative-design.md``
— made each WP-row cell on Actuals hold the posted-cumulative value
through that month (gaps allowed inside the active window). Three
load-bearing helpers drive every downstream consumer:

* ``posted_idx``   — SUMPRODUCT-MAX-ISNUMBER scan returning the latest
                     1-based month index with a numeric posting at-or-
                     before ``current_month``. Banned-pattern-safe form
                     (MATCH / LOOKUP fail under openpyxl + LO recalc).
* ``posted_value`` — ``IFERROR(INDEX(rng, posted_idx), "")``: latest
                     filled cumulative through ``current_month`` or ""
                     sentinel.
* ``slope``        — ``IF(posted_idx=0, "", posted_value/posted_idx)``:
                     avg-burn (EUR per project-month). Used by Summary
                     project-end Forecast EUR + Trends reconciliation.

This file pins:
  1. Banned-pattern-safe shape on the helpers + slope (regression guard
     against MATCH/LOOKUP refactors that parse but break under recalc).
  2. ``""`` sentinel + cell-level redundancy between hidden ``posted_value``
     helper and visible ``Posted -i`` rollup.
  3. End-to-end gap handling under recalc — explicit ``[fill, blank, fill,
     blank, fill]`` pattern, asserts ``posted_idx`` = LATEST-filled (not
     first-filled, not count, not max-of-values).
"""
import json
import subprocess
from pathlib import Path

import openpyxl
import pytest

from budget_planner.config import PROJECT_MONTHS, DEMO_CAPACITY_ASSIGNMENTS, MONTH_COL_START
from budget_planner.layout import val_columns
from budget_planner.styles import col_letter
from budget_planner.workbook import init_workbook

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


@pytest.fixture(scope="module")
def actuals_ws(demo_wb_path):
    """Pre-recalc demo workbook (formulas in cells)."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    return wb["Actuals"]


@pytest.fixture(scope="module")
def vcols():
    return val_columns(PROJECT_MONTHS)


def _find_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r + 1
    raise AssertionError("No WP row found in Actuals")


# ── Helper-cell formula shape (banned-pattern protection) ──


def test_posted_idx_helper_uses_sumproduct_max_isnumber(actuals_ws, vcols):
    """``posted_idx`` MUST use SUMPRODUCT(MAX(...ISNUMBER(...))) array form
    — MATCH and LOOKUP variants parse but silently break under
    openpyxl + LO recalc (see ``.claude/rules/openpyxl.md`` § Core
    patterns). This is the specific regression class the test guards.
    """
    row = _find_first_wp_row(actuals_ws)
    formula = str(actuals_ws.cell(row, vcols["posted_idx"]).value)
    assert formula.startswith("=SUMPRODUCT(MAX("), formula
    assert "ISNUMBER(" in formula, formula
    assert "current_month" in formula, formula
    # Banned forms: any of these would parse but fail under recalc.
    assert "MATCH(" not in formula, formula
    assert "LOOKUP(" not in formula, formula


def test_posted_value_helper_indexes_grid_at_posted_idx(actuals_ws, vcols):
    """Hidden ``posted_value`` helper is ``IFERROR(INDEX(<wp_range>,
    <posted_idx_cell>), "")``. The IFERROR catches ``INDEX(rng, 0)`` when
    no postings; the ``""`` sentinel is the design-§2 "no data yet"
    marker that downstream consumers treat as blank.
    """
    row = _find_first_wp_row(actuals_ws)
    formula = str(actuals_ws.cell(row, vcols["posted_value"]).value)
    idx_l = col_letter(vcols["posted_idx"])
    assert formula.startswith("=IFERROR(INDEX("), formula
    assert f"{idx_l}{row}" in formula, formula
    assert formula.endswith(',"")'), formula


def test_visible_posted_mirrors_helper_formula(actuals_ws, vcols):
    """Visible ``Posted -i`` cell repeats the helper's INDEX formula
    verbatim (intentional cell-level redundancy per design §2:
    'click cell, see calc' UX). If the visible cell ever drifts to
    just ``=<helper_cell>``, the user loses formula provenance.
    """
    row = _find_first_wp_row(actuals_ws)
    helper = str(actuals_ws.cell(row, vcols["posted_value"]).value)
    visible = str(actuals_ws.cell(row, vcols["posted"]).value)
    assert visible == helper, (
        f"Posted -i should mirror posted_value helper verbatim.\n"
        f"  helper:  {helper}\n"
        f"  visible: {visible}"
    )


def test_slope_formula_divides_posted_value_by_posted_idx(actuals_ws, vcols):
    """``Slope -i`` = ``IF(posted_idx=0, "", posted_value/posted_idx)``.
    The /0 guard surfaces blank (not ``#DIV/0!``) when no postings —
    downstream Summary Forecast EUR + Trends reconciliation depend on
    that sentinel passing through cleanly. test_actuals.py:88 has a
    shallower version; this pins the literal division form.
    """
    row = _find_first_wp_row(actuals_ws)
    formula = str(actuals_ws.cell(row, vcols["slope"]).value)
    idx_l = col_letter(vcols["posted_idx"])
    val_l = col_letter(vcols["posted_value"])
    assert formula == (
        f'=IF({idx_l}{row}=0,"",{val_l}{row}/{idx_l}{row})'
    ), formula


# ── Mid-window gap handling under recalc (load-bearing) ──


@pytest.fixture(scope="module")
def gap_recalc_path(tmp_path_factory):
    """Workbook seeded with an explicit ``[fill, blank, fill, blank, fill]``
    pattern at M1/M3/M5 on the first WP row, then recalc'd. Pins that
    the SUMPRODUCT-MAX scan returns the LATEST filled month — not
    first-filled, count-of-filled, or max-of-values.

    The values 200/380/500 are arbitrary integers; only the *positions*
    (M1, M3, M5) are load-bearing for the latest-scan test. The
    ``slope=500/5=100`` assertion uses round numbers so a denominator
    or operand-swap regression yields an obviously wrong answer (e.g.
    swapped: 5/500=0.01), not to assert posted-cumulative value
    semantics on the seeded numbers.
    """
    out = tmp_path_factory.mktemp("posted_gap_wb") / "gap.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(out)
    ws = wb["Actuals"]
    first_wp_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break
    assert first_wp_row is not None, "No WP ID header row found"

    # Clear any pre-existing values on the WP row across the active range,
    # then write the gap pattern at months 1, 3, 5.
    for m in range(PROJECT_MONTHS):
        ws.cell(row=first_wp_row, column=MONTH_COL_START + m, value=None)
    ws.cell(row=first_wp_row, column=MONTH_COL_START + 0, value=200)    # M1
    ws.cell(row=first_wp_row, column=MONTH_COL_START + 2, value=380)    # M3
    ws.cell(row=first_wp_row, column=MONTH_COL_START + 4, value=500)    # M5

    # Override BOTH through-month controls so the SUMPRODUCT scan filter
    # (`<= current_month`) is deterministic. The ``current_month`` named
    # range resolves to ``Trends!$B$1`` (workbook.py Stage 5 — Trends
    # stub). Without this, ``current_month`` is date-derived from
    # ``project_start`` + today() and would shift over time, making the
    # ``posted_idx=5`` assertion silently brittle. The Actuals!B1
    # override drives ``eur_actuals_through_month`` for sibling EUR
    # consumers; not strictly required here but kept for clarity since
    # both control strips serve through-month overrides.
    wb["Trends"].cell(row=1, column=2, value=10)
    ws.cell(row=1, column=2, value=10)

    wb.save(str(out))
    wb.close()

    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), str(out)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip(f"LibreOffice recalc unavailable: {result.stderr}")
    json.loads(result.stdout)  # validates recalc output shape
    return out, first_wp_row


def test_gap_posted_idx_is_latest_filled(gap_recalc_path, vcols):
    """``posted_idx`` returns 5 (the M5 fill) — NOT 1 (first fill), NOT 3
    (count of fills), NOT 500 (max value)."""
    out, row = gap_recalc_path
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["Actuals"]
    assert ws.cell(row, vcols["posted_idx"]).value == 5


def test_gap_posted_value_is_latest_filled_value(gap_recalc_path, vcols):
    """``posted_value`` returns 500 (the M5 cell) — guards against an
    INDEX-into-wrong-row regression (would pull from another WP)."""
    out, row = gap_recalc_path
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["Actuals"]
    assert ws.cell(row, vcols["posted_value"]).value == 500


def test_gap_slope_is_posted_value_over_idx(gap_recalc_path, vcols):
    """``slope`` = 500 / 5 = 100 — pins both the /0 guard non-firing
    (idx > 0) AND the division order (regressions that swapped the
    operands would yield 5/500 = 0.01 instead)."""
    out, row = gap_recalc_path
    wb = openpyxl.load_workbook(out, data_only=True)
    ws = wb["Actuals"]
    assert ws.cell(row, vcols["slope"]).value == 100
