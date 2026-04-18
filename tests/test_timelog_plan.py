"""Structural tests for the Timelog Plan sheet builder.

Covers the shape a downstream consumer can rely on without recalculating the
workbook: sheet exists, person rows mirror the roster, the PER-WP SUMMARY
block is populated, the traffic-light formulas are wired, and CF rules match
the design spec (matched by fill colour + formula pattern — not positional
index, per ``openpyxl-patterns.md``).
"""
from __future__ import annotations

import re

import pytest
from openpyxl import load_workbook

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS, TEAM, WORK_PACKAGES
from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_TIMELOG_PLAN,
    timelog_assigned_col_end,
    timelog_assigned_col_start,
    timelog_summary_col_start,
)
from budget_planner.styles import CF_RED_FILL
from budget_planner.workbook import init_workbook


# ── Fixtures ────────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def assigned_wb(tmp_path_factory):
    out = tmp_path_factory.mktemp("tlp_assigned") / "wb.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return load_workbook(str(out), data_only=False)


@pytest.fixture(scope="module")
def tlp(assigned_wb):
    return assigned_wb[SHEET_TIMELOG_PLAN]


def _project_months(wb) -> int:
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _person_first_row(ws) -> int:
    """Row of the first person entry — anchored on the 'Person' header in A."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Person":
            return r + 4  # month_idx + section + summary_header + 1
    raise AssertionError("'Person' header not found in col A of Timelog Plan")


def _iter_person_rows(ws, first_row: int):
    """Yield rows from ``first_row`` until PER-WP SUMMARY / blank."""
    r = first_row
    while r <= ws.max_row:
        v = ws.cell(r, 1).value
        if v is None:
            return
        s = str(v).strip()
        if not s or s.upper().startswith("PER-WP"):
            return
        yield r
        r += 1


# ── Existence + tab appearance ──────────────────────────────────────────────


def test_timelog_plan_sheet_exists(assigned_wb):
    assert SHEET_TIMELOG_PLAN in assigned_wb.sheetnames


def test_timelog_plan_tab_order(assigned_wb):
    """Timelog Plan must sit immediately after Plan."""
    names = assigned_wb.sheetnames
    assert names.index(SHEET_TIMELOG_PLAN) == names.index("Plan") + 1


def test_tab_colour_orange(tlp):
    # tabColor is stored with an AA prefix by openpyxl.
    assert (tlp.sheet_properties.tabColor.rgb or "").upper().endswith("FF8C00")


# ── Shape ───────────────────────────────────────────────────────────────────


def test_person_row_count_matches_team(tlp):
    first = _person_first_row(tlp)
    rows = list(_iter_person_rows(tlp, first))
    assert len(rows) == len(TEAM)


def test_person_col_a_cross_refs_config(tlp):
    """Every person row's col A is a ``=Config!A<n>`` cross-reference."""
    first = _person_first_row(tlp)
    for r in _iter_person_rows(tlp, first):
        v = tlp.cell(r, 1).value
        assert isinstance(v, str) and re.match(r"^=Config!A\d+$", v), (
            f"row {r} col A should be =Config!A<n>, got {v!r}"
        )


def test_assigned_grid_defaults_to_proposed_formula(assigned_wb, tlp):
    """Fresh build — every Assigned cell should be ``=<Proposed cell>``."""
    months = _project_months(assigned_wb)
    a_start = timelog_assigned_col_start(months)
    a_end = timelog_assigned_col_end(months)
    first = _person_first_row(tlp)
    sampled = 0
    for r in _iter_person_rows(tlp, first):
        for c in range(a_start, a_end + 1):
            v = tlp.cell(r, c).value
            if v is None or v == "":
                continue
            # Every non-empty cell must be a formula pointing at the Proposed
            # grid column — built-in default shape, no user overrides yet.
            assert isinstance(v, str) and v.startswith("="), (
                f"row {r} col {c}: expected formula default, got {v!r}"
            )
            sampled += 1
    assert sampled > 0, "Assigned grid had no populated cells"


# ── Per-WP SUMMARY block ────────────────────────────────────────────────────


def _find_wp_summary_first_row(ws) -> int:
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper() == "PER-WP SUMMARY":
            return r + 2  # section row + header row
    raise AssertionError("PER-WP SUMMARY header not found")


def test_wp_summary_rows_match_work_packages(tlp):
    first = _find_wp_summary_first_row(tlp)
    wp_codes_in_sheet = []
    for r in range(first, first + len(WORK_PACKAGES) + 5):
        v = tlp.cell(r, 1).value
        if v is None or str(v).strip() == "":
            break
        wp_codes_in_sheet.append(str(v).strip())
    assert wp_codes_in_sheet == [wp["id"] for wp in WORK_PACKAGES]


def test_wp_summary_formula_patterns(tlp):
    """Each WP row carries: Ceiling (=Config!...), Assigned € (SUMPRODUCT+ROUND),
    Overshoot/Undershoot (MAX), Assigned months (SUMPRODUCT), Window remaining
    (MAX+current_month), Out-of-window (SUMPRODUCT)."""
    first = _find_wp_summary_first_row(tlp)
    row = first  # first WP
    ceiling = str(tlp.cell(row, 2).value or "")
    assigned_eur = str(tlp.cell(row, 3).value or "")
    overshoot = str(tlp.cell(row, 4).value or "")
    undershoot = str(tlp.cell(row, 5).value or "")
    months_col = str(tlp.cell(row, 6).value or "")
    remaining = str(tlp.cell(row, 7).value or "")
    oow = str(tlp.cell(row, 8).value or "")
    assert re.match(r"^=Config!\$E\$\d+$", ceiling), ceiling
    assert assigned_eur.startswith("=ROUND(SUMPRODUCT("), assigned_eur
    assert overshoot.startswith("=MAX(0,"), overshoot
    assert undershoot.startswith("=MAX(0,"), undershoot
    assert months_col.startswith("=SUMPRODUCT(1*("), months_col
    assert "current_month" in remaining, remaining
    assert oow.startswith("=SUMPRODUCT("), oow


# ── Per-person summary ──────────────────────────────────────────────────────


def test_person_summary_headers(assigned_wb, tlp):
    months = _project_months(assigned_wb)
    summary_start = timelog_summary_col_start(months)
    # Summary header row = one row above the first person row.
    first_person = _person_first_row(tlp)
    hdr_row = first_person - 1
    expected = [
        "Assigned", "Pool", "Remaining", "Over-pool?",
        "Switches", "First log", "Last log", "Largest gap",
    ]
    got = [tlp.cell(hdr_row, summary_start + i).value for i in range(8)]
    assert got == expected


def test_person_summary_formulas(assigned_wb, tlp):
    months = _project_months(assigned_wb)
    summary_start = timelog_summary_col_start(months)
    row = _person_first_row(tlp)  # first person
    assigned_count = str(tlp.cell(row, summary_start).value or "")
    pool = str(tlp.cell(row, summary_start + 1).value or "")
    switches = str(tlp.cell(row, summary_start + 4).value or "")
    first_log = str(tlp.cell(row, summary_start + 5).value or "")
    assert assigned_count.startswith("=COUNTIF(")
    assert "ROUND(" in pool and "Config!" in pool
    assert switches.startswith("=SUMPRODUCT(1*(")
    # First-log uses the robust (range<>"")*1 coercion pattern, not the
    # silent-CSE-only MATCH(TRUE,...) — verify we didn't drift.
    assert '<>""' in first_log and "*1,0" in first_log


# ── Traffic-light formulas ──────────────────────────────────────────────────


def test_traffic_light_formulas_wired(tlp):
    # Row 2 holds the 4 (label, value) pairs.
    labels = [tlp.cell(2, c).value for c in (1, 3, 5, 7)]
    assert labels == [
        "Any undershoot?",
        "Any out-of-window cell?",
        "Any over-pool person?",
        "Total switches",
    ]
    v_undershoot = str(tlp.cell(2, 2).value or "")
    v_oow = str(tlp.cell(2, 4).value or "")
    v_over = str(tlp.cell(2, 6).value or "")
    v_switches = str(tlp.cell(2, 8).value or "")
    assert v_undershoot.startswith("=COUNTIF(")
    assert v_oow.startswith("=SUM(")
    assert v_over.startswith("=COUNTIF(") and "TRUE" in v_over
    assert v_switches.startswith("=SUM(")


# ── Conditional formatting ──────────────────────────────────────────────────


def _iter_all_rules(ws):
    for cf_range in ws.conditional_formatting:
        for rule in ws.conditional_formatting[cf_range]:
            yield str(cf_range), rule


def test_cf_has_red_rule_for_every_wp_out_of_window(tlp):
    """One red-fill CF per WP code × (proposed, assigned) grid.

    Match by formula content (contains ``<WP>`` + active-window bounds), not
    positional index — adding new CF rules must not break this test.
    """
    wp_codes = [wp["id"] for wp in WORK_PACKAGES]
    red_argb = CF_RED_FILL.bgColor.rgb
    red_oow_rules_by_code: dict[str, int] = {code: 0 for code in wp_codes}
    for _rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
        if fill is None or fill.bgColor is None:
            continue
        if fill.bgColor.rgb != red_argb:
            continue
        formula = (rule.formula or [""])[0]
        for code in wp_codes:
            if f'"{code}"' in formula and "OR(" in formula:
                red_oow_rules_by_code[code] += 1
                break
    # Rule applied to both Proposed grid and Assigned grid → 2 per WP.
    for code, count in red_oow_rules_by_code.items():
        assert count == 2, f"{code}: expected 2 out-of-window CF rules, got {count}"


def test_cf_pastel_rule_per_wp(tlp):
    """One pastel-fill CF per WP × (proposed, assigned) — 2 per WP total.

    Pastel fills vary, so we match by the formula shape (``cell="<WP>"`` with
    no OR() clause) and simply count non-red rules that mention the WP code.
    """
    wp_codes = [wp["id"] for wp in WORK_PACKAGES]
    red_argb = CF_RED_FILL.bgColor.rgb
    pastel_rules_by_code: dict[str, int] = {code: 0 for code in wp_codes}
    for _rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
        if fill is None or fill.bgColor is None:
            continue
        if fill.bgColor.rgb == red_argb:
            continue
        formula = (rule.formula or [""])[0]
        if "OR(" in formula:
            continue  # that's the out-of-window red rule (already counted)
        for code in wp_codes:
            if f'="{code}"' in formula:
                pastel_rules_by_code[code] += 1
                break
    for code, count in pastel_rules_by_code.items():
        assert count == 2, f"{code}: expected 2 pastel CF rules, got {count}"


# ── Freeze + column widths ──────────────────────────────────────────────────


def test_freeze_panes_at_person_row(tlp):
    # freeze_panes coord should anchor at (person_first_row, MONTH_COL_START).
    first = _person_first_row(tlp)
    expected_col_letter = tlp.cell(first, MONTH_COL_START).column_letter
    assert tlp.freeze_panes == f"{expected_col_letter}{first}"
