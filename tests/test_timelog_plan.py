"""Structural + recalc tests for the Timelog Plan v2 merged-grid layout.

Covers the shape a downstream consumer can rely on without recalculating the
workbook (layout, literal grid, toggle, CF rules) plus recalc-dependent
correctness for the First log / Last log SUMPRODUCT pattern (LO-stable —
openpyxl round-trip alone can't verify this).

CF rules are matched by fill colour + formula pattern, not positional index,
per ``.claude/rules/openpyxl-patterns.md``.
"""
from __future__ import annotations

import json
import re
import subprocess
from pathlib import Path

import pytest
from openpyxl import load_workbook

from budget_planner.layout import (
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
)
from budget_planner.styles import CF_GREEN_FILL, CF_RED_FILL
from budget_planner.timelog_plan import (
    SECTION_TITLE_GRID,
    SECTION_TITLE_PERSON_SUMMARY,
    SECTION_TITLE_WP_SUMMARY,
)

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


# ── Fixtures ────────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def assigned_wb(assigned_wb_path):
    return load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def tlp(assigned_wb):
    return assigned_wb[SHEET_TIMELOG_PLAN]


@pytest.fixture(scope="module")
def recalced_assigned_path(assigned_wb_path, tmp_path_factory):
    """Copy of assigned_wb_path with LibreOffice recalc applied.

    Recalc mutates the source file in place, and ``assigned_wb_path`` is a
    session fixture consumed by other tests — copy first so recalc doesn't
    mutate a shared artifact. Skips when LO is unavailable.
    """
    import shutil
    dst = tmp_path_factory.mktemp("tlp_recalc") / "assigned_recalced.xlsx"
    shutil.copyfile(assigned_wb_path, dst)
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), str(dst)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    try:
        json.loads(result.stdout)
    except json.JSONDecodeError:
        pytest.skip("recalc.py produced non-JSON output")
    return str(dst)


def _project_months(wb) -> int:
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _read_team(wb) -> list[dict]:
    from budget_planner.reader import read_team_from_config
    return read_team_from_config(wb["Config"])


def _read_wps(wb) -> list[dict]:
    from budget_planner.reader import read_config_sheet
    _params, wps = read_config_sheet(wb["Config"])
    return wps


def _find_section_title_row(ws, title: str) -> int:
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == title:
            return r
    raise AssertionError(f"Section title {title!r} not found in col A")


def _grid_person_first_row(ws) -> int:
    """GRID band: title, year, month, midx, persons → title + 4."""
    return _find_section_title_row(ws, SECTION_TITLE_GRID) + 4


def _summary_person_first_row(ws) -> int:
    """PER-PERSON SUMMARY: title, header, persons → title + 2."""
    return _find_section_title_row(ws, SECTION_TITLE_PERSON_SUMMARY) + 2


def _wp_summary_first_row(ws) -> int:
    """PER-WP SUMMARY: title, header, WPs → title + 2."""
    return _find_section_title_row(ws, SECTION_TITLE_WP_SUMMARY) + 2


def _iter_person_rows(ws, first_row: int):
    """Yield rows from ``first_row`` until next section title / blank / total row.

    Stops at PER- section titles, blank col A, and the "Monthly total +i" label
    (phase-3 row that sits directly under the grid in the same outline group).
    """
    r = first_row
    while r <= ws.max_row:
        v = ws.cell(r, 1).value
        if v is None:
            return
        s = str(v).strip()
        if not s or s.upper().startswith("PER-"):
            return
        if s.lower().startswith("monthly total"):
            return
        yield r
        r += 1


def _iter_all_rules(ws):
    for cf_range in ws.conditional_formatting:
        for rule in ws.conditional_formatting[cf_range]:
            yield str(cf_range.sqref), rule


# ── Existence + tab appearance ──────────────────────────────────────────────


def test_timelog_plan_sheet_exists(assigned_wb):
    assert SHEET_TIMELOG_PLAN in assigned_wb.sheetnames


def test_tab_colour_orange(tlp):
    assert (tlp.sheet_properties.tabColor.rgb or "").upper().endswith("FF8C00")


# ── Layout col shift ────────────────────────────────────────────────────────


def test_layout_col_shift(tlp, assigned_wb):
    """M1 starts at col B (was D in v1); Rate is last col (H) of PER-PERSON SUMMARY."""
    assert TIMELOG_MONTH_COL_START == 2

    # Merged-grid month header row = grid title + 2 (title, year, month).
    grid_title = _find_section_title_row(tlp, SECTION_TITLE_GRID)
    month_row = grid_title + 2
    m1_letter = tlp.cell(month_row, TIMELOG_MONTH_COL_START).column_letter
    assert m1_letter == "B"

    # Rate is at PER-PERSON SUMMARY col H (shifted left one after Largest gap drop).
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY) + 1
    assert tlp.cell(hdr_row, 8).value == "Rate"
    first_person = _summary_person_first_row(tlp)
    rate_val = str(tlp.cell(first_person, 8).value or "")
    assert rate_val.startswith("=VLOOKUP(") and "roster_lookup" in rate_val, rate_val


def test_freeze_panes_at_b3(tlp):
    """Freeze = ``B3`` — rows 1-2 + col A stay visible when scrolling."""
    assert tlp.freeze_panes == "B3"


# ── Merged grid ─────────────────────────────────────────────────────────────


def test_merged_grid_literal(tlp, assigned_wb):
    """Every merged-grid month cell holds a literal value, not a formula."""
    months = _project_months(assigned_wb)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + months - 1
    first = _grid_person_first_row(tlp)
    found = 0
    for r in _iter_person_rows(tlp, first):
        for c in range(m_start, m_end + 1):
            v = tlp.cell(r, c).value
            if v is None or v == "":
                continue
            assert not (isinstance(v, str) and v.startswith("=")), (
                f"row {r} col {c}: expected literal, got formula {v!r}"
            )
            found += 1
    assert found > 0, "merged grid had no populated cells (solver didn't run?)"


def test_grid_person_col_a_cross_refs_config(tlp):
    """Every person row's col A is a ``=Config!A<n>`` cross-reference."""
    for r in _iter_person_rows(tlp, _grid_person_first_row(tlp)):
        v = tlp.cell(r, 1).value
        assert isinstance(v, str) and re.match(r"^=Config!A\d+$", v), (
            f"row {r} col A should be =Config!A<n>, got {v!r}"
        )


def test_grid_person_row_count_matches_team(tlp, assigned_wb):
    team = _read_team(assigned_wb)
    rows = list(_iter_person_rows(tlp, _grid_person_first_row(tlp)))
    assert len(rows) == len(team)


# ── Underload removed ───────────────────────────────────────────────────────


def test_underload_removed(tlp, assigned_wb):
    """No Underload (Pool − Assigned) formula on any band — v2 dropped it."""
    months = _project_months(assigned_wb)

    # Merged grid: no col B metric formula. Col B is the first month cell
    # (should be literal or empty, never a cross-row subtraction).
    for r in _iter_person_rows(tlp, _grid_person_first_row(tlp)):
        v = tlp.cell(r, 2).value
        if isinstance(v, str) and v.startswith("="):
            assert "COUNTIF" not in v, (
                f"row {r} col B (M1) still carries Underload-like formula: {v!r}"
            )

    # PER-PERSON SUMMARY: col B is Assigned (a COUNTIF), not Underload.
    # The distinguishing signature of v1 Underload was "=$D{r}-…COUNTIF".
    # v2 Assigned is "=COUNTIF(…,\"WP*\")".
    for r in _iter_person_rows(tlp, _summary_person_first_row(tlp)):
        v = str(tlp.cell(r, 2).value or "")
        assert v.startswith("=COUNTIF("), (
            f"PER-PERSON SUMMARY col B should be Assigned=COUNTIF, got {v!r}"
        )
        assert "=$D" not in v, f"col B carries v1 Underload formula: {v!r}"

    # Header confirms semantic rename.
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY) + 1
    assert tlp.cell(hdr_row, 2).value == "Assigned"
    # Sanity: 'Underload' nowhere in headers.
    headers = [tlp.cell(hdr_row, c).value for c in range(1, 10)]
    assert "Underload" not in headers, headers

    # No CF rules with B{r}<0 / B{r}>0 pattern (the v1 Underload font-colour
    # CF). Band-range rules on col B from v1 should be gone.
    lt_rules = 0
    gt_rules = 0
    for rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        if not (rng.startswith("B") and ":" in rng):
            continue
        formula = (rule.formula or [""])[0]
        if re.match(r"^B\d+<0$", formula):
            lt_rules += 1
        elif re.match(r"^B\d+>0$", formula):
            gt_rules += 1
    assert lt_rules == 0, f"expected 0 'B<0' Underload rules, got {lt_rules}"
    assert gt_rules == 0, f"expected 0 'B>0' Underload rules, got {gt_rules}"

    # Months-loop used for bounds sanity only (ensures a non-empty grid).
    assert months > 0


# ── Rebuild-from-month Config param ─────────────────────────────────────────


def test_rebuild_from_month_named_range(assigned_wb):
    """``timelog_rebuild_from_month`` named range targets Config!$B$9 (row 9)
    and resolves to an integer default of 1. Replaces the v2 toggle cell.
    """
    dn = assigned_wb.defined_names["timelog_rebuild_from_month"]
    dests = list(dn.destinations)
    assert len(dests) == 1, dests
    sheet, ref = dests[0]
    assert sheet == "Config", sheet
    assert ref.replace("$", "").upper() == "B9", ref
    # Value: the BLUE_INPUT default is 1 (solve-all).
    cfg = assigned_wb["Config"]
    cell_val = cfg[ref.replace("$", "")].value
    assert cell_val == 1, f"expected default 1, got {cell_val!r}"


def test_no_toggle_row2_label_or_value(tlp):
    """Row 2 cols 9-10 (old toggle label/value) are not written by the builder
    on a fresh init — no BC shim for the legacy ``timelog_resolve_flag``.
    """
    for c in (9, 10):
        v = tlp.cell(2, c).value
        assert v in (None, ""), (
            f"row 2 col {c} expected empty after toggle removal, got {v!r}"
        )


def test_legacy_resolve_flag_named_range_absent(assigned_wb):
    """Fresh workbook: the legacy ``timelog_resolve_flag`` name is never written."""
    assert "timelog_resolve_flag" not in assigned_wb.defined_names


# ── Traffic-light formulas (cols 2/4/6/8) ───────────────────────────────────


def test_traffic_light_formulas_wired(tlp):
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
    # v2 wiring: Undershoot counts Buffer<0, Over-pool counts col D (was E),
    # Switches sums col E (was F).
    assert v_undershoot.startswith("=COUNTIF(") and '"<0"' in v_undershoot, v_undershoot
    assert v_oow.startswith("=SUM(")
    assert v_over.startswith("=COUNTIF(") and "TRUE" in v_over
    assert v_switches.startswith("=SUM(")


# ── PER-WP SUMMARY: headers + schema ────────────────────────────────────────


def test_wp_summary_headers_v2(tlp):
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_WP_SUMMARY) + 1
    # Aiko-port T3 (2026-04-23): col B renamed "Budget €" → "Budget +i"
    # (direct Config!F xref = +i). Buffer % inserted at col E; cols
    # Assigned months/WP sched months/Burn rate/OOW shifted one col right.
    expected = [
        "WP code", "Budget +i", "Assigned €", "Buffer €", "Buffer %",
        "Assigned months", "WP sched months", "Burn rate needed",
        "Any out-of-window?",
    ]
    got = [tlp.cell(hdr_row, 1 + i).value for i in range(len(expected))]
    # Accept "Any out-of-window?" or a shorter variant — the 'OOW' semantic.
    assert got[:8] == expected[:8], got
    assert got[8] and "out-of-window" in str(got[8]).lower(), got[8]


def test_wp_summary_rows_match_work_packages(tlp, assigned_wb):
    wps = _read_wps(assigned_wb)
    first = _wp_summary_first_row(tlp)
    sheet_codes = []
    for r in range(first, first + len(wps) + 5):
        v = tlp.cell(r, 1).value
        if v is None or str(v).strip() == "":
            break
        s = str(v).strip()
        # Phase-4 "Yearly total €" / "Yearly budget €" rows sit directly below
        # the last WP row (no spacer, same outline group) — stop before them.
        if s.lower().startswith("yearly "):
            break
        sheet_codes.append(s)
    assert sheet_codes == [wp["id"] for wp in wps]


# ── PER-WP SUMMARY: Buffer € ────────────────────────────────────────────────


def test_buffer_signed(tlp, assigned_wb):
    """Buffer € at col D = C − B (Assigned − Ceiling). CF red<0 / green>0."""
    wps = _read_wps(assigned_wb)
    first = _wp_summary_first_row(tlp)
    for i, _wp in enumerate(wps):
        r = first + i
        v = str(tlp.cell(r, 4).value or "")
        assert v == f"=C{r}-B{r}", f"row {r} Buffer €: {v!r}"

    # CF: band-range rule on col D, red fill for <0 and green fill for >0.
    last = first + len(wps) - 1
    red_argb = CF_RED_FILL.bgColor.rgb
    green_argb = CF_GREEN_FILL.bgColor.rgb
    red_hits = 0
    green_hits = 0
    for rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        if not rng.startswith("D") or ":" not in rng:
            continue
        if f"D{first}:D{last}" not in rng:
            continue
        fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
        if fill is None or fill.bgColor is None:
            continue
        formula = (rule.formula or [""])[0]
        if fill.bgColor.rgb == red_argb and "<0" in formula:
            red_hits += 1
        elif fill.bgColor.rgb == green_argb and ">0" in formula:
            green_hits += 1
    assert red_hits >= 1, f"expected a red Buffer€<0 CF rule, found {red_hits}"
    assert green_hits >= 1, f"expected a green Buffer€>0 CF rule, found {green_hits}"


# ── PER-WP SUMMARY: Burn rate ───────────────────────────────────────────────


def test_burn_rate_conditional(tlp, assigned_wb):
    """Col H formula: IF(AND(D<0, wallclock>0), -D/wallclock, ""). CF red on excess.

    Aiko-port T3 (2026-04-23): Burn rate shifted col G → H when Buffer % was
    inserted at col E. Buffer column at col D unchanged.
    """
    wps = _read_wps(assigned_wb)
    first = _wp_summary_first_row(tlp)
    last = first + len(wps) - 1
    for i, _wp in enumerate(wps):
        r = first + i
        v = str(tlp.cell(r, 8).value or "")
        assert v.startswith("=IF(AND(D"), f"row {r} Burn rate: {v!r}"
        assert f"D{r}<0" in v
        assert "current_month" in v
        assert '"")' in v or '"",' in v, v  # empty-string branch present

    # CF: red fill + MAX-rate comparison on a range including col H.
    red_argb = CF_RED_FILL.bgColor.rgb
    red_hits = 0
    for rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        if not rng.startswith("H") or ":" not in rng:
            continue
        if f"H{first}:H{last}" not in rng:
            continue
        fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
        if fill is None or fill.bgColor is None:
            continue
        if fill.bgColor.rgb != red_argb:
            continue
        formula = (rule.formula or [""])[0]
        if "MAX(" in formula and "H" in formula:
            red_hits += 1
    assert red_hits >= 1, f"expected red Burn-rate CF on H-band, found {red_hits}"


# ── PER-PERSON SUMMARY: v2 schema ───────────────────────────────────────────


def test_person_summary_headers_v2(tlp):
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY) + 1
    assert tlp.cell(hdr_row, 1).value == "Person"
    expected = [
        "Assigned", "Pool", "Over-pool?", "Switches",
        "First log", "Last log", "Rate",
    ]
    got = [tlp.cell(hdr_row, 2 + i).value for i in range(len(expected))]
    assert got == expected, got


def test_person_summary_shape(tlp):
    """PER-PERSON SUMMARY is 8 cols (Person + 7 metrics); no Largest gap header."""
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY) + 1
    headers = [tlp.cell(hdr_row, c).value for c in range(1, 10)]
    # Cols 1..8 are populated; col 9 must be blank/None (was 9-col v2).
    assert headers[0] == "Person"
    assert all(headers[i] is not None and headers[i] != "" for i in range(8)), headers
    assert headers[8] in (None, ""), (
        f"PER-PERSON SUMMARY col 9 should be empty after Largest gap drop, got {headers[8]!r}"
    )
    # "Largest gap" must not appear anywhere in the header row.
    assert "Largest gap" not in headers, headers


def test_rate_col_h(tlp, assigned_wb):
    """Rate lands at PER-PERSON SUMMARY col H; Assigned € SUMPRODUCT references col-H range."""
    # Rate formula at col H on the first person row.
    first_person = _summary_person_first_row(tlp)
    rate_val = str(tlp.cell(first_person, 8).value or "")
    assert rate_val.startswith("=VLOOKUP(") and "roster_lookup" in rate_val, rate_val

    # PER-WP SUMMARY Assigned € SUMPRODUCT (col C on first WP row) must reference
    # the col-H range for its rate vector, not col I.
    wp_first = _wp_summary_first_row(tlp)
    assigned_eur = str(tlp.cell(wp_first, 3).value or "")
    assert assigned_eur.startswith("=ROUND(SUMPRODUCT("), assigned_eur
    # Rate-vector segment must be anchored at col H on the summary-person rows.
    summary_last_row = first_person
    for r in _iter_person_rows(tlp, first_person):
        summary_last_row = r
    expected_fragment = f"H{first_person}:H{summary_last_row}"
    assert expected_fragment in assigned_eur, (
        f"expected rate-vector range {expected_fragment!r} in Assigned € formula, "
        f"got {assigned_eur!r}"
    )
    # Anti-assertion: the v2 col-I range must be gone.
    stale_fragment = f"I{first_person}:I{summary_last_row}"
    assert stale_fragment not in assigned_eur, (
        f"Assigned € still references stale col-I rate vector: {assigned_eur!r}"
    )


def test_first_log_formula_uses_sumproduct_min(tlp):
    """First log uses SUMPRODUCT(MIN(…)) pattern, NOT MATCH/LOOKUP."""
    first = _summary_person_first_row(tlp)
    v = str(tlp.cell(first, 6).value or "")  # col F
    assert v.startswith("=IFERROR(INDEX("), v
    assert "SUMPRODUCT(MIN(" in v, v
    assert "*99999" in v, v
    # Banned v1 patterns.
    assert "MATCH(TRUE" not in v and "MATCH(1," not in v, v


def test_last_log_formula_uses_sumproduct_max(tlp):
    """Last log uses SUMPRODUCT(MAX(…)) pattern, NOT LOOKUP(2,1/…)."""
    first = _summary_person_first_row(tlp)
    v = str(tlp.cell(first, 7).value or "")  # col G
    # Guarded against LO INDEX(range,0) implicit-intersection bug —
    # the MAX=0 branch must short-circuit before INDEX runs.
    assert v.startswith("=IFERROR(IF("), v
    assert "SUMPRODUCT(MAX(" in v, v
    assert "INDEX(" in v, v
    # Banned v1 patterns.
    assert "LOOKUP(2," not in v, v


# ── First/Last log recalc correctness (LO-specific) ─────────────────────────


def test_first_log_first_non_empty_recalc(recalced_assigned_path):
    """For every person with ≥1 assignment, First log returns the month-header
    value of their first non-empty grid cell (NOT blank, NOT last-month)."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + months - 1

    # Grid rows in assigned-sheet order; summary rows in the same order.
    grid_rows = list(_iter_person_rows(ws, _grid_person_first_row(ws)))
    summary_rows = list(_iter_person_rows(ws, _summary_person_first_row(ws)))
    assert len(grid_rows) == len(summary_rows)

    month_header_row = _find_section_title_row(ws, SECTION_TITLE_GRID) + 2

    checked = 0
    for g_row, s_row in zip(grid_rows, summary_rows):
        first_non_empty_col = None
        for c in range(m_start, m_end + 1):
            v = ws.cell(g_row, c).value
            if v not in (None, ""):
                first_non_empty_col = c
                break
        first_log = ws.cell(s_row, 6).value
        if first_non_empty_col is None:
            assert first_log in (None, ""), (
                f"summary row {s_row}: empty grid but First log = {first_log!r}"
            )
            continue
        expected = ws.cell(month_header_row, first_non_empty_col).value
        assert first_log == expected, (
            f"summary row {s_row}: expected First log = {expected!r}, got {first_log!r}"
        )
        checked += 1
    assert checked > 0, "no populated grid rows — fixture didn't exercise the formula"
    wb.close()


def test_last_log_last_non_empty_recalc(recalced_assigned_path):
    """For every person with ≥1 assignment, Last log returns the month-header
    value of their LAST non-empty grid cell (the v1 bug was always-last-month)."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + months - 1

    grid_rows = list(_iter_person_rows(ws, _grid_person_first_row(ws)))
    summary_rows = list(_iter_person_rows(ws, _summary_person_first_row(ws)))
    month_header_row = _find_section_title_row(ws, SECTION_TITLE_GRID) + 2
    project_last_header = ws.cell(month_header_row, m_end).value

    checked = 0
    any_partial_row_checked = False  # bystander: prove not always last month
    for g_row, s_row in zip(grid_rows, summary_rows):
        last_non_empty_col = None
        for c in range(m_end, m_start - 1, -1):
            v = ws.cell(g_row, c).value
            if v not in (None, ""):
                last_non_empty_col = c
                break
        last_log = ws.cell(s_row, 7).value
        if last_non_empty_col is None:
            assert last_log in (None, ""), (
                f"summary row {s_row}: empty grid but Last log = {last_log!r}"
            )
            continue
        expected = ws.cell(month_header_row, last_non_empty_col).value
        assert last_log == expected, (
            f"summary row {s_row}: expected Last log = {expected!r}, got {last_log!r}"
        )
        if last_non_empty_col < m_end:
            any_partial_row_checked = True
        checked += 1
    assert checked > 0, "no populated grid rows — fixture didn't exercise the formula"
    # If at least one person doesn't extend to the project's last month, we've
    # proven Last log isn't the v1 always-last-month bug.
    assert any_partial_row_checked, (
        "every populated row reached the final month — can't distinguish "
        "correct Last log from the v1 always-last-month bug; fixture needs a "
        "person whose assignments end before MN"
    )
    # Extra bystander: project_last_header is defined (sanity).
    assert project_last_header is not None
    wb.close()


# ── PER-WP SUMMARY: year cols (Phase 3) ─────────────────────────────────────


def _expected_year_groups(project_start, project_months: int) -> list[tuple[int, int, int]]:
    """Mirror of _calendar_year_groups: (year, first_month, last_month) 1-indexed."""
    from dateutil.relativedelta import relativedelta
    out: list[tuple[int, int, int]] = []
    m = 0
    while m < project_months:
        cal_date = project_start + relativedelta(months=m)
        months_left_in_year = 12 - cal_date.month + 1
        span = min(months_left_in_year, project_months - m)
        out.append((cal_date.year, m + 1, m + span))
        m += span
    return out


def test_year_cols_count_and_headers(tlp, assigned_wb):
    """Year col count = distinct calendar years spanned; headers = Y{yyyy} +i
    with WITH_INDIRECT_HEADER_FILL (T7.3 display-layer +i flip)."""
    from budget_planner.styles import WITH_INDIRECT_HEADER_FILL
    params, _wps = _read_config(assigned_wb)
    months = params["months"]
    groups = _expected_year_groups(params["start"], months)

    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_WP_SUMMARY) + 1
    base_cols = 9  # cols A..I hold the base metrics (Aiko-port T3 added Buffer % at col E)
    expected_fill_rgb = WITH_INDIRECT_HEADER_FILL.fgColor.rgb
    for i, (year, _fm, _lm) in enumerate(groups):
        cell = tlp.cell(hdr_row, base_cols + 1 + i)
        assert cell.value == f"Y{year} +i", (
            f"col {base_cols + 1 + i}: {cell.value!r}"
        )
        assert cell.fill.fgColor.rgb == expected_fill_rgb, (
            f"col {base_cols + 1 + i}: expected WITH_INDIRECT_HEADER_FILL "
            f"{expected_fill_rgb}, got {cell.fill.fgColor.rgb!r}"
        )
    # No stray header past the last year col.
    trailing = tlp.cell(hdr_row, base_cols + len(groups) + 1).value
    assert trailing in (None, ""), f"unexpected header past year cols: {trailing!r}"


def test_year_cols_overlap_blanks(tlp, assigned_wb):
    """Cells where WP's active window doesn't overlap the year are literal ``""``."""
    params, wps = _read_config(assigned_wb)
    groups = _expected_year_groups(params["start"], params["months"])
    first = _wp_summary_first_row(tlp)
    base_cols = 9

    checked_nonoverlap = 0
    checked_overlap = 0
    for wi, wp in enumerate(wps):
        active_start, active_end = int(wp["start"]), int(wp["end"])
        for yi, (_year, fm, lm) in enumerate(groups):
            col = base_cols + 1 + yi
            val = tlp.cell(first + wi, col).value
            overlaps = fm <= active_end and lm >= active_start
            if overlaps:
                assert isinstance(val, str) and val.startswith("="), (
                    f"WP {wp['id']} year {groups[yi][0]}: expected formula, got {val!r}"
                )
                checked_overlap += 1
            else:
                assert val in ("", None), (
                    f"WP {wp['id']} year {groups[yi][0]}: expected blank, got {val!r}"
                )
                checked_nonoverlap += 1
    # Bystander: ensure we actually exercised both branches.
    assert checked_overlap > 0
    # A non-overlap bystander is only possible if some WP doesn't span every
    # year. In the demo fixture WPs are staggered, so expect ≥1 non-overlap.
    assert checked_nonoverlap > 0, (
        "no non-overlap year cells in fixture — selectivity test is vacuous"
    )


def test_year_cols_sum_matches_assigned(recalced_assigned_path):
    """Sum across year cols on each WP row ≈ that WP's Assigned € × 1.25.

    Post-T7.3: year cells are +i (display-layer ×1.25 wrap inside
    ROUND/SUMPRODUCT); col C Assigned € stays -i. So year-sum ≈ Assigned × 1.25.
    Each year cell is independently ROUND'd per the EUR leaf-rounding rule,
    so the sum of rounded-per-year can differ from the expected +i Assigned
    by at most one EUR per year (round-half-even). We allow a tolerance of
    ``num_years`` EUR.
    """
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    params, wps = _read_config(wb)
    groups = _expected_year_groups(params["start"], params["months"])
    first = _wp_summary_first_row(ws)
    base_cols = 9
    tolerance = len(groups)

    checked = 0
    for wi, wp in enumerate(wps):
        assigned_eur = ws.cell(first + wi, 3).value
        if not isinstance(assigned_eur, (int, float)):
            continue
        year_sum = 0.0
        for yi in range(len(groups)):
            v = ws.cell(first + wi, base_cols + 1 + yi).value
            if isinstance(v, (int, float)):
                year_sum += float(v)
        expected_plus_i = float(assigned_eur) * 1.25
        assert abs(year_sum - expected_plus_i) <= tolerance, (
            f"WP {wp['id']}: sum of year cols {year_sum} vs Assigned × 1.25 "
            f"{expected_plus_i} (tolerance ±{tolerance})"
        )
        checked += 1
    assert checked > 0, "no numeric Assigned € values — recalc didn't populate"
    wb.close()


def _read_config(wb):
    from budget_planner.reader import read_config_sheet
    return read_config_sheet(wb["Config"])


# ── Vacation-month model (Phase 2 enhancements) ─────────────────────────────


def _vacation_project_months(project_start, project_months: int, vac_cal_month: int) -> set[int]:
    """Return 1-indexed project months whose calendar month == vac_cal_month."""
    from dateutil.relativedelta import relativedelta
    out: set[int] = set()
    for i in range(project_months):
        d = project_start + relativedelta(months=i)
        if d.month == vac_cal_month:
            out.add(i + 1)
    return out


def test_vacation_config_scalar(assigned_wb):
    """`vacation_month` named range resolves to Config!$B$<row>, default 'Jul'."""
    assert "vacation_month" in assigned_wb.defined_names
    dn = assigned_wb.defined_names["vacation_month"]
    dests = list(dn.destinations)
    assert len(dests) == 1
    sheet, ref = dests[0]
    assert sheet == "Config"
    cfg = assigned_wb["Config"]
    val = cfg[ref.replace("$", "")].value
    # Default fixture builds with default project params → vacation = "Jul".
    assert val == "Jul", f"expected default vacation_month 'Jul', got {val!r}"


def test_num_vacation_months_literal(assigned_wb):
    """`num_vacation_months` Config literal matches count of July months in span."""
    assert "num_vacation_months" in assigned_wb.defined_names
    params, _wps = _read_config(assigned_wb)
    dn = assigned_wb.defined_names["num_vacation_months"]
    dests = list(dn.destinations)
    assert len(dests) == 1
    _sheet, ref = dests[0]
    cfg = assigned_wb["Config"]
    val = cfg[ref.replace("$", "")].value
    expected = len(_vacation_project_months(params["start"], params["months"], 7))
    assert val == expected, (
        f"num_vacation_months: expected {expected}, got {val!r}"
    )


def test_vacation_cells_locked(assigned_wb, tlp):
    """Every July cell across person rows is '', has VACATION_FILL, locked=True."""
    from budget_planner.styles import VACATION_FILL
    params, _wps = _read_config(assigned_wb)
    m_start = TIMELOG_MONTH_COL_START
    vac_months = _vacation_project_months(params["start"], params["months"], 7)
    expected_fill_rgb = (VACATION_FILL.fgColor.rgb or "").upper()

    grid_rows = list(_iter_person_rows(tlp, _grid_person_first_row(tlp)))
    assert grid_rows, "no person rows in grid"
    assert vac_months, "fixture span has no July months — vacation tier untested"

    checked_vac = 0
    checked_nonvac = 0
    for r in grid_rows:
        for month_num in vac_months:
            col = m_start + month_num - 1
            cell = tlp.cell(r, col)
            assert cell.value in (None, ""), (
                f"row {r} col {col} (month {month_num}): vacation cell "
                f"must be empty, got {cell.value!r}"
            )
            rgb = (cell.fill.fgColor.rgb or "").upper() if cell.fill and cell.fill.fgColor else ""
            assert rgb == expected_fill_rgb, (
                f"row {r} col {col}: expected VACATION_FILL {expected_fill_rgb}, got {rgb}"
            )
            assert cell.protection.locked is True, (
                f"row {r} col {col}: vacation cell must be locked=True"
            )
            checked_vac += 1
        # Bystander: at least one non-vacation in-window cell on the same row
        # must NOT have VACATION_FILL — proves the fill is vacation-specific.
        for mi in range(params["months"]):
            month_num = mi + 1
            if month_num in vac_months:
                continue
            cell = tlp.cell(r, m_start + mi)
            rgb = (cell.fill.fgColor.rgb or "").upper() if cell.fill and cell.fill.fgColor else ""
            if rgb == expected_fill_rgb:
                pytest.fail(
                    f"row {r} col {m_start + mi} (month {month_num}): "
                    "non-vacation cell incorrectly carries VACATION_FILL"
                )
            checked_nonvac += 1
            break
    assert checked_vac > 0 and checked_nonvac > 0


def test_pool_formula_accounts_for_vacation(tlp):
    """Pool formula on PER-PERSON SUMMARY col C references num_vacation_months."""
    first = _summary_person_first_row(tlp)
    pool_val = str(tlp.cell(first, 3).value or "")
    assert pool_val.startswith("=ROUND("), pool_val
    assert "num_vacation_months" in pool_val, (
        f"Pool formula must subtract num_vacation_months, got {pool_val!r}"
    )
    # Anti-assertion: the pre-enhancement form (plain multiply) is gone.
    assert "project_months,0)" not in pool_val or "-num_vacation_months" in pool_val


def test_pool_value_after_recalc(recalced_assigned_path):
    """Recalced Pool = round(max_fte × (project_months − num_vacation_months))."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    cfg = wb["Config"]
    params, _wps = _read_config(wb)
    num_vac = len(_vacation_project_months(params["start"], params["months"], 7))
    available_months = params["months"] - num_vac

    # Walk roster → max_fte, then match PER-PERSON SUMMARY row.
    team = _read_team(wb)
    name_to_fte = {m["name"]: float(m["max_fte"]) for m in team}

    summary_rows = list(_iter_person_rows(ws, _summary_person_first_row(ws)))
    assert summary_rows, "no summary rows"

    cfg_ws = cfg  # silence pyright
    checked = 0
    for r in summary_rows:
        col_a = ws.cell(r, 1).value
        # Resolve cross-ref to the Config person name.
        name: str | None = None
        if isinstance(col_a, str) and col_a.startswith("="):
            import re as _re
            m = _re.match(r"=Config!A(\d+)", col_a)
            if m:
                name = str(cfg_ws.cell(int(m.group(1)), 1).value or "")
        else:
            name = str(col_a or "")
        if not name or name not in name_to_fte:
            continue
        expected = round(name_to_fte[name] * available_months)
        actual = ws.cell(r, 3).value
        assert actual == expected, (
            f"Pool for {name!r}: expected {expected}, got {actual!r} "
            f"(max_fte={name_to_fte[name]}, available={available_months})"
        )
        checked += 1
    assert checked > 0, "no pool cells validated — recalc didn't populate"
    wb.close()


def test_solver_never_places_wp_in_vacation_cell(assigned_wb, tlp):
    """TIMELOG GRID: no WP code lands in any July cell across all person rows."""
    params, _wps = _read_config(assigned_wb)
    vac_months = _vacation_project_months(params["start"], params["months"], 7)
    assert vac_months, "no July months in span — vacation suppression untested"
    m_start = TIMELOG_MONTH_COL_START

    violations: list[tuple[int, int, object]] = []
    for r in _iter_person_rows(tlp, _grid_person_first_row(tlp)):
        for month_num in vac_months:
            v = tlp.cell(r, m_start + month_num - 1).value
            if v not in (None, ""):
                violations.append((r, month_num, v))
    assert not violations, f"solver placed codes in vacation cells: {violations}"


# ── Monthly total € row (Phase 3 enhancement) ───────────────────────────────


def _monthly_total_row(ws) -> int:
    """Grid band layout: title, year, month, midx, persons…, monthly total."""
    grid_rows = list(_iter_person_rows(ws, _grid_person_first_row(ws)))
    assert grid_rows, "no person rows — monthly total row placement undefined"
    return grid_rows[-1] + 1


def test_monthly_total_row_present(tlp, assigned_wb):
    """Monthly total +i row sits directly under the last person row with
    ROUND(SUMPRODUCT(…)*1.25,0) formulas per month col and
    WITH_INDIRECT_HEADER_FILL on the label cell (T7.3)."""
    from budget_planner.styles import WITH_INDIRECT_HEADER_FILL
    months = _project_months(assigned_wb)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + months - 1

    total_row = _monthly_total_row(tlp)
    label_cell = tlp.cell(total_row, 1)
    assert label_cell.value == "Monthly total +i"
    assert (
        label_cell.fill.fgColor.rgb == WITH_INDIRECT_HEADER_FILL.fgColor.rgb
    ), f"label fill {label_cell.fill.fgColor.rgb!r}"

    # Every month col must carry a ROUND(SUMPRODUCT(…)*1.25,0) formula.
    checked = 0
    for c in range(m_start, m_end + 1):
        v = str(tlp.cell(total_row, c).value or "")
        assert v.startswith("=ROUND(SUMPRODUCT("), (
            f"col {c}: expected ROUND(SUMPRODUCT(…), got {v!r}"
        )
        assert '<>""' in v, v
        assert "*1.25" in v, f"col {c}: expected ×1.25 wrap, got {v!r}"
        checked += 1
    assert checked == months


def test_monthly_total_row_vacation_fill(tlp, assigned_wb):
    """July cells on the Monthly total +i row carry VACATION_FILL; non-vacation
    month cells on the same row do NOT (bystander for selectivity)."""
    from budget_planner.styles import VACATION_FILL
    params, _wps = _read_config(assigned_wb)
    vac_months = _vacation_project_months(params["start"], params["months"], 7)
    assert vac_months, "fixture has no July months — fill tier untestable"

    m_start = TIMELOG_MONTH_COL_START
    total_row = _monthly_total_row(tlp)
    expected_fill_rgb = (VACATION_FILL.fgColor.rgb or "").upper()

    checked_vac = 0
    checked_nonvac = 0
    for mi in range(params["months"]):
        month_num = mi + 1
        cell = tlp.cell(total_row, m_start + mi)
        rgb = (
            (cell.fill.fgColor.rgb or "").upper()
            if cell.fill and cell.fill.fgColor else ""
        )
        if month_num in vac_months:
            assert rgb == expected_fill_rgb, (
                f"total-row col {m_start + mi} (month {month_num}): "
                f"expected VACATION_FILL {expected_fill_rgb}, got {rgb}"
            )
            checked_vac += 1
        else:
            assert rgb != expected_fill_rgb, (
                f"total-row col {m_start + mi} (month {month_num}): "
                "non-vacation cell incorrectly carries VACATION_FILL"
            )
            checked_nonvac += 1
    assert checked_vac > 0 and checked_nonvac > 0


def test_monthly_total_vacation_zero(recalced_assigned_path):
    """Every vacation-month cell on the Monthly total +i row recalcs to 0."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    params, _wps = _read_config(wb)
    vac_months = _vacation_project_months(params["start"], params["months"], 7)
    assert vac_months, "no July months — vacation cell total untested"
    m_start = TIMELOG_MONTH_COL_START
    total_row = _monthly_total_row(ws)

    for month_num in vac_months:
        v = ws.cell(total_row, m_start + month_num - 1).value
        assert v in (0, 0.0), (
            f"vacation month {month_num} total expected 0, got {v!r}"
        )
    wb.close()


def test_monthly_total_non_vacation_nonzero(recalced_assigned_path):
    """For each non-vacation month with ≥1 assigned person, the Monthly total +i
    cell equals round(sum of those persons' Rate values × 1.25). Rates are -i
    per D7; display-layer ×1.25 wrap flips to +i (T7.3)."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    params, _wps = _read_config(wb)
    vac_months = _vacation_project_months(params["start"], params["months"], 7)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + params["months"] - 1
    total_row = _monthly_total_row(ws)

    # Map each grid row → Rate value from PER-PERSON SUMMARY col H.
    grid_rows = list(_iter_person_rows(ws, _grid_person_first_row(ws)))
    summary_rows = list(_iter_person_rows(ws, _summary_person_first_row(ws)))
    assert len(grid_rows) == len(summary_rows)
    grid_rate: dict[int, float] = {}
    for g_row, s_row in zip(grid_rows, summary_rows):
        rate = ws.cell(s_row, 8).value
        if isinstance(rate, (int, float)):
            grid_rate[g_row] = float(rate)

    checked = 0
    for c in range(m_start, m_end + 1):
        month_num = c - m_start + 1
        if month_num in vac_months:
            continue
        expected = 0.0
        for g_row, rate in grid_rate.items():
            v = ws.cell(g_row, c).value
            if v not in (None, ""):
                expected += rate
        got = ws.cell(total_row, c).value
        expected_plus_i = round(expected * 1.25)
        assert got == expected_plus_i, (
            f"month {month_num} (col {c}): expected {expected_plus_i} "
            f"(= round({expected} × 1.25)), got {got!r}"
        )
        if expected > 0:
            checked += 1
    assert checked > 0, (
        "no non-vacation month had any assigned persons — bystander gate vacuous"
    )
    wb.close()


# ── Monthly person cost +i row (Aiko-port T4) ───────────────────────────────


def test_monthly_person_cost_row_present(tlp, assigned_wb):
    """Monthly person cost +i row sits directly under Monthly total +i with
    the same ROUND(SUMPRODUCT(...)*1.25,0) formula shape per month col and
    WITH_INDIRECT_HEADER_FILL on the label cell (Aiko-port T4, Q-D)."""
    from budget_planner.styles import WITH_INDIRECT_HEADER_FILL
    months = _project_months(assigned_wb)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + months - 1

    person_cost_row = _monthly_total_row(tlp) + 1
    label_cell = tlp.cell(person_cost_row, 1)
    assert label_cell.value == "Monthly person cost +i"
    assert (
        label_cell.fill.fgColor.rgb == WITH_INDIRECT_HEADER_FILL.fgColor.rgb
    ), f"label fill {label_cell.fill.fgColor.rgb!r}"

    checked = 0
    for c in range(m_start, m_end + 1):
        v = str(tlp.cell(person_cost_row, c).value or "")
        assert v.startswith("=ROUND(SUMPRODUCT("), (
            f"col {c}: expected ROUND(SUMPRODUCT(…), got {v!r}"
        )
        assert '<>""' in v, v
        assert "*1.25" in v, f"col {c}: expected ×1.25 wrap, got {v!r}"
        checked += 1
    assert checked == months


def test_monthly_person_cost_row_precedes_summary_title(tlp):
    """Person-cost row abuts PER-PERSON SUMMARY title with no spacer — this
    anchors summary_person_first_row at its historical offset, which the
    row-30 formula depends on (formula refs the Rate col range at
    summary_person_first_row..summary_person_last_row). Bystander: Monthly
    total +i row is still at person_cost_row - 1."""
    person_cost_row = _monthly_total_row(tlp) + 1
    # Monthly total +i row is the row above — confirms layout order.
    assert (
        str(tlp.cell(person_cost_row - 1, 1).value or "")
        .lower().startswith("monthly total")
    )
    # Next row is the PER-PERSON SUMMARY section title (no spacer between).
    next_val = tlp.cell(person_cost_row + 1, 1).value
    assert next_val == SECTION_TITLE_PERSON_SUMMARY, (
        f"expected PER-PERSON SUMMARY title at row "
        f"{person_cost_row + 1}, got {next_val!r}"
    )


# ── Yearly total € / Yearly budget € rows (Phase 4 enhancement) ─────────────


def _yearly_rows(tlp) -> tuple[int, int]:
    """Return (yearly_total_row, yearly_budget_row) — two rows directly beneath
    the last WP row in PER-WP SUMMARY (no spacer, same outline group).

    Post-D9 (2026-04-22): row labels live at col H (TIMELOG_YEARLY_LABEL_COL),
    not col A."""
    from budget_planner.layout import TIMELOG_YEARLY_LABEL_COL
    first = _wp_summary_first_row(tlp)
    # Walk down col H until we hit a cell whose text starts with "Yearly ".
    # Blank cells at WP rows are normal (col H at WP rows is the
    # "Any out-of-window?" formula output, not a label) — keep scanning.
    r = first
    while r <= tlp.max_row:
        v = tlp.cell(r, TIMELOG_YEARLY_LABEL_COL).value
        if v is not None and str(v).strip().lower().startswith("yearly "):
            break
        r += 1
    else:
        raise AssertionError("no Yearly rows found in PER-WP SUMMARY")
    return r, r + 1


def test_yearly_rows_present(tlp):
    """Both rows exist directly below the last WP row with expected labels
    at col I (post-D9 col-H shift + Aiko-port T3 one-col right shift)."""
    from budget_planner.layout import TIMELOG_YEARLY_LABEL_COL
    from budget_planner.styles import WITH_INDIRECT_HEADER_FILL
    y_total, y_budget = _yearly_rows(tlp)
    # Aiko-port T3 (2026-04-23) renamed "Yearly total +i" → "Yearly Assigned +i"
    # per brainstorm locked spec.
    assert tlp.cell(y_total, TIMELOG_YEARLY_LABEL_COL).value == "Yearly Assigned +i"
    assert tlp.cell(y_budget, TIMELOG_YEARLY_LABEL_COL).value == "Yearly Budget +i"

    # Col A on these rows is blank (label shifted to col H).
    assert tlp.cell(y_total, 1).value in (None, "")
    assert tlp.cell(y_budget, 1).value in (None, "")

    # With-indirect header fill on the label cell.
    assert (
        tlp.cell(y_total, TIMELOG_YEARLY_LABEL_COL).fill.fgColor.rgb
        == WITH_INDIRECT_HEADER_FILL.fgColor.rgb
    )
    assert (
        tlp.cell(y_budget, TIMELOG_YEARLY_LABEL_COL).fill.fgColor.rgb
        == WITH_INDIRECT_HEADER_FILL.fgColor.rgb
    )

    # Outline level: rows must be in the PER-WP SUMMARY outline group (level 1).
    assert tlp.row_dimensions[y_total].outline_level == 1
    assert tlp.row_dimensions[y_budget].outline_level == 1


def test_yearly_total_col_c_formula(tlp, assigned_wb):
    """Yearly total col C cross-check = raw SUM(C{first}:C{last}) (Aiko-port Q-B).

    PER-WP SUMMARY col C stays -i; the cross-check is -i by construction.
    Per-year cells to the right of col C stay +i via the ×1.25 wrap at the
    SUMPRODUCT leaf (T7.3), but the cross-check no longer wraps in ×1.25."""
    wps = _read_wps(assigned_wb)
    y_total, _ = _yearly_rows(tlp)
    first = _wp_summary_first_row(tlp)
    last = first + len(wps) - 1
    v = str(tlp.cell(y_total, 3).value or "")
    assert v == f"=SUM(C{first}:C{last})", v


def test_yearly_budget_col_c_formula(tlp, assigned_wb):
    """Yearly Budget +i col C cross-check = raw SUM(B{first}:B{last}) (Aiko-port Q-B).

    PER-WP SUMMARY col B now xrefs Config col F (Budget +i, COL_WP_BUDGET_PLUS_I)
    directly, so SUM(B) is already +i — no ×1.25 multiplier needed."""
    wps = _read_wps(assigned_wb)
    _, y_budget = _yearly_rows(tlp)
    first = _wp_summary_first_row(tlp)
    last = first + len(wps) - 1
    v = str(tlp.cell(y_budget, 3).value or "")
    assert v == f"=SUM(B{first}:B{last})", v


def test_yearly_total_col_c_crosscheck(recalced_assigned_path):
    """Col C value equals row-sum of year cells (within leaf-rounding tolerance).

    Aiko-port Q-B dropped the ×1.25 from col C, so col C is -i and year cells
    are +i. Drift-detection cross-check now compares ``col_c * 1.25`` vs
    year_sum:
      - col_c = SUM(C{first}:C{last}) — sums per-WP Assigned € leaves, each
        independently rounded; error ≤ len(wps) * 0.5. ×1.25 scales both the
        signal and the error for +i-parity comparison.
      - year_sum = Σ ROUND(SUMPRODUCT_{wp,y} * 1.25, 0) over len(wps)×len(groups)
        independently-rounded cells; error ≤ len(wps) * len(groups) * 0.5.
    Safe tolerance: len(wps) * (len(groups) + 2) covers both paths with
    slack for float64 noise."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    params, wps = _read_config(wb)
    groups = _expected_year_groups(params["start"], params["months"])
    y_total, _ = _yearly_rows(ws)
    base_cols = 9

    col_c = ws.cell(y_total, 3).value
    year_sum = 0.0
    for yi in range(len(groups)):
        v = ws.cell(y_total, base_cols + 1 + yi).value
        if isinstance(v, (int, float)):
            year_sum += float(v)
    assert isinstance(col_c, (int, float)), f"col C not numeric: {col_c!r}"
    tolerance = len(wps) * (len(groups) + 2)
    col_c_plus_i = float(col_c) * 1.25
    assert abs(year_sum - col_c_plus_i) <= tolerance, (
        f"Yearly total: col C×1.25 {col_c_plus_i} vs year-sum {year_sum} "
        f"(tolerance ±{tolerance})"
    )
    wb.close()


def test_yearly_budget_col_c_crosscheck(recalced_assigned_path):
    """Col C value equals row-sum of year cells (within fp tolerance)."""
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    params, _wps = _read_config(wb)
    groups = _expected_year_groups(params["start"], params["months"])
    _, y_budget = _yearly_rows(ws)
    base_cols = 9

    col_c = ws.cell(y_budget, 3).value
    year_sum = 0.0
    for yi in range(len(groups)):
        v = ws.cell(y_budget, base_cols + 1 + yi).value
        if isinstance(v, (int, float)):
            year_sum += float(v)
    assert isinstance(col_c, (int, float)), f"col C not numeric: {col_c!r}"
    # Yearly Budget +i uses un-ROUND'd term sums — per design §5 the formula is
    # `=Σ prop_ky × Config!F{wp_row}` with no outer ROUND (col F = Budget +i).
    # In real arithmetic Σ_y prop_ky = 1 for every WP, so Σ year cells =
    # Σ budgets_+i. Col C post-Aiko-port Q-B is `=SUM(B{first}:B{last})` with
    # col B now direct Config!F (Budget +i) — so SUM(B) = Σ budgets_+i exactly.
    # Equality holds; fp rounding errors are the only drift source.
    # ±1 EUR tolerance covers float64 noise.
    # If the per-year formula is later changed to wrap each cell in
    # `ROUND(..., 0)` (e.g. to align with the Yearly total +i convention),
    # bump this tolerance to `len(groups)` EUR to accommodate ±0.5 EUR per year.
    assert abs(year_sum - float(col_c)) <= 1.0, (
        f"Yearly budget: col C {col_c} vs year-sum {year_sum}"
    )
    wb.close()


def test_yearly_budget_single_wp_single_year(tlp, assigned_wb):
    """For any WP whose active window lies inside one calendar year, that year's
    Yearly Budget +i term = 1.0 × Config!F{wp_row} (col F = Budget +i)."""
    params, wps = _read_wps_and_schedule_and_params(assigned_wb)
    groups = _expected_year_groups(params["start"], params["months"])
    _, y_budget = _yearly_rows(tlp)
    base_cols = 9

    # Find a WP fully contained in one year (existence, not every-row).
    candidates = []
    for wi, wp in enumerate(wps):
        active_start, active_end = int(wp["start"]), int(wp["end"])
        for yi, (_yr, fm, lm) in enumerate(groups):
            if fm <= active_start and active_end <= lm:
                candidates.append((wi, wp, yi, fm, lm))
                break
    if not candidates:
        pytest.skip("no WP fully contained in a single calendar year in fixture")

    wi, wp, yi, _fm, _lm = candidates[0]
    cfg_row = _config_wp_row(assigned_wb, wp["id"])
    formula = str(tlp.cell(y_budget, base_cols + 1 + yi).value or "")
    # Term should contain "1.0*Config!$F${cfg_row}" exactly once (cell_ref uses
    # absolute refs; Config col F = Budget +i per D4). Other WP terms may
    # coexist in the same year's formula.
    expected_term = f"1.0*Config!$F${cfg_row}"
    assert expected_term in formula, (
        f"WP {wp['id']} fully in year idx {yi}: expected {expected_term!r} "
        f"in yearly budget formula, got {formula!r}"
    )


def _read_wps_and_schedule_and_params(wb):
    from budget_planner.reader import read_config_sheet
    params, wps = read_config_sheet(wb["Config"])
    return params, wps


def _config_wp_row(wb, wp_id: str) -> int:
    """Return the 1-indexed row on Config where this WP's row lives."""
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == wp_id:
            return r
    raise AssertionError(f"WP {wp_id!r} not found on Config sheet")
