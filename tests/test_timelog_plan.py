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


def _read_wps_and_schedule(wb) -> tuple[list[dict], list[dict]]:
    from budget_planner.reader import read_config_sheet
    _params, wps, sched = read_config_sheet(wb["Config"])
    return wps, sched


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
    """Yield rows from ``first_row`` until next section title / blank."""
    r = first_row
    while r <= ws.max_row:
        v = ws.cell(r, 1).value
        if v is None:
            return
        s = str(v).strip()
        if not s or s.upper().startswith("PER-"):
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
    """M1 starts at col B (was D in v1); Rate is last col (I) of PER-PERSON SUMMARY."""
    assert TIMELOG_MONTH_COL_START == 2

    # Merged-grid month header row = grid title + 2 (title, year, month).
    grid_title = _find_section_title_row(tlp, SECTION_TITLE_GRID)
    month_row = grid_title + 2
    m1_letter = tlp.cell(month_row, TIMELOG_MONTH_COL_START).column_letter
    assert m1_letter == "B"

    # Rate is at PER-PERSON SUMMARY col I.
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY) + 1
    assert tlp.cell(hdr_row, 9).value == "Rate"
    first_person = _summary_person_first_row(tlp)
    rate_val = str(tlp.cell(first_person, 9).value or "")
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


# ── Toggle cell ─────────────────────────────────────────────────────────────


def test_toggle_cell_present(assigned_wb, tlp):
    """Row 2 col J (col 10) holds literal TRUE; named range resolves."""
    cell = tlp.cell(2, 10)
    assert cell.value is True, (
        f"toggle cell expected True, got {cell.value!r} ({type(cell.value)})"
    )

    # Label in col I (col 9) anchors what the toggle is.
    label = str(tlp.cell(2, 9).value or "")
    assert "re-run" in label.lower() or "rebuild" in label.lower(), label

    # Named range resolves to the same cell.
    dn = assigned_wb.defined_names["timelog_resolve_flag"]
    dests = list(dn.destinations)
    assert len(dests) == 1
    sheet, ref = dests[0]
    assert sheet == SHEET_TIMELOG_PLAN
    # Accept either $J$2 or 'Timelog Plan'!$J$2 after Excel/LO munging —
    # destinations already strips the sheet prefix.
    assert ref.replace("$", "").upper() == "J2", ref


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
    expected = [
        "WP code", "Ceiling €", "Assigned €", "Buffer €",
        "Assigned months", "WP sched months", "Burn rate needed",
        "Any out-of-window?",
    ]
    got = [tlp.cell(hdr_row, 1 + i).value for i in range(len(expected))]
    # Accept "Any out-of-window?" or a shorter variant — the 'OOW' semantic.
    assert got[:7] == expected[:7], got
    assert got[7] and "out-of-window" in str(got[7]).lower(), got[7]


def test_wp_summary_rows_match_work_packages(tlp, assigned_wb):
    wps, _sched = _read_wps_and_schedule(assigned_wb)
    first = _wp_summary_first_row(tlp)
    sheet_codes = []
    for r in range(first, first + len(wps) + 5):
        v = tlp.cell(r, 1).value
        if v is None or str(v).strip() == "":
            break
        sheet_codes.append(str(v).strip())
    assert sheet_codes == [wp["id"] for wp in wps]


# ── PER-WP SUMMARY: Buffer € ────────────────────────────────────────────────


def test_buffer_signed(tlp, assigned_wb):
    """Buffer € at col D = C − B (Assigned − Ceiling). CF red<0 / green>0."""
    wps, _sched = _read_wps_and_schedule(assigned_wb)
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


# ── PER-WP SUMMARY: WP sched months literal ─────────────────────────────────


def test_wp_sched_months_literal(tlp, assigned_wb):
    """Col F = literal int = sum of active-segment lengths (multi-segment WPs respected)."""
    wps, sched = _read_wps_and_schedule(assigned_wb)
    first = _wp_summary_first_row(tlp)
    for i, wp in enumerate(wps):
        r = first + i
        code = wp["id"]
        expected = sum(
            s["end"] - s["start"] + 1
            for s in sched
            if s["id"] == code
        )
        got = tlp.cell(r, 6).value
        assert got == expected, f"{code} sched months: expected {expected}, got {got}"

    # Bystander: the total across all WPs should match the schedule total,
    # proving the formula isn't always returning the same value.
    sheet_total = sum(tlp.cell(first + i, 6).value for i in range(len(wps)))
    sched_total = sum(s["end"] - s["start"] + 1 for s in sched)
    assert sheet_total == sched_total


# ── PER-WP SUMMARY: Burn rate ───────────────────────────────────────────────


def test_burn_rate_conditional(tlp, assigned_wb):
    """Col G formula: IF(AND(D<0, wallclock>0), -D/wallclock, ""). CF red on excess."""
    wps, _sched = _read_wps_and_schedule(assigned_wb)
    first = _wp_summary_first_row(tlp)
    last = first + len(wps) - 1
    for i, _wp in enumerate(wps):
        r = first + i
        v = str(tlp.cell(r, 7).value or "")
        assert v.startswith("=IF(AND(D"), f"row {r} Burn rate: {v!r}"
        assert f"D{r}<0" in v
        assert "current_month" in v
        assert '"")' in v or '"",' in v, v  # empty-string branch present

    # CF: red fill + MAX-rate comparison on a range including col G.
    red_argb = CF_RED_FILL.bgColor.rgb
    red_hits = 0
    for rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        if not rng.startswith("G") or ":" not in rng:
            continue
        if f"G{first}:G{last}" not in rng:
            continue
        fill = getattr(rule.dxf, "fill", None) if rule.dxf else None
        if fill is None or fill.bgColor is None:
            continue
        if fill.bgColor.rgb != red_argb:
            continue
        formula = (rule.formula or [""])[0]
        if "MAX(" in formula and "G" in formula:
            red_hits += 1
    assert red_hits >= 1, f"expected red Burn-rate CF on G-band, found {red_hits}"


# ── PER-PERSON SUMMARY: v2 schema ───────────────────────────────────────────


def test_person_summary_headers_v2(tlp):
    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY) + 1
    assert tlp.cell(hdr_row, 1).value == "Person"
    expected = [
        "Assigned", "Pool", "Over-pool?", "Switches",
        "First log", "Last log", "Largest gap", "Rate",
    ]
    got = [tlp.cell(hdr_row, 2 + i).value for i in range(len(expected))]
    assert got == expected, got


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
    """Year col count = distinct calendar years spanned; headers = Y{yyyy} €."""
    params, _wps, _sched = _read_config(assigned_wb)
    months = params["months"]
    groups = _expected_year_groups(params["start"], months)

    hdr_row = _find_section_title_row(tlp, SECTION_TITLE_WP_SUMMARY) + 1
    base_cols = 8  # cols A..H hold the base metrics
    for i, (year, _fm, _lm) in enumerate(groups):
        got = tlp.cell(hdr_row, base_cols + 1 + i).value
        assert got == f"Y{year} €", f"col {base_cols + 1 + i}: {got!r}"
    # No stray header past the last year col.
    trailing = tlp.cell(hdr_row, base_cols + len(groups) + 1).value
    assert trailing in (None, ""), f"unexpected header past year cols: {trailing!r}"


def _wp_active_window(wp_id: str, wp_schedule: list[dict]) -> tuple[int, int]:
    """Derive (active_start, active_end) from the schedule — multi-segment aware."""
    segs = [s for s in wp_schedule if s["id"] == wp_id]
    assert segs, f"no schedule rows for WP {wp_id!r}"
    return min(int(s["start"]) for s in segs), max(int(s["end"]) for s in segs)


def test_year_cols_overlap_blanks(tlp, assigned_wb):
    """Cells where WP's active window doesn't overlap the year are literal ``""``."""
    params, wps, sched = _read_config(assigned_wb)
    groups = _expected_year_groups(params["start"], params["months"])
    first = _wp_summary_first_row(tlp)
    base_cols = 8

    checked_nonoverlap = 0
    checked_overlap = 0
    for wi, wp in enumerate(wps):
        active_start, active_end = _wp_active_window(wp["id"], sched)
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
    """Sum across year cols on each WP row ≈ that WP's Assigned €.

    Each year cell is independently ROUND'd per the EUR leaf-rounding rule,
    so the sum of rounded-per-year can differ from the rounded-overall
    Assigned € by at most one EUR per year (round-half-even). We allow a
    tolerance of ``num_years`` EUR.
    """
    wb = load_workbook(recalced_assigned_path, data_only=True)
    ws = wb[SHEET_TIMELOG_PLAN]
    params, wps, _sched = _read_config(wb)
    groups = _expected_year_groups(params["start"], params["months"])
    first = _wp_summary_first_row(ws)
    base_cols = 8
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
        assert abs(year_sum - float(assigned_eur)) <= tolerance, (
            f"WP {wp['id']}: sum of year cols {year_sum} vs Assigned € "
            f"{assigned_eur} (tolerance ±{tolerance})"
        )
        checked += 1
    assert checked > 0, "no numeric Assigned € values — recalc didn't populate"
    wb.close()


def _read_config(wb):
    from budget_planner.reader import read_config_sheet
    return read_config_sheet(wb["Config"])
