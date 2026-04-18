"""Structural tests for the Timelog Plan sheet builder — vertical-stacked layout.

Covers the shape a downstream consumer can rely on without recalculating the
workbook: sheet exists, person rows mirror the roster in each band, the
PER-WP and PER-PERSON SUMMARY blocks are populated, traffic-light formulas
are wired, and CF rules match the design spec (matched by fill colour +
formula pattern — not positional index, per ``openpyxl-patterns.md``).
"""
from __future__ import annotations

import re

import pytest
from openpyxl import load_workbook

from budget_planner.config import TEAM, WORK_PACKAGES
from budget_planner.layout import (
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
    TIMELOG_RATE_HELPER_COL,
)
from budget_planner.styles import CF_RED_FILL
from budget_planner.timelog_plan import (
    SECTION_TITLE_ASSIGNED,
    SECTION_TITLE_PERSON_SUMMARY,
    SECTION_TITLE_PROPOSED,
    SECTION_TITLE_WP_SUMMARY,
)


# ── Fixtures ────────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def assigned_wb(assigned_wb_path):
    return load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def tlp(assigned_wb):
    return assigned_wb[SHEET_TIMELOG_PLAN]


def _project_months(wb) -> int:
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _find_section_title_row(ws, title: str) -> int:
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == title:
            return r
    raise AssertionError(f"Section title {title!r} not found in col A")


def _band_person_first_row(ws, title: str) -> int:
    """First person row of a band (ASSIGNED / PROPOSED) = title + 4.

    Band layout: title, year, month, midx(hidden), persons...
    """
    return _find_section_title_row(ws, title) + 4


def _summary_person_first_row(ws) -> int:
    """PER-PERSON SUMMARY band: title, header, persons → title + 2."""
    return _find_section_title_row(ws, SECTION_TITLE_PERSON_SUMMARY) + 2


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


# ── Vertical-band layout ────────────────────────────────────────────────────


def test_all_four_section_titles_present(tlp):
    """PER-WP SUMMARY, PROPOSED, ASSIGNED, PER-PERSON SUMMARY all on col A."""
    for title in (
        SECTION_TITLE_WP_SUMMARY,
        SECTION_TITLE_PROPOSED,
        SECTION_TITLE_ASSIGNED,
        SECTION_TITLE_PERSON_SUMMARY,
    ):
        _find_section_title_row(tlp, title)  # raises if missing


def test_section_title_order_top_to_bottom(tlp):
    """PER-WP SUMMARY < PROPOSED < ASSIGNED < PER-PERSON SUMMARY by row."""
    rows = [
        _find_section_title_row(tlp, SECTION_TITLE_WP_SUMMARY),
        _find_section_title_row(tlp, SECTION_TITLE_PROPOSED),
        _find_section_title_row(tlp, SECTION_TITLE_ASSIGNED),
        _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY),
    ]
    assert rows == sorted(rows), f"sections out of order: {rows}"


# ── Shape ───────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("title", [SECTION_TITLE_PROPOSED, SECTION_TITLE_ASSIGNED])
def test_person_row_count_matches_team_per_band(tlp, title):
    first = _band_person_first_row(tlp, title)
    rows = list(_iter_person_rows(tlp, first))
    assert len(rows) == len(TEAM), f"{title}: rows={len(rows)}, team={len(TEAM)}"


def test_summary_band_person_row_count(tlp):
    first = _summary_person_first_row(tlp)
    rows = list(_iter_person_rows(tlp, first))
    assert len(rows) == len(TEAM)


@pytest.mark.parametrize("title", [SECTION_TITLE_PROPOSED, SECTION_TITLE_ASSIGNED])
def test_person_col_a_cross_refs_config_per_band(tlp, title):
    """Every person row's col A is a ``=Config!A<n>`` cross-reference."""
    first = _band_person_first_row(tlp, title)
    for r in _iter_person_rows(tlp, first):
        v = tlp.cell(r, 1).value
        assert isinstance(v, str) and re.match(r"^=Config!A\d+$", v), (
            f"{title}: row {r} col A should be =Config!A<n>, got {v!r}"
        )


def test_assigned_grid_defaults_to_proposed_formula(tlp):
    """Fresh build — every Assigned cell points at the Proposed band's col/row.

    Formula shape: ``=<ColLetter>{proposed_row}`` where ``<ColLetter>`` is the
    same column in the Assigned grid's row (the two grids share columns).
    """
    months = _project_months(tlp.parent)
    m_start = TIMELOG_MONTH_COL_START
    m_end = m_start + months - 1
    assigned_first = _band_person_first_row(tlp, SECTION_TITLE_ASSIGNED)
    proposed_first = _band_person_first_row(tlp, SECTION_TITLE_PROPOSED)
    sampled = 0
    for i, r in enumerate(_iter_person_rows(tlp, assigned_first)):
        proposed_row = proposed_first + i
        for c in range(m_start, m_end + 1):
            v = tlp.cell(r, c).value
            if v is None or v == "":
                continue
            # Every non-empty cell must be a cross-band ref into PROPOSED.
            expected = f"={tlp.cell(r, c).column_letter}{proposed_row}"
            assert v == expected, (
                f"row {r} col {c}: expected {expected!r}, got {v!r}"
            )
            sampled += 1
    assert sampled > 0, "Assigned grid had no populated cells"


# ── PER-WP SUMMARY band ─────────────────────────────────────────────────────


def _wp_summary_first_row(ws) -> int:
    """PER-WP SUMMARY band: title, header, WPs → title + 2."""
    return _find_section_title_row(ws, SECTION_TITLE_WP_SUMMARY) + 2


def test_wp_summary_rows_match_work_packages(tlp):
    first = _wp_summary_first_row(tlp)
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
    first = _wp_summary_first_row(tlp)
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


# ── PER-PERSON SUMMARY band ─────────────────────────────────────────────────


def test_person_summary_headers(tlp):
    """Header row: col A=Person, B=Underload, C-I=7 metrics (Remaining dropped)."""
    summary_title = _find_section_title_row(tlp, SECTION_TITLE_PERSON_SUMMARY)
    hdr_row = summary_title + 1
    # Col A and B headers.
    assert tlp.cell(hdr_row, 1).value == "Person"
    assert tlp.cell(hdr_row, 2).value == "Underload"
    # Metric headers C..I (7 cols).
    expected = [
        "Assigned", "Pool", "Over-pool?",
        "Switches", "First log", "Last log", "Largest gap",
    ]
    got = [tlp.cell(hdr_row, 3 + i).value for i in range(len(expected))]
    assert got == expected


def test_person_summary_formulas(tlp):
    """Metric cols C..I and Underload col B carry the expected formula shapes."""
    row = _summary_person_first_row(tlp)  # first person in summary band
    underload = str(tlp.cell(row, 2).value or "")
    assigned_count = str(tlp.cell(row, 3).value or "")    # C: Assigned
    pool = str(tlp.cell(row, 4).value or "")              # D: Pool
    over_pool = str(tlp.cell(row, 5).value or "")         # E: Over-pool?
    switches = str(tlp.cell(row, 6).value or "")          # F: Switches
    first_log = str(tlp.cell(row, 7).value or "")         # G: First log
    # Underload = Pool(D) − Assigned(C) same row.
    assert underload.startswith("=D") and "-C" in underload, underload
    assert assigned_count.startswith("=COUNTIF(")
    assert "ROUND(" in pool and "Config!" in pool
    assert over_pool.startswith("=C") and ">D" in over_pool
    assert switches.startswith("=SUMPRODUCT(1*(")
    # First-log uses the robust (range<>"")*1 coercion pattern, not the
    # silent-CSE-only MATCH(TRUE,...) — verify we didn't drift.
    assert '<>""' in first_log and "*1,0" in first_log


@pytest.mark.parametrize("title", [SECTION_TITLE_PROPOSED, SECTION_TITLE_ASSIGNED])
def test_underload_col_b_in_data_bands(tlp, title):
    """PROPOSED and ASSIGNED bands: col B = $D{summary_row} − COUNTIF(row, "WP*")."""
    first = _band_person_first_row(tlp, title)
    first_row = next(iter(_iter_person_rows(tlp, first)))
    val = str(tlp.cell(first_row, 2).value or "")
    # Pool lives in col D of the summary band — absolute col ref, relative row.
    assert val.startswith("=$D"), val
    assert 'COUNTIF(' in val and '"WP*"' in val, val


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
            yield str(cf_range.sqref), rule


def test_cf_has_red_rule_for_every_wp_out_of_window(tlp):
    """One red-fill CF per WP code × (proposed, assigned) grid."""
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
    """One pastel-fill CF per WP × (proposed, assigned) — 2 per WP total."""
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


def test_cf_underload_font_rules_per_band(tlp):
    """Underload col B carries font CF in all 3 bands (red < 0, green > 0).

    6 rules total (3 bands × 2). Match by formula shape ("B{r}<0" or "B{r}>0")
    and ignore any fill-based rules (pastel / OOW red live on the month grid,
    not col B).
    """
    lt_rules = 0
    gt_rules = 0
    for rng, rule in _iter_all_rules(tlp):
        if rule.type != "expression":
            continue
        # Underload CF lives on multi-row B<first>:B<last> ranges (one per
        # band). Single-cell B<row> rules are the dashboard indicators at
        # the top of the sheet — different purpose, same column.
        if not (rng.startswith("B") and ":" in rng):
            continue
        formula = (rule.formula or [""])[0]
        if re.match(r"^B\d+<0$", formula):
            lt_rules += 1
        elif re.match(r"^B\d+>0$", formula):
            gt_rules += 1
    assert lt_rules == 3, f"expected 3 'B<0' underload rules, got {lt_rules}"
    assert gt_rules == 3, f"expected 3 'B>0' underload rules, got {gt_rules}"


# ── Outline grouping ────────────────────────────────────────────────────────


def test_band_rows_are_outline_level_1(tlp):
    """Data-band rows (headers + data) collapse under outline_level=1;
    section-title rows stay at level 0."""
    proposed_title = _find_section_title_row(tlp, SECTION_TITLE_PROPOSED)
    proposed_year = proposed_title + 1
    proposed_first_person = proposed_title + 4
    # Section title: level 0 (stays visible when collapsed).
    assert tlp.row_dimensions[proposed_title].outline_level in (None, 0)
    # Year row: level 1.
    assert tlp.row_dimensions[proposed_year].outline_level == 1
    # First person row: level 1.
    assert tlp.row_dimensions[proposed_first_person].outline_level == 1


def test_rate_helper_col_c_collapsible(tlp):
    """Col C (rate helper) is collapsible via outline_level=1."""
    assert tlp.column_dimensions["C"].outline_level == 1


# ── Freeze + rate-helper column coordinates ─────────────────────────────────


def test_freeze_panes_at_d3(tlp):
    """Freeze = "D3" — rows 1-2 + cols A-C stay visible when scrolling."""
    expected_col_letter = tlp.cell(3, TIMELOG_MONTH_COL_START).column_letter
    assert tlp.freeze_panes == f"{expected_col_letter}3"
    assert tlp.freeze_panes == "D3"


def test_rate_helper_col_is_c(tlp):
    """After the vertical restructure, rate helper lives in col C (not B)."""
    assert TIMELOG_RATE_HELPER_COL == 3
    # Data-band first person row: col C carries a VLOOKUP(A..., roster_lookup...).
    first = _band_person_first_row(tlp, SECTION_TITLE_ASSIGNED)
    first_row = next(iter(_iter_person_rows(tlp, first)))
    rate_val = str(tlp.cell(first_row, TIMELOG_RATE_HELPER_COL).value or "")
    assert rate_val.startswith("=VLOOKUP(") and "roster_lookup" in rate_val, rate_val
