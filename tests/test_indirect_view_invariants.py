"""Indirect-view cross-sheet invariants (see
``docs/spec.md`` § Indirect-view EUR header suffix).

Enforces the ``+i`` / ``-i`` convention across the in-scope sheets
(Config, Actuals, Capacity, Actuals-Hours, Timelog Plan):

* ``test_minus_i_fill_carries_minus_i_token`` — every cell with
  ``WITHOUT_INDIRECT_HEADER_FILL`` (#FBE5D6, unique to the indirect-view
  convention) contains a ``-i`` token in its text, or is an empty
  fill-continuation cell. The reverse ``WITH_INDIRECT_HEADER_FILL``
  fill→text check is NOT symmetric — #E2EFDA aliases
  ``LIGHT_GREEN_FILL``, which is reused on TOTAL rows and helper columns
  (e.g. ``Cum hours (to date)`` on Actuals-Hours). The text→fill
  direction covers the +i side.
* ``test_header_fill_matches_suffix`` — every bold header cell whose
  text carries a ``+i`` / ``-i`` token has the matching fill. Section
  titles (SECTION_FILL) and calendar-year anchor cells ("2026 -i" on the
  Actuals year_row) are exempt per D8.
* ``test_budget_plus_i_minus_i_round_trip`` — per WP, Config
  ``Budget +i / 1.25 == Budget -i`` after LibreOffice recalc.

Trends, Summary, and Plan are intentionally out of scope — they will be
aligned in the follow-on overhaul plan noted in
``memory/project_state.md``.
"""

import re

import openpyxl

from budget_planner.layout import (
    COL_WP_BUDGET_MINUS_I,
    COL_WP_BUDGET_PLUS_I,
    SHEET_ACTUALS,
    SHEET_ACTUALS_HOURS,
    SHEET_CAPACITY,
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
)
from budget_planner.styles import (
    LIGHT_BLUE_FILL,
    SECTION_FILL,
    WITH_INDIRECT_HEADER_FILL,
    WITHOUT_INDIRECT_HEADER_FILL,
)

IN_SCOPE_SHEETS = [
    SHEET_CONFIG,
    SHEET_ACTUALS,
    SHEET_ACTUALS_HOURS,
    SHEET_CAPACITY,
    SHEET_TIMELOG_PLAN,
]


def _rgb6(color):
    rgb = getattr(color, "rgb", None)
    if not isinstance(rgb, str):
        return None
    s = rgb.upper()
    return s[-6:] if len(s) >= 6 else None


PLUS_I_RGB = _rgb6(WITH_INDIRECT_HEADER_FILL.fgColor)
MINUS_I_RGB = _rgb6(WITHOUT_INDIRECT_HEADER_FILL.fgColor)
SECTION_RGB = _rgb6(SECTION_FILL.fgColor)
LIGHT_BLUE_RGB = _rgb6(LIGHT_BLUE_FILL.fgColor)


def _cell_rgb(cell):
    fill = cell.fill
    if fill is None or fill.patternType != "solid":
        return None
    return _rgb6(fill.fgColor)


# Calendar-year anchor: the year_row cells on Actuals carry a
# "<YYYY> -i" label that visually labels the monthly grid band.
# The grid cells themselves and this label use LIGHT_BLUE_FILL (inherited
# from ``build_year_month_headers``). Exempt per D8.
_CALENDAR_ANCHOR_RE = re.compile(r"^\d{4}\s[+\-]i$")

# Decorative band-mirror labels — calendar years ("2026") and month
# tags ("M12-Oct27") mirrored onto the Actuals-Hours forecast-EUR band
# by ``_mirror_year_month_headers_to_band``. They inherit the band's
# WITHOUT_INDIRECT fill but don't themselves carry the -i token; the
# band-label row above ("FORECAST EUR -i (from hours)") does.
_BAND_MIRROR_LABEL_RE = re.compile(r"^(?:\d{4}|M\d+-[A-Za-z]{3}\d{2})$")

# Indirect-view token: "+i" or "-i" preceded by whitespace / start, and
# followed by a non-word char or end of string.
_INDIRECT_TOKEN_RE = re.compile(r"(?:^|\s)([+\-]i)(?=\W|$)")


def test_minus_i_fill_carries_minus_i_token(assigned_wb_path):
    """Every cell on the in-scope sheets with WITHOUT_INDIRECT_HEADER_FILL
    (unique orange #FBE5D6) contains a '-i' token in its text."""
    wb = openpyxl.load_workbook(assigned_wb_path, data_only=False)
    offenders = []
    for sheet_name in IN_SCOPE_SHEETS:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if _cell_rgb(cell) != MINUS_I_RGB:
                    continue
                text = cell.value
                if not isinstance(text, str) or not text.strip():
                    continue
                if _BAND_MIRROR_LABEL_RE.match(text.strip()):
                    continue
                if "-i" not in text:
                    offenders.append(f"{sheet_name}!{cell.coordinate}={text!r} has -i fill but no -i token")
    assert not offenders, "-i fill without -i token:\n" + "\n".join(offenders)


def test_header_fill_matches_suffix(assigned_wb_path):
    """Every header cell whose text carries a +i / -i token on the
    in-scope sheets has the matching indirect-view fill. Exempts
    section-title cells (SECTION_FILL) and calendar-year anchor cells."""
    wb = openpyxl.load_workbook(assigned_wb_path, data_only=False)
    offenders = []
    for sheet_name in IN_SCOPE_SHEETS:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                text = cell.value
                if not isinstance(text, str) or not text.strip():
                    continue
                # Formula cells can embed " -i" inside structured-ref
                # brackets like ``[Rate (€/PM) -i]`` — that's a WP Budget
                # column reference, not a header suffix. Skip formulas.
                if text.startswith("="):
                    continue
                rgb = _cell_rgb(cell)
                if rgb == SECTION_RGB:
                    continue
                if _CALENDAR_ANCHOR_RE.match(text.strip()):
                    continue
                tokens = set(_INDIRECT_TOKEN_RE.findall(text))
                if "+i" in tokens and rgb != PLUS_I_RGB:
                    offenders.append(
                        f"{sheet_name}!{cell.coordinate}={text!r} carries +i "
                        f"token but fill rgb={rgb!r} (expected {PLUS_I_RGB!r})"
                    )
                elif "-i" in tokens and rgb != MINUS_I_RGB:
                    offenders.append(
                        f"{sheet_name}!{cell.coordinate}={text!r} carries -i "
                        f"token but fill rgb={rgb!r} (expected {MINUS_I_RGB!r})"
                    )
    assert not offenders, "suffix→fill mismatches:\n" + "\n".join(offenders)


def test_budget_plus_i_minus_i_round_trip(demo_recalc_path):
    """Per WP on Config, Budget +i / 1.25 ≈ Budget -i after recalc.

    Skips implicitly when LibreOffice is unavailable (demo_recalc_path
    chains through demo_recalc, which pytest.skip()s on LO missing).
    """
    wb = openpyxl.load_workbook(demo_recalc_path, data_only=True)
    ws = wb[SHEET_CONFIG]

    header_row = None
    for r in range(1, 60):
        v = ws.cell(r, COL_WP_BUDGET_MINUS_I).value
        if isinstance(v, str) and v.strip() == "Budget -i":
            header_row = r
            break
    assert header_row is not None, f"Could not locate 'Budget -i' header on Config col {COL_WP_BUDGET_MINUS_I}"

    plus_i_header = ws.cell(header_row, COL_WP_BUDGET_PLUS_I).value
    assert isinstance(plus_i_header, str) and plus_i_header.strip() == "Budget +i", (
        f"Expected 'Budget +i' header at col {COL_WP_BUDGET_PLUS_I}, got {plus_i_header!r}"
    )

    failures = []
    for r in range(header_row + 1, header_row + 100):
        wp_id = ws.cell(r, 1).value
        if wp_id in (None, "", "TOTAL"):
            break
        minus_i = ws.cell(r, COL_WP_BUDGET_MINUS_I).value
        plus_i = ws.cell(r, COL_WP_BUDGET_PLUS_I).value
        if not isinstance(minus_i, (int, float)) or not isinstance(plus_i, (int, float)):
            failures.append(
                f"WP {wp_id!r} row {r}: non-numeric Budget (-i={minus_i!r}, +i={plus_i!r}) — recalc incomplete?"
            )
            continue
        expected = plus_i / 1.25
        if abs(expected - minus_i) > 1e-6:
            failures.append(f"WP {wp_id!r}: +i {plus_i} / 1.25 = {expected} != -i {minus_i}")
    assert failures == [], "Budget +i/-i round-trip failures:\n" + "\n".join(failures)
