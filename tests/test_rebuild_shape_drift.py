"""Round-trip tests for Excel-side row insertions in Config tables.

Mirrors the user's workflow: generate a workbook, hand-edit Config in Excel
(simulated via openpyxl), run rebuild, assert no data loss.
"""
from pathlib import Path

import openpyxl

from budget_planner import config as cfg
from budget_planner.reader import (
    read_assignments_from_config,
    read_team_from_config,
)
from budget_planner.workbook import init_workbook, rebuild_workbook
from tests._shape_invariants import assert_meta_matches_live


def _find_roster_first_data_row(ws):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == "TEAM ROSTER":
            return r + 2
    raise AssertionError("TEAM ROSTER section not found")


def _simulate_excel_insert_in_named_range(wb, name: str, before_row: int, amount: int = 1):
    """Simulate Excel's row-insert semantics: extend the named range's last row.

    openpyxl's ``ws.insert_rows`` shifts cells but does not update defined-name
    references. Excel, by contrast, auto-extends a named range when rows are
    inserted inside it. We emulate that here so the post-edit workbook matches
    what the user's Excel session produced.

    Tolerates missing names so callers can iterate a fixed list of candidate
    ranges without guarding each one.
    """
    import re

    if name not in wb.defined_names:
        return
    dn = wb.defined_names[name]
    dest = list(dn.destinations)[0]
    sheet, ref = dest
    m = re.match(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", ref)
    assert m, f"unexpected ref form: {ref}"
    col1, row1, col2, row2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    last = int(row2) if row2 else row1
    # Extend only when the insertion is strictly inside [row1, last]. Excel
    # does NOT auto-extend when inserting at last+1 (immediately below the
    # range); that's an append and leaves the name untouched.
    # Inserts above row1 are no-ops here (Excel behaves the same way).
    if row1 <= before_row <= last:
        new_last = last + amount
        col2_part = f":${col2}${new_last}" if col2 else f":${col1}${new_last}"
        dn.attr_text = f"{sheet}!${col1}${row1}{col2_part}"


def test_blank_row_mid_roster_does_not_truncate_rebuild(tmp_path: Path):
    """User inserts an empty row mid-roster (common in Excel); rebuild must keep persons after the blank."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    first_data = _find_roster_first_data_row(config_ws)
    # Capture existing roster names for later assertion.
    existing_names: list[str] = []
    r = first_data
    while r <= config_ws.max_row:
        v = config_ws.cell(r, 1).value
        if v is None or (isinstance(v, str) and not v.strip()):
            break
        existing_names.append(str(v).strip())
        r += 1
    assert len(existing_names) >= 2, "seed roster must have at least 2 people for this test"

    # Insert a truly blank row in the middle of the roster.
    insert_at = first_data + 1
    config_ws.insert_rows(insert_at, amount=1)
    _simulate_excel_insert_in_named_range(wb, "roster_ftes", insert_at, 1)
    _simulate_excel_insert_in_named_range(wb, "roster_grades", insert_at, 1)
    _simulate_excel_insert_in_named_range(wb, "roster_lookup", insert_at, 1)
    wb.save(path)
    wb.close()

    # Pre-rebuild reader must see ALL original persons despite the blank row.
    wb = openpyxl.load_workbook(path)
    roster = read_team_from_config(wb["Config"])
    wb.close()
    names = [p["name"] for p in roster]
    for n in existing_names:
        assert n in names, (
            f"Pre-rebuild read dropped person {n!r} after blank-row insertion. "
            f"Roster read: {names}"
        )

    # Rebuild must succeed and preserve all persons.
    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    roster_after = read_team_from_config(wb["Config"])
    assignments_after = read_assignments_from_config(wb["Config"])
    names_after = [p["name"] for p in roster_after]
    for n in existing_names:
        assert n in names_after, (
            f"Rebuild dropped person {n!r} across blank-row insertion. "
            f"Roster after rebuild: {names_after}"
        )
    dropped_assignees = {a["person"] for a in assignments_after} - set(names_after)
    assert not dropped_assignees, (
        f"Rebuild has assignments pointing at persons no longer in roster: {dropped_assignees}"
    )
    assert_meta_matches_live(wb)
    wb.close()


def test_hours_actuals_through_month_round_trip(tmp_path: Path):
    """T2: user-edited hours_actuals_through_month survives rebuild.

    Covers the full init -> user edit -> rebuild -> reopen flow for the new
    v2 Config rows (B8 blue input + B9-B11 derived formulas + four named
    ranges). Also bumps B7 (eur_actuals_through_month) so the MIN in B11 has
    something non-trivial to resolve. Second rebuild verifies stability
    (the preserved value survives when it is itself the starting point).
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    # ── Simulate user edit in Excel: set B7=3 and B8=5 ──
    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    config_ws["B7"] = 3  # eur_actuals_through_month
    config_ws["B8"] = 5  # hours_actuals_through_month
    wb.save(path)
    wb.close()

    # ── First rebuild ──
    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]

    # (a) User value preserved in B8.
    assert config_ws["B8"].value == 5, (
        f"Rebuild clobbered hours_actuals_through_month: "
        f"expected 5, got {config_ws['B8'].value!r}"
    )

    # (b) Derived formulas still exact-string match in B9-B11.
    assert config_ws["B9"].value == "=MAX(1,$B$8)", (
        f"B9 derived formula drift: {config_ws['B9'].value!r}"
    )
    assert config_ws["B10"].value == "=MAX(1,$B$7)", (
        f"B10 derived formula drift: {config_ws['B10'].value!r}"
    )
    assert config_ws["B11"].value == "=MIN($B$7,$B$8)", (
        f"B11 derived formula drift: {config_ws['B11'].value!r}"
    )

    # (c) All four named ranges still point at Config!$B$8..$B$11.
    expected_targets = {
        "hours_actuals_through_month": "Config!$B$8",
        "months_elapsed_hours": "Config!$B$9",
        "months_elapsed_eur": "Config!$B$10",
        "reconciled_through_month": "Config!$B$11",
    }
    for name, expected in expected_targets.items():
        assert name in wb.defined_names, (
            f"Named range {name!r} lost across rebuild"
        )
        assert wb.defined_names[name].attr_text == expected, (
            f"Named range {name!r} drifted: "
            f"expected {expected!r}, got {wb.defined_names[name].attr_text!r}"
        )

    wb.close()

    # ── Second rebuild: stability check ──
    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    assert config_ws["B8"].value == 5, (
        f"Second rebuild clobbered preserved hours_actuals_through_month: "
        f"expected 5, got {config_ws['B8'].value!r}"
    )
    # Derived formulas still present (their values would be filled by recalc,
    # but on a freshly rebuilt workbook openpyxl sees the formula string).
    assert config_ws["B9"].value == "=MAX(1,$B$8)"
    # Bystander: an untouched param row (Hours per FTE Month, B6=143) stays put —
    # proves the preservation is selective, not wholesale.
    assert config_ws["B6"].value == 143, (
        f"Hours per FTE Month (B6) drifted unexpectedly: {config_ws['B6'].value!r}"
    )
    wb.close()


def test_mid_insert_in_team_roster_preserved_through_rebuild(tmp_path: Path):
    """User inserts a new person row in the middle of Team Roster; rebuild must keep them."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    first_data = _find_roster_first_data_row(config_ws)
    insert_at = first_data + 1
    config_ws.insert_rows(insert_at, amount=1)
    config_ws.cell(insert_at, 1, "Midway Mary")
    config_ws.cell(insert_at, 2, 0.8)
    config_ws.cell(insert_at, 3, next(iter(cfg.GRADE_RATES.keys())))
    _simulate_excel_insert_in_named_range(wb, "roster_ftes", insert_at, 1)
    _simulate_excel_insert_in_named_range(wb, "roster_grades", insert_at, 1)
    _simulate_excel_insert_in_named_range(wb, "roster_lookup", insert_at, 1)
    wb.save(path)
    wb.close()

    wb = openpyxl.load_workbook(path)
    roster = read_team_from_config(wb["Config"])
    wb.close()
    names = [p["name"] for p in roster]
    assert "Midway Mary" in names, (
        f"Pre-rebuild read already loses inserted person — reader bug. "
        f"Roster read: {names}"
    )

    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    roster_after = read_team_from_config(wb["Config"])
    assignments_after = read_assignments_from_config(wb["Config"])
    names_after = [p["name"] for p in roster_after]
    assert "Midway Mary" in names_after, (
        f"Rebuild dropped the mid-inserted person. Roster after rebuild: {names_after}"
    )
    mary = next(p for p in roster_after if p["name"] == "Midway Mary")
    assert mary["grade"] == next(iter(cfg.GRADE_RATES.keys())), (
        f"Rebuild dropped the mid-inserted person's grade. Got {mary['grade']!r}"
    )
    dropped_assignees = {a["person"] for a in assignments_after} - set(names_after)
    assert not dropped_assignees, (
        f"Rebuild has assignments pointing at persons no longer in roster: {dropped_assignees}"
    )
    assert_meta_matches_live(wb)
    wb.close()
