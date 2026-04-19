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
    wb.close()


def test_through_month_overrides_round_trip(tmp_path: Path):
    """User-typed overrides in Actuals!B1 and Actuals-Hours!B1 survive rebuild.

    Post-2026-04-18 cleanup: through-month control lives on the Actuals /
    Actuals-Hours sheets (row 1 = label / blue-input override / auto /
    effective). ``read_actuals_override`` extracts B1 and threads it through
    ``_build_all_sheets`` → builders — second rebuild stability included.
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    # ── User edits row-1 B1 on each Actuals sheet ──
    wb = openpyxl.load_workbook(path)
    wb["Actuals"]["B1"] = 3
    wb["Actuals-Hours"]["B1"] = 5
    wb.save(path)
    wb.close()

    # ── First rebuild ──
    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    assert wb["Actuals"]["B1"].value == 3, (
        f"Rebuild clobbered eur through-month override: "
        f"expected 3, got {wb['Actuals']['B1'].value!r}"
    )
    assert wb["Actuals-Hours"]["B1"].value == 5, (
        f"Rebuild clobbered hours through-month override: "
        f"expected 5, got {wb['Actuals-Hours']['B1'].value!r}"
    )

    # Named ranges still point at the effective cells (D1) on each sheet.
    assert wb.defined_names["eur_actuals_through_month"].attr_text == (
        "Actuals!$D$1"
    )
    assert wb.defined_names["hours_actuals_through_month"].attr_text == (
        "'Actuals-Hours'!$D$1"
    )
    wb.close()

    # ── Second rebuild: stability ──
    rebuild_workbook(str(path))
    wb = openpyxl.load_workbook(path)
    assert wb["Actuals"]["B1"].value == 3
    assert wb["Actuals-Hours"]["B1"].value == 5
    # Bystander: an untouched Config param (Hours per FTE Month, B6=143) stays
    # — proves rebuild isn't wholesale-preserving workbook state.
    assert wb["Config"]["B6"].value == 143
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
    wb.close()
