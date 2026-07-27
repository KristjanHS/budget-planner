"""Follow-up coverage for rebuild paths not exercised by earlier tests.

- Config-absent fallback: when the Config sheet lacks the PERSON-WP
  ASSIGNMENTS / EXTERNAL COMMITMENTS section headers (simulating a
  pre-Phase-1 workbook), the rebuild falls back to the legacy Capacity
  sections via ``read_capacity_sheet``.
- Actuals-Hours orphan pruning: when the preserved hours dict contains
  a person who is no longer on the roster (or a WP that has been
  removed), the pruning block in ``rebuild_workbook`` drops those
  entries and emits an "Orphaned Hours" warning instead of crashing.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl
from openpyxl.cell import Cell

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import MONTH_COL_START
from budget_planner.reader import (
    _find_header_row,
    read_actuals_hours_sheet,
    read_assignments_from_config,
    read_externals_from_config,
)
from budget_planner.workbook import init_workbook, rebuild_workbook


def _set(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _blank_config_section_headers(ws, labels: list[str]) -> None:
    """Clear col A cells whose value matches any of ``labels`` so that
    ``_find_header_row(required=False)`` returns None."""
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None and str(val).strip() in labels:
            _set(ws, r, 1, None)


def _write_legacy_assignments_on_capacity(ws, assignments: list[dict], start_row: int) -> None:
    """Emulate a pre-Phase-1 workbook by writing a PERSON-WP ASSIGNMENTS
    config block on the Capacity sheet. Layout mirrors what the legacy
    reader expects: header row, column-header row, then data rows in
    ``(person, wp, total_fte, start, end)`` order (cols 1-5).
    """
    _set(ws, start_row, 1, "PERSON-WP ASSIGNMENTS")
    _set(ws, start_row + 1, 1, "Person")
    _set(ws, start_row + 1, 2, "WP")
    _set(ws, start_row + 1, 3, "Total FTE")
    _set(ws, start_row + 1, 4, "Start")
    _set(ws, start_row + 1, 5, "End")
    for i, a in enumerate(assignments, start=start_row + 2):
        _set(ws, i, 1, a["person"])
        _set(ws, i, 2, a["wp"])
        _set(ws, i, 3, a["total_fte"])
        _set(ws, i, 4, a["start"])
        _set(ws, i, 5, a["end"])


def test_rebuild_falls_back_to_capacity_when_config_sections_absent(
    tmp_path: Path,
):
    """Blanking the Config assignment/external section headers triggers the
    Capacity-based legacy path; rebuild then repopulates the Config sections
    from the fallback data."""
    path = str(tmp_path / "legacy.xlsx")
    init_workbook(output_path=path, assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Surrogate legacy scenario: Config sections gone, Capacity carries data.
    wb = openpyxl.load_workbook(path)
    _blank_config_section_headers(wb["Config"], ["PERSON-WP ASSIGNMENTS", "EXTERNAL COMMITMENTS"])
    # Write a legacy-style assignments block on Capacity below existing content.
    # Use a subset (first 2) so we can verify the fallback actually parses it.
    legacy_subset = DEMO_CAPACITY_ASSIGNMENTS[:2]
    _write_legacy_assignments_on_capacity(wb["Capacity"], legacy_subset, start_row=wb["Capacity"].max_row + 3)
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=path)

    # After rebuild, Config sections must exist again and contain the legacy
    # data that the fallback recovered from Capacity.
    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    assert _find_header_row(config_ws, "PERSON-WP ASSIGNMENTS", required=False) is not None, (
        "Config PERSON-WP ASSIGNMENTS header missing after rebuild"
    )
    assert _find_header_row(config_ws, "EXTERNAL COMMITMENTS", required=False) is not None, (
        "Config EXTERNAL COMMITMENTS header missing after rebuild"
    )

    rebuilt = read_assignments_from_config(config_ws)
    # Externals on Capacity were untouched in this test — externals section
    # on Config is expected to be empty (fallback returned []).
    assert read_externals_from_config(config_ws) == []
    wb.close()

    recovered = {(a["person"], a["wp"]) for a in rebuilt}
    expected = {(a["person"], a["wp"]) for a in legacy_subset}
    assert expected <= recovered, f"Fallback failed to recover legacy assignments: expected {expected}, got {recovered}"


def _find_hours_wp_row(ws, config_ws, person: str, wp_id: str) -> int:
    """Walk Actuals-Hours rows to find the (person, wp) data row.

    Person headers carry the literal name in col A (not a formula); WP
    rows below carry ``=Config!A{n}`` which resolves to the WP ID in col
    A of the Config sheet.
    """
    current_person_row: int | None = None
    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(r, 1).value
        if isinstance(a_val, str) and not a_val.startswith("="):
            if a_val.strip() == person:
                current_person_row = r
            continue
        if current_person_row is None:
            continue
        if isinstance(a_val, str) and a_val.startswith("="):
            m = re.match(r"=Config!A(\d+)", a_val)
            if m:
                resolved = config_ws.cell(int(m.group(1)), 1).value
                if resolved and str(resolved).strip() == wp_id:
                    return r
    raise AssertionError(f"Could not locate Actuals-Hours row for {person!r}/{wp_id!r}")


def test_rebuild_prunes_orphaned_hours_when_person_removed(tmp_path: Path, capsys):
    """Hours logged against a person later removed from the roster are
    dropped on rebuild with an 'Orphaned Hours' warning; hours logged
    for a bystander person with no roster changes survive the rebuild."""
    path = str(tmp_path / "orphan.xlsx")
    init_workbook(output_path=path, assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Pick a target (to be orphaned) and a bystander (must survive pruning).
    target_asgn = DEMO_CAPACITY_ASSIGNMENTS[0]
    target_person = target_asgn["person"]
    bystander_asgn = next(
        (a for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] != target_person),
        None,
    )
    assert bystander_asgn is not None, "demo data must contain at least two distinct assigned people"

    target_hours = 96.0
    bystander_hours = 48.5

    wb = openpyxl.load_workbook(path)
    ah = wb["Actuals-Hours"]
    config_ws = wb["Config"]

    target_row = _find_hours_wp_row(ah, config_ws, target_person, target_asgn["wp"])
    bystander_row = _find_hours_wp_row(ah, config_ws, bystander_asgn["person"], bystander_asgn["wp"])
    _set(ah, target_row, MONTH_COL_START + target_asgn["start"] - 1, target_hours)
    bystander_month_idx = bystander_asgn["start"] - 1
    _set(
        ah,
        bystander_row,
        MONTH_COL_START + bystander_month_idx,
        bystander_hours,
    )

    # Remove the target from the Config roster + their assignment rows so
    # team_names drops them and _validate_config doesn't reject the rebuild.
    roster_section = _find_header_row(config_ws, "TEAM ROSTER", required=True)
    assert roster_section is not None
    removed = False
    for r in range(roster_section + 2, config_ws.max_row + 1):
        if config_ws.cell(r, 1).value == target_person:
            _set(config_ws, r, 1, None)
            _set(config_ws, r, 2, None)
            removed = True
            break
    assert removed, f"{target_person} not found in Config TEAM ROSTER"

    assign_section = _find_header_row(config_ws, "PERSON-WP ASSIGNMENTS", required=False)
    assert assign_section is not None
    for r in range(assign_section + 2, config_ws.max_row + 1):
        if config_ws.cell(r, 1).value == target_person:
            _set(config_ws, r, 1, None)
            _set(config_ws, r, 3, None)

    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=path)
    captured = capsys.readouterr()
    assert "Orphaned Hours" in captured.out, f"Expected 'Orphaned Hours' warning on stdout; got: {captured.out!r}"
    assert target_person in captured.out

    wb = openpyxl.load_workbook(path)
    hours_data = read_actuals_hours_sheet(wb["Actuals-Hours"])
    wb.close()

    # Target was pruned; bystander survived with the sentinel value intact.
    # The bystander check is the real correctness gate — without it, the test
    # would pass even if pruning dropped every person.
    assert target_person not in hours_data, (
        f"{target_person}'s hours should be pruned as orphaned; still present in: {list(hours_data.keys())}"
    )
    bystander = bystander_asgn["person"]
    assert bystander in hours_data, f"bystander {bystander!r} missing after rebuild: {list(hours_data.keys())}"
    assert bystander_asgn["wp"] in hours_data[bystander]
    got = hours_data[bystander][bystander_asgn["wp"]][bystander_month_idx]
    assert got == bystander_hours, f"bystander hours clobbered: expected {bystander_hours}, got {got}"
