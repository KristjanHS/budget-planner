"""Zero-FTE WP-row suppression — Q2.1=(a) of capacity-row-cleanup plan.

Assignments with ``total_fte == 0`` stay in the Config PERSON-WP ASSIGNMENTS
auditor table (so the row is still a citable entry) but their per-person
grid row is suppressed on Capacity AND Actuals-Hours so the row-order
mirror invariant holds across both sheets.

Empty-window intersection (non-zero ``total_fte`` but the assignment window
has no overlap with the WP active range) is **not** addressed by this
filter — those rows still render as all-grey rows. See plan
``docs/plans/archive/2026-04-28-capacity-row-cleanup-and-skip-cumulate-plan.md`` § Risks.
"""

import openpyxl
import pytest
from openpyxl.cell import Cell

from budget_planner.person_section import filter_zero_fte_wps
from budget_planner.workbook import init_workbook, rebuild_workbook
from budget_planner.reader import (
    read_capacity_sheet,
    read_config_sheet,
    read_actuals_hours_sheet,
)


def _set(ws, row: int, col: int, value):
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


# ── Unit test ──


def test_filter_zero_fte_wps_drops_zero_keeps_nonzero():
    """Zero-FTE WP IDs are dropped; non-zero bystanders survive in original order.

    Bystander assertion (per testing-anti-patterns rule): the second
    assignment is non-zero and MUST be preserved — proves the filter is
    selective, not "remove everything".
    """
    person_assignments = [
        {"person": "Alice", "wp": "WP1", "total_fte": 0.0, "start": 1, "end": 12},
        {"person": "Alice", "wp": "WP2", "total_fte": 3.0, "start": 1, "end": 12},
        {"person": "Alice", "wp": "WP3", "total_fte": 0, "start": 1, "end": 12},
    ]
    ordered = ["WP1", "WP2", "WP3"]
    assert filter_zero_fte_wps(person_assignments, ordered) == ["WP2"]


# ── Integration: build a workbook with one zero-FTE assignment ──


@pytest.fixture(scope="module")
def zero_fte_wb_path(tmp_path_factory):
    """Workbook where Alice has WP1 zero-FTE + WP2 non-zero (bystander).

    Bob has WP5 non-zero only — gives a second person whose row count is
    unaffected by the filter, anchoring the mirror-invariant assertion to
    a multi-person fixture rather than a single-person edge.
    """
    out = tmp_path_factory.mktemp("zero_fte_wb") / "zero_fte.xlsx"
    assignments = [
        {"person": "Alice", "wp": "WP1", "total_fte": 0.0, "start": 1, "end": 12},
        {"person": "Alice", "wp": "WP2", "total_fte": 3.0, "start": 1, "end": 18},
        {"person": "Bob", "wp": "WP5", "total_fte": 6.0, "start": 4, "end": 13},
    ]
    init_workbook(str(out), assignments=assignments)
    return str(out)


def test_capacity_drops_zero_fte_wp_row(zero_fte_wb_path):
    """Capacity per-person grid omits Alice's WP1 row but keeps WP2."""
    wb = openpyxl.load_workbook(zero_fte_wb_path)
    _, wps = read_config_sheet(wb["Config"])
    cap = read_capacity_sheet(wb["Capacity"], wps)
    alice_wps = cap["person_groups"]["Alice"]["wp_values"]
    assert "WP1" not in alice_wps, "Zero-FTE WP1 row should not appear on Capacity"
    assert "WP2" in alice_wps, "Non-zero bystander WP2 must still be present"
    wb.close()


def test_actuals_hours_drops_zero_fte_wp_row(zero_fte_wb_path):
    """Actuals-Hours per-person grid omits Alice's WP1 row but keeps WP2."""
    wb = openpyxl.load_workbook(zero_fte_wb_path)
    hrs = read_actuals_hours_sheet(wb["Actuals-Hours"])
    alice_rows = hrs["Alice"]
    assert "WP1" not in alice_rows
    assert "WP2" in alice_rows
    wb.close()


def test_capacity_and_actuals_hours_wp_row_order_mirror(zero_fte_wb_path):
    """Capacity and Actuals-Hours expose the same WP-row set per person.

    Load-bearing per ``docs/spec.md`` § Row-order contracts — cross-sheet
    refs target WP positions on both sheets and rely on the order matching.
    Asserts identical WP-id sequences (intersection-with-Config-WPs to
    drop external rows that read_actuals_hours_sheet folds in).

    Order-preservation assumption: both readers return ``dict[wp_id, …]``
    where insertion order matches the top-to-bottom row scan, which in
    turn matches the builder's ``ordered_wp_ids`` write order. If a
    future reader change buckets WPs differently (e.g. lookup-table
    style), this comparison can pass with out-of-order keys while the
    actual row positions are wrong — a stronger anchor would be
    asserting against the row numbers in col-A of each sheet.
    """
    wb = openpyxl.load_workbook(zero_fte_wb_path)
    _, wps = read_config_sheet(wb["Config"])
    wp_ids = {wp["id"] for wp in wps}
    cap = read_capacity_sheet(wb["Capacity"], wps)
    hrs = read_actuals_hours_sheet(wb["Actuals-Hours"])
    for name in cap["person_groups"]:
        cap_wp_seq = [w for w in cap["person_groups"][name]["wp_values"]]
        hrs_wp_seq = [w for w in hrs.get(name, {}) if w in wp_ids]
        assert cap_wp_seq == hrs_wp_seq, (
            f"WP-row order differs for {name}: Capacity={cap_wp_seq} vs Actuals-Hours={hrs_wp_seq}"
        )
    wb.close()


# ── Rebuild contract: A1 silent data loss (Q2.2) ──


def test_rebuild_drops_hours_when_assignment_total_fte_zeroed(tmp_path):
    """A1 (accepted): hours logged on a row whose Config assignment is later
    set to ``total_fte == 0`` are silently discarded on rebuild.

    Per ``docs/plans/archive/2026-04-28-capacity-row-cleanup-and-skip-cumulate-plan.md`` Q2.2,
    there is no preserved-hours carve-out for the zero-FTE filter — the
    row drops everywhere on rebuild even if hours were previously entered.
    No telemetry, no warning. ``git log output/budget_plan*.xlsx`` is
    the audit trail.

    THIS IS NOT A REGRESSION. If this test fails because hours survive
    on a zeroed-out WP row, check whether someone added a preservation
    carve-out without updating the design — the row-order mirror with
    Capacity will break if Actuals-Hours keeps rows Capacity drops.
    """
    path = str(tmp_path / "rebuild_zero.xlsx")
    init_workbook(
        output_path=path,
        assignments=[
            {"person": "Alice", "wp": "WP2", "total_fte": 3.0, "start": 1, "end": 12},
        ],
    )

    # Log some hours on Alice's WP2 row in Actuals-Hours
    wb = openpyxl.load_workbook(path)
    hrs_ws = wb["Actuals-Hours"]
    initial = read_actuals_hours_sheet(hrs_ws)
    assert "WP2" in initial["Alice"], "Setup: Alice should have WP2 row"

    # Find Alice's WP2 row by re-scanning (literal name in col A on header)
    alice_header_row = None
    for r in range(1, hrs_ws.max_row + 1):
        if hrs_ws.cell(r, 1).value == "Alice":
            alice_header_row = r
            break
    assert alice_header_row is not None, "Could not locate Alice header"

    # WP2 is the only WP row, sits one row below header.
    wp2_row = alice_header_row + 1
    _set(hrs_ws, wp2_row, 3, 40.0)  # month 1
    _set(hrs_ws, wp2_row, 4, 60.0)  # month 2
    wb.save(path)
    wb.close()

    # Confirm hours landed in the read.
    wb = openpyxl.load_workbook(path)
    after_seed = read_actuals_hours_sheet(wb["Actuals-Hours"])
    assert after_seed["Alice"]["WP2"][:2] == [40.0, 60.0]

    # Now zero out the assignment via Config and rebuild.
    config_ws = wb["Config"]
    for r in range(1, config_ws.max_row + 1):
        if config_ws.cell(r, 1).value == "Alice" and config_ws.cell(r, 2).value == "WP2":
            _set(config_ws, r, 3, 0.0)
            break
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=path)

    wb = openpyxl.load_workbook(path)
    rebuilt = read_actuals_hours_sheet(wb["Actuals-Hours"])
    assert "WP2" not in rebuilt.get("Alice", {}), "Zero-FTE row should be dropped from Actuals-Hours after rebuild"

    _, wps = read_config_sheet(wb["Config"])
    cap_after = read_capacity_sheet(wb["Capacity"], wps)
    assert "WP2" not in cap_after["person_groups"]["Alice"]["wp_values"], (
        "Zero-FTE row should also be dropped from Capacity"
    )
    wb.close()
