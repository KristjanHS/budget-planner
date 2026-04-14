"""Row-grouping tests for Capacity + Actuals-Hours (W1 row grouping).

Each person block collapses as a unit: header row stays at ``outline_level
== 0``; every body row (WP rows, external rows, TOTAL, optional REMAINING,
optional EU TOTAL) has ``outline_level == 1``. The sheet's
``outlinePr.summaryBelow`` is set to ``False`` so the collapse button sits
next to the person header (above the block), not below it.

Round-trip test asserts ``rebuild_workbook`` preserves the outline levels
— openpyxl stores row_dimensions natively, but the full init→edit→rebuild
→reopen loop is the contract we care about.
"""

from openpyxl import load_workbook

from budget_planner.workbook import init_workbook, rebuild_workbook
from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS


def _find_person_header_rows(ws) -> list[int]:
    """Return rows whose col A is the start of a person header (merged A:B).

    Person headers are the only rows in the person-grid area that are
    merged across columns 1–2. We scan merged_cells for A:B horizontal
    merges below the year/month header rows.
    """
    headers: list[int] = []
    for mr in ws.merged_cells.ranges:
        if (mr.min_col == 1 and mr.max_col == 2
                and mr.min_row == mr.max_row):
            headers.append(mr.min_row)
    return sorted(headers)


def _person_body_rows(ws, header_row: int, next_header_row: int | None) -> list[int]:
    """Return the contiguous body rows for one person block.

    Body = ``header_row + 1 .. next_header_row - 1`` (or last populated row).
    """
    end = (next_header_row - 1) if next_header_row is not None else ws.max_row
    return list(range(header_row + 1, end + 1))


def _assert_person_blocks_grouped(ws, sheet_name: str) -> None:
    headers = _find_person_header_rows(ws)
    # Drop any merged A:B above the main grid — those are band labels/
    # section headers written before the person grid begins. Person
    # headers on Capacity/Actuals-Hours start a few rows in; identify
    # them as the headers that have a non-empty body between consecutive
    # merges.
    assert headers, f"{sheet_name}: expected at least one merged A:B row"

    # Keep only the subset that looks like a person header: next-row col B
    # is either blank/formula (WP row) or TOTAL/EU TOTAL label is within
    # body before next header. Simplest filter: any header whose +1 row
    # carries data (non-empty col A) or is a TOTAL row within the span.
    # In practice the demo workbook has clear person headers; filter by
    # any level-1 row occurring within (header, next_header).
    for i, hr in enumerate(headers):
        next_hr = headers[i + 1] if i + 1 < len(headers) else None
        body = _person_body_rows(ws, hr, next_hr)
        grouped_in_body = [
            r for r in body
            if (ws.row_dimensions[r].outline_level or 0) == 1
        ]
        if not grouped_in_body:
            # Not a person header — e.g. a pre-grid section label merge.
            continue

        # Header row itself must be level 0.
        assert (ws.row_dimensions[hr].outline_level or 0) == 0, (
            f"{sheet_name}: person-header row {hr} must stay at "
            f"outline_level 0 (actual: "
            f"{ws.row_dimensions[hr].outline_level})"
        )
        # Every body row must be level 1 — including WP rows, externals,
        # TOTAL, REMAINING (capacity only), and EU TOTAL.
        for br in body:
            level = ws.row_dimensions[br].outline_level or 0
            assert level == 1, (
                f"{sheet_name}: body row {br} under person header {hr} "
                f"must be outline_level 1 (actual: {level})"
            )


def _assert_summary_below_false(ws, sheet_name: str) -> None:
    op = ws.sheet_properties.outlinePr
    assert op is not None, f"{sheet_name}: outlinePr should be set"
    assert op.summaryBelow is False, (
        f"{sheet_name}: summaryBelow must be False so the row-outline "
        f"button sits next to each person header"
    )


def test_capacity_row_grouping(assigned_wb_path):
    wb = load_workbook(assigned_wb_path)
    ws = wb["Capacity"]
    _assert_summary_below_false(ws, "Capacity")
    _assert_person_blocks_grouped(ws, "Capacity")
    wb.close()


def test_actuals_hours_row_grouping(assigned_wb_path):
    wb = load_workbook(assigned_wb_path)
    ws = wb["Actuals-Hours"]
    _assert_summary_below_false(ws, "Actuals-Hours")
    _assert_person_blocks_grouped(ws, "Actuals-Hours")
    wb.close()


def test_row_grouping_survives_rebuild(tmp_path):
    """Init → rebuild → reopen: outline levels + summaryBelow preserved.

    openpyxl stores ``row_dimensions`` natively, but the rebuild path
    re-runs every builder — this test pins the full round-trip.
    """
    path = str(tmp_path / "rebuild.xlsx")
    init_workbook(path, assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Capture initial grouped-row sets per sheet.
    wb = load_workbook(path)
    initial: dict[str, frozenset[int]] = {}
    for sheet in ("Capacity", "Actuals-Hours"):
        ws = wb[sheet]
        initial[sheet] = frozenset(
            r for r, dim in ws.row_dimensions.items()
            if (dim.outline_level or 0) == 1
        )
    wb.close()

    # Sanity: initial grouping actually happened.
    for sheet, rows in initial.items():
        assert rows, f"{sheet}: initial build should have grouped rows"

    rebuild_workbook(path)

    wb = load_workbook(path)
    for sheet in ("Capacity", "Actuals-Hours"):
        ws = wb[sheet]
        _assert_summary_below_false(ws, sheet)
        rebuilt = frozenset(
            r for r, dim in ws.row_dimensions.items()
            if (dim.outline_level or 0) == 1
        )
        assert rebuilt == initial[sheet], (
            f"{sheet}: grouped-row set drifted across rebuild\n"
            f"  initial: {sorted(initial[sheet])}\n"
            f"  rebuilt: {sorted(rebuilt)}"
        )
    wb.close()
