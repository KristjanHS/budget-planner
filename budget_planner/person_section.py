"""Shared per-person row geometry helpers.

Used by the Capacity sheet and the Actuals-Hours sheet so both share the same
per-person block layout (header → WP rows → external rows → TOTAL → optional
REMAINING → optional EU TOTAL). Capacity needs the REMAINING row (budget
view); Actuals-Hours does not (hours log, not a budget).
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class PersonGridContext:
    """Loop-invariant inputs shared across every per-person grid write.

    Used by ``capacity._write_person_grid`` and
    ``actuals_hours._write_person_hours_grid`` — each builder receives one
    context instance plus its sheet-specific extras (per-person data,
    val-column offsets unique to that sheet, etc.). Fields mirror the
    orchestrator's pre-loop computations; bundling them avoids threading
    8 kwargs through every person-loop call site.
    """
    work_packages: list[dict]
    external_commitments: list[dict]
    wp_active_map: dict[str, frozenset[int]]
    project_params: dict
    month_col_start: int
    month_col_end: int
    val_col_sum: int
    config_first_wp_row: int


@dataclass
class PersonLayout:
    header_row: int
    wp_rows: dict[str, int]
    external_rows: dict[str, int]
    total_row: int
    remaining_row: int | None
    eu_total_row: int | None
    next_row: int


def plan_person_layout(
    *,
    start_row: int,
    wp_ids: list[str],
    external_names: list[str],
    include_remaining: bool,
    include_eu_total: bool,
) -> PersonLayout:
    """Compute sequential row positions for a single person's section.

    Layout order: header, one row per WP id, one row per external name,
    TOTAL, optional REMAINING, optional EU TOTAL.
    """
    r = start_row
    header_row = r
    r += 1
    wp_rows = {wp: r + i for i, wp in enumerate(wp_ids)}
    r += len(wp_ids)
    external_rows = {name: r + i for i, name in enumerate(external_names)}
    r += len(external_names)
    total_row = r
    r += 1
    remaining_row = r if include_remaining else None
    if include_remaining:
        r += 1
    eu_total_row = r if include_eu_total else None
    if include_eu_total:
        r += 1
    return PersonLayout(
        header_row=header_row,
        wp_rows=wp_rows,
        external_rows=external_rows,
        total_row=total_row,
        remaining_row=remaining_row,
        eu_total_row=eu_total_row,
        next_row=r,
    )


def apply_person_row_outline(ws, layout: PersonLayout) -> None:
    """Mark every body row of a person block as ``outline_level=1``.

    Body rows = everything between the header row (exclusive) and
    ``next_row`` (exclusive) — i.e. WP rows, external rows, TOTAL,
    optional REMAINING, optional EU TOTAL. Header row stays at the
    default level 0 so the outline button (with ``summaryBelow=False``
    on the sheet) sits next to the header and collapses the block
    beneath it.

    Idempotent: calling twice leaves the final level at 1.
    """
    for r in range(layout.header_row + 1, layout.next_row):
        ws.row_dimensions[r].outline_level = 1
