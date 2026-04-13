"""Shared per-person row geometry helpers.

Used by the Capacity sheet and the Actuals-Hours sheet so both share the same
per-person block layout (header → WP rows → external rows → TOTAL → optional
REMAINING → optional EU TOTAL). Capacity needs the REMAINING row (budget
view); Actuals-Hours does not (hours log, not a budget).
"""

from dataclasses import dataclass


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
