"""Demo data seeder — Actuals EUR + Actuals-Hours + EUR through-month override.

Capacity FTE values are auto-distributed by the builder via assignments;
this module seeds the Actuals EUR spend and per-person Actuals-Hours so
the reconciliation visuals (greyed forecast-EUR band, Capacity to-date
block) have something to render. Writes a blue-input override to
``Actuals!$B$1`` at ``DEMO_CURRENT_MONTH − 3`` to model the accountant's
3-month finance-close lag (EUR cadence trails hours). Hours through-month
is left auto-derived — the row-1 scan on Actuals-Hours naturally returns
``DEMO_CURRENT_MONTH`` because hours are seeded through that month.
"""

from openpyxl import load_workbook

from budget_planner.config import (
    DEMO_ACTUALS_MULTIPLIERS,
    DEMO_CAPACITY_ASSIGNMENTS,
    DEMO_CURRENT_MONTH,
    TEAM,
    WORK_PACKAGES,
)
from budget_planner.layout import (
    MONTH_COL_START,
    SECTION_HEADER_LABELS,
    SHEET_ACTUALS,
    SHEET_ACTUALS_HOURS,
    SHEET_CONFIG,
    wp_active_months,
    wp_month_bounds,
)
from budget_planner.readers.layout_discovery import _data_row_after, _find_header_row


def seed_demo_data(workbook_path: str) -> str:
    """Seed Actuals EUR and Actuals-Hours, plus the EUR through-month override.

    Args:
        workbook_path: Path to an existing workbook (from init_workbook).

    Returns:
        The workbook path.
    """
    wb = load_workbook(workbook_path)

    # ── Seed Actuals EUR values ──
    act_ws = wb[SHEET_ACTUALS]

    # Find first WP row in Actuals (row after "WP ID" header)
    first_wp_row = None
    for r in range(1, act_ws.max_row + 1):
        if act_ws.cell(r, 1).value == "WP ID":
            first_wp_row = r + 1
            break

    if first_wp_row is not None:
        # Read per-WP rates from Config col D
        config_ws = wb[SHEET_CONFIG]
        wp_budget_row = _data_row_after(
            _find_header_row(config_ws, SECTION_HEADER_LABELS["wp_budget"])
        )
        for i, wp in enumerate(WORK_PACKAGES):
            row = first_wp_row + i
            wp_rate = config_ws.cell(wp_budget_row + i, 4).value or 6500
            multiplier = DEMO_ACTUALS_MULTIPLIERS[i % len(DEMO_ACTUALS_MULTIPLIERS)]
            wp_start, wp_end = wp_month_bounds(wp["id"], WORK_PACKAGES)
            active_months = wp_end - wp_start + 1
            monthly_planned = (wp["fte"] / active_months) * wp_rate
            for m_idx in range(wp_start - 1, min(wp_end, DEMO_CURRENT_MONTH)):
                col = MONTH_COL_START + m_idx
                eur_value = round(monthly_planned * multiplier)
                act_ws.cell(row, col).value = eur_value

    # ── Seed EUR through-month override on Actuals!$B$1 ──
    # Models a 3-month finance-close lag (EUR cadence trails hours). Hours
    # through-month resolves via the auto-derived scan on Actuals-Hours;
    # no override needed because hours are seeded through DEMO_CURRENT_MONTH.
    hours_through = DEMO_CURRENT_MONTH
    act_ws.cell(row=1, column=2, value=max(1, DEMO_CURRENT_MONTH - 3))

    # ── Seed Actuals-Hours per-person monthly hours ──
    # Reconstruct per-person WP ordering from DEMO_CAPACITY_ASSIGNMENTS
    # (dedup by person+wp preserving first occurrence) — mirrors the
    # order actuals_hours.py builds rows in.
    HOURS_PER_FTE_MONTH = 143
    per_person: dict[str, list[dict]] = {}
    seen: set[tuple[str, str]] = set()
    for asgn in DEMO_CAPACITY_ASSIGNMENTS:
        key = (asgn["person"], asgn["wp"])
        if key in seen:
            continue
        seen.add(key)
        per_person.setdefault(asgn["person"], []).append(asgn)

    team_names = {m["name"] for m in TEAM}
    wp_index = {w["id"]: i for i, w in enumerate(WORK_PACKAGES)}
    wp_active_set = {
        w["id"]: wp_active_months(w["id"], WORK_PACKAGES)
        for w in WORK_PACKAGES
    }

    if SHEET_ACTUALS_HOURS in wb.sheetnames:
        hours_ws = wb[SHEET_ACTUALS_HOURS]
        name_to_header_row: dict[str, int] = {}
        for r in range(1, hours_ws.max_row + 1):
            v = hours_ws.cell(r, 1).value
            if isinstance(v, str) and v in team_names:
                name_to_header_row[v] = r

        for member in TEAM:
            name = member["name"]
            header_row = name_to_header_row.get(name)
            if header_row is None:
                continue
            for i, asgn in enumerate(per_person.get(name, [])):
                wp_id = asgn["wp"]
                if wp_id not in wp_index:
                    continue
                row = header_row + 1 + i
                wp_start, wp_end = wp_month_bounds(wp_id, WORK_PACKAGES)
                a_start = max(wp_start, asgn["start"])
                a_end = min(wp_end, asgn["end"])
                active_months = a_end - a_start + 1
                if active_months <= 0:
                    continue
                monthly_fte = asgn["total_fte"] / active_months
                multiplier = DEMO_ACTUALS_MULTIPLIERS[
                    wp_index[wp_id] % len(DEMO_ACTUALS_MULTIPLIERS)
                ]
                active_set = wp_active_set[wp_id]
                for m_idx in range(a_start - 1, min(a_end, hours_through)):
                    month_num = m_idx + 1
                    if month_num not in active_set:
                        continue
                    col = MONTH_COL_START + m_idx
                    hours = round(monthly_fte * HOURS_PER_FTE_MONTH * multiplier)
                    hours_ws.cell(row, col).value = hours

    wb.save(workbook_path)
    wb.close()
    print(f"Demo data seeded: {workbook_path}")
    return workbook_path
