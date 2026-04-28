"""Read the Actuals-Hours sheet — preserved per-person blue-input hours."""

import re

from openpyxl.utils import get_column_letter

from budget_planner.layout import (
    MONTH_COL_START,
    SECTION_HEADER_LABELS,
    SHEET_CONFIG,
)

from .layout_discovery import _find_header_row

__all__ = ["read_actuals_hours_sheet"]


def read_actuals_hours_sheet(
    ws,
    project_months: int | None = None,
    month_col_start: int = MONTH_COL_START,
) -> dict[str, dict[str, list[float | None]]]:
    """Read Actuals-Hours blue-input hours for round-trip preservation.

    Anchors on the "WP ID" month-header row, then walks person groups. Each
    person section starts with a merged A:B header whose col A holds the
    literal person name (written as a literal, not a formula, so the
    forecast-EUR VLOOKUP works under ``data_only``). Under each person, WP
    rows carry ``=Config!A{n}`` in col A; external rows carry a literal
    project name. TOTAL and EU TOTAL rows are skipped.

    Columns with ``outline_level >= 1`` belong to the collapsible Forecast
    EUR band — those are derived formulas, never user input, so they are
    skipped when reading monthly values. Tolerates internal blank rows
    inside a person section (per the rebuild shape-drift design).

    Args:
        ws: Actuals-Hours worksheet.
        project_months: Number of project months. If None, falls back to
            ``config.PROJECT_MONTHS`` for backward compatibility.
        month_col_start: First month column (default 3 = col C).

    Returns:
        ``{person_name: {wp_id_or_external_project: [h_m1, h_m2, ...]}}``.
        Each monthly entry is a float (user input) or None (blank cell /
        greyed inactive cell / formula).
    """
    if project_months is None:
        from budget_planner.config import PROJECT_MONTHS

        project_months = PROJECT_MONTHS

    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None

    # Anchor: the month-header row ("WP ID" in col A). Person sections
    # begin strictly below it.
    month_header_row = _find_header_row(ws, SECTION_HEADER_LABELS["wp_id_column"])

    # Pre-compute the set of hour columns (skip outline-level >= 1 columns
    # — those are the forecast-EUR band mirror, recomputed on rebuild).
    hour_cols: list[int] = []
    for mi in range(project_months):
        col = month_col_start + mi
        if ws.column_dimensions[get_column_letter(col)].outline_level >= 1:
            continue
        hour_cols.append(col)

    # ── Find person header rows ──
    # Convention (actuals_hours.py): person headers are A:B merged with a
    # literal name in col A (not a formula) + SECTION_FILL banner. Locate
    # them by matching the A:B merge anchored at col 1 — no reliance on
    # font colour, which openpyxl doesn't always round-trip cleanly.
    merged_a_b_rows: set[int] = set()
    for mr in ws.merged_cells.ranges:
        if mr.min_col == 1 and mr.max_col >= 2 and mr.min_row == mr.max_row:
            merged_a_b_rows.add(mr.min_row)

    person_headers: list[tuple[int, str]] = []
    for r in sorted(merged_a_b_rows):
        if r <= month_header_row:
            continue
        val = ws.cell(r, 1).value
        if val is None:
            continue
        if isinstance(val, str) and val.startswith("="):
            # Defensive: newer builder writes literal names, but tolerate
            # older formula-headers by resolving through Config.
            m = re.match(rf"={SHEET_CONFIG}!A(\d+)", val)
            if m and config_ws is not None:
                resolved = config_ws.cell(int(m.group(1)), 1).value
                if resolved:
                    person_headers.append((r, str(resolved).strip()))
            continue
        name = str(val).strip()
        if not name:
            continue
        # Skip merged section-header rows like "FORECAST EUR -i (from hours)"
        # that might also span A:B in future layouts. Person headers use
        # SECTION_FILL; the forecast-band label row uses the -i orange fill.
        # Simple heuristic: section banner is on row 1; skip it.
        if r == 1:
            continue
        person_headers.append((r, name))

    def _read_hours_row(row: int) -> list[float | None]:
        """Read hour cells for one data row, skipping outline-level columns."""
        out: list[float | None] = []
        col_idx = 0
        for mi in range(project_months):
            col = month_col_start + mi
            if col_idx < len(hour_cols) and hour_cols[col_idx] == col:
                val = ws.cell(row, col).value
                if val is None or val == "":
                    out.append(None)
                elif isinstance(val, str):
                    if val.startswith("="):
                        out.append(None)
                    else:
                        try:
                            out.append(float(val))
                        except ValueError:
                            out.append(None)
                else:
                    try:
                        out.append(float(val))
                    except (TypeError, ValueError):
                        out.append(None)
                col_idx += 1
            else:
                # Column was suppressed by outline filter — reader would
                # have no hour-input to read here by design. Emit None to
                # keep the per-month list length stable.
                out.append(None)
        return out

    def _any_nonnull(series: list[float | None]) -> bool:
        """True if the row has any logged value (including 0.0). Used to
        prefer a later duplicate-key row that actually has data over an
        earlier all-blank one — a user logging 0 is still "present"."""
        return any(v is not None for v in series)

    hours: dict[str, dict[str, list[float | None]]] = {}

    for hi, (header_row, name) in enumerate(person_headers):
        if hi + 1 < len(person_headers):
            group_end = person_headers[hi + 1][0] - 1
        else:
            group_end = ws.max_row

        per_wp: dict[str, list[float | None]] = {}

        for dr in range(header_row + 1, group_end + 1):
            col_b = ws.cell(dr, 2).value
            # Stop-or-skip on TOTAL / EU TOTAL rows.
            if col_b is not None and str(col_b).strip().upper() in (
                "TOTAL",
                "EU TOTAL",
            ):
                continue

            col_a = ws.cell(dr, 1).value
            if col_a is None or (isinstance(col_a, str) and not col_a.strip()):
                # Internal blank row — tolerate and keep scanning.
                continue

            key: str | None = None
            if isinstance(col_a, str) and col_a.startswith("="):
                # WP row: =Config!A{n}
                m = re.match(rf"={SHEET_CONFIG}!A(\d+)", col_a)
                if m and config_ws is not None:
                    resolved = config_ws.cell(int(m.group(1)), 1).value
                    if resolved:
                        key = str(resolved).strip()
            else:
                # External project row: literal project name.
                key = str(col_a).strip()

            if not key:
                continue

            series = _read_hours_row(dr)
            # Preserve every identifiable row (including all-None) so the
            # builder sees an explicit "present but untouched" signal —
            # matches the ``read_actuals_values`` preservation contract.
            # First-wins if duplicate keys appear in one section.
            if key not in per_wp or _any_nonnull(series):
                per_wp[key] = series

        if per_wp:
            hours[name] = per_wp

    return hours
