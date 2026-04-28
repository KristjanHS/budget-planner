"""Layout discovery + shared helpers used by every per-sheet reader.

Owns the cross-sheet discovery primitive (``discover_layout``) plus the
parsing/scanning helpers (``_find_header_row``, ``_data_row_after``,
``_iter_named_range_rows``, ``_parse_config_float``, ``_parse_fte_rate``,
``_safe_float``, ``_resolve_plan_ref``, ``_read_monthly_values``) that the
per-sheet reader modules import via relative ``from .layout_discovery``.
"""

import re
from decimal import Decimal
from typing import Literal, overload

from budget_planner.layout import (
    SECTION_HEADER_LABELS,
    SECTION_HEADER_ROWS,
    SHEET_CONFIG,
    SHEET_PLAN,
)

__all__ = ["discover_layout"]


def _safe_float(value, default: float) -> float:
    """Convert value to float, returning default if it's a formula or None."""
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def _parse_config_float(value) -> float | None:
    """Parse Config numeric cells: real numbers and text (e.g. European ``0,35``).

    Excel often stores locale-formatted values as *text* (green triangle). The
    roster reader used to stop at the first such cell, so later names were
    omitted from ``team`` and assignments looked "orphaned".

    Returns:
        ``0.0`` when *value* is ``None`` (empty cell on a row the caller has
        already decided is valid). A float when *value* is numeric or a
        parseable numeric string (US/EU locale, NBSP-tolerant). ``None`` when
        *value* is a bool, a formula string (``=...``), or any other non-numeric
        string — signalling the caller to stop scanning the section.
    """
    # bool is a subclass of int; reject it explicitly before the int/float branch
    # so True/False don't silently coerce to 1.0/0.0 in a numeric context.
    if isinstance(value, bool):
        return None
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, Decimal):
        return float(value)
    if not isinstance(value, str):
        return None
    s = value.strip()
    if not s:
        # Empty/whitespace string: treat as "non-numeric sentinel" (matches pre-
        # _parse_config_float break-on-type-guard behaviour). Returning 0.0 here
        # would append a zero-capacity ghost row for a named person with a blank
        # FTE cell.
        return None
    if s.startswith("="):
        return None
    s = s.replace(" ", "").replace(" ", "")
    norm = s
    if "," in norm and "." not in norm:
        norm = norm.replace(",", ".")
    elif "," in norm and "." in norm:
        if norm.rfind(",") > norm.rfind("."):
            norm = norm.replace(".", "").replace(",", ".")
        else:
            norm = norm.replace(",", "")
    try:
        return float(norm)
    except ValueError:
        return None


def _parse_fte_rate(value) -> float:
    """Parse the FTE Rate cell value as a numeric fallback.

    The builder always regenerates B4 as a SUMPRODUCT over the live grade/roster
    layout, so formula strings in the cell are intentionally discarded — the
    rate is a readout of the grade table, not an independent input. Numeric
    values (and None) are coerced to float, defaulting to 7500.
    """
    if isinstance(value, (int, float)):
        return float(value or 7500)
    return 7500.0


@overload
def _find_header_row(
    ws,
    pattern: str,
    col: int = ...,
    max_row: int | None = ...,
    required: Literal[True] = ...,
) -> int: ...


@overload
def _find_header_row(
    ws,
    pattern: str,
    col: int = ...,
    max_row: int | None = ...,
    required: Literal[False] = ...,
) -> int | None: ...


def _find_header_row(ws, pattern, col=1, max_row=None, required=True):
    """Find row where cell exactly matches pattern (case-insensitive, stripped).

    Uses exact match (not substring) to avoid layout-drift silent failures —
    e.g. a renamed or nested section header won't accidentally satisfy the search.

    Args:
        ws: openpyxl worksheet.
        pattern: String to search for (exact match after strip + upper-case).
        col: Column number to scan (default 1 = col A).
        max_row: Upper row limit. If None, scans the whole sheet.
        required: If True and pattern not found, raises ValueError.

    Returns:
        Row number (int) or None if not found and required=False.

    Raises:
        ValueError: If required=True and pattern is not found.
    """
    limit = max_row if max_row is not None else ws.max_row
    target = pattern.strip().upper()
    for r in range(1, limit + 1):
        val = ws.cell(r, col).value
        if val is not None and str(val).strip().upper() == target:
            return r
    if required:
        raise ValueError(
            f"Section header {pattern!r} not found in column {col} "
            f"(rows 1-{limit}). Was the workbook structure modified?"
        )
    return None


def _data_row_after(header_row: int, header_rows: int = SECTION_HEADER_ROWS) -> int:
    """First data row after a section header (title row + column header row)."""
    return header_row + header_rows


_NAMED_RANGE_REF_RE = re.compile(r"\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?$")


def _iter_named_range_rows(wb, name: str) -> tuple[int, int] | None:
    """Resolve a defined name to its (first_row, last_row) on the first target range.

    Returns None if the name is not defined or cannot be parsed. Callers fall back
    to sentinel-scan reading when this returns None. Excel auto-extends defined
    names on row insertion, so this is the authoritative live span after
    Excel-side edits.
    """
    dn = None
    try:
        dn = wb.defined_names[name]
    except (KeyError, TypeError):
        return None
    if dn is None:
        return None
    try:
        destinations = list(dn.destinations)
    except Exception:
        return None
    if not destinations:
        return None
    _sheet, ref = destinations[0]
    m = _NAMED_RANGE_REF_RE.match(ref)
    if not m:
        return None
    first = int(m.group(1))
    last = int(m.group(2)) if m.group(2) else first
    if last < first:
        return None
    return (first, last)


def _read_monthly_values(ws, row, col_start, num_months):
    """Read monthly values from a row, returning None for formulas/empty."""
    values = []
    for m in range(num_months):
        col = col_start + m
        val = ws.cell(row, col).value
        if val is not None and isinstance(val, str) and val.startswith("="):
            values.append(None)
        elif val is None or val == "":
            values.append(None)
        else:
            try:
                values.append(float(val))
            except (ValueError, TypeError):
                values.append(None)
    return values


def _resolve_plan_ref(cell_value, plan_ws, config_ws=None):
    """Resolve a cell value that might be a =Plan!A{n} or =Config!A{n} formula."""
    if cell_value is None:
        return None
    val = str(cell_value)
    # Try =Plan!A{n} — resolve by reading the Plan cell (which itself may be =Config!A{n})
    m = re.match(rf"={SHEET_PLAN}!A(\d+)", val)
    if m and plan_ws:
        plan_row = int(m.group(1))
        resolved = plan_ws.cell(plan_row, 1).value
        if resolved and isinstance(resolved, str) and resolved.startswith("="):
            # Follow the chain: Plan cell is a formula like =Config!A8
            m2 = re.match(rf"={SHEET_CONFIG}!A(\d+)", resolved)
            if m2 and config_ws:
                config_row = int(m2.group(1))
                final = config_ws.cell(config_row, 1).value
                return str(final) if final else None
        return str(resolved) if resolved else None
    # Try =Config!A{n}
    m3 = re.match(rf"={SHEET_CONFIG}!A(\d+)", val)
    if m3 and config_ws:
        config_row = int(m3.group(1))
        resolved = config_ws.cell(config_row, 1).value
        return str(resolved) if resolved else None
    if not val.startswith("="):
        return val
    return None


def discover_layout(ws) -> dict:
    """Scan existing Plan worksheet, reconstruct layout dict.

    After the Plan config table was removed, the Plan sheet has:
    - Row 1: compact header
    - Row 2: blank
    - Row 3: PLANNED FTE section header
    - Rows 4-5: year/month headers
    - Rows 6+: WP data rows

    The config_first_wp_row/config_last_wp_row refer to the Config sheet's
    WP table rows (row 8+ by default).
    """
    # Find PLANNED block header
    planned_header_row = _find_header_row(
        ws, SECTION_HEADER_LABELS["planned_fte"], required=True
    )

    # Find the first WP data row after the PLANNED header.
    # WP rows have a formula (e.g. =Config!A8) or literal WP ID in column A.
    planned_first_wp_row = None
    for r in range(planned_header_row + 1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is None or str(val).strip() == "":
            continue
        val_str = str(val)
        # A formula referencing a Config row (e.g. =Config!A8) means WP data
        if val_str.startswith("="):
            planned_first_wp_row = r
            break
        # A literal WP ID (skip known header labels)
        if not any(
            kw in val_str.upper()
            for kw in ("PLANNED", "YEAR", "SUM", "BUDGET", "WP ID", "WP NAME")
        ):
            planned_first_wp_row = r
            break
    if planned_first_wp_row is None:
        raise ValueError(
            "Cannot find first WP data row in PLANNED block. "
            "Expected formula (=Config!A{n}) or literal WP ID after "
            f"PLANNED header at row {planned_header_row}."
        )

    # Count WP rows: scan until empty col A or TOTAL row
    planned_last_wp_row = planned_first_wp_row
    while True:
        next_val = ws.cell(planned_last_wp_row + 1, 1).value
        next_name = str(ws.cell(planned_last_wp_row + 1, 2).value or "")
        if not next_val or next_name.upper().startswith("TOTAL"):
            break
        planned_last_wp_row += 1
    num_wps = planned_last_wp_row - planned_first_wp_row + 1
    planned_total_row = planned_last_wp_row + 1

    # Derive Config sheet WP rows from the first WP formula reference
    # Formula pattern: =Config!A{row}
    first_val = ws.cell(planned_first_wp_row, 1).value
    cfg_wp_first = None
    if first_val and isinstance(first_val, str):
        m = re.match(rf"={SHEET_CONFIG}!A(\d+)", str(first_val))
        if m:
            cfg_wp_first = int(m.group(1))
    if cfg_wp_first is None:
        # Fallback: discover from Config sheet directly (legacy workbooks
        # with literal WP IDs instead of =Config!A{n} formulas)
        config_ws = (
            ws.parent[SHEET_CONFIG] if SHEET_CONFIG in ws.parent.sheetnames else None
        )
        wp_section = (
            _find_header_row(
                config_ws, SECTION_HEADER_LABELS["wp_budget"], required=False
            )
            if config_ws is not None
            else None
        )
        if wp_section is None:
            raise ValueError(
                "Cannot locate Config WP Budget section. "
                "Regenerate workbook with `python -m budget_planner init`."
            )
        cfg_wp_first = _data_row_after(wp_section)
    cfg_wp_last = cfg_wp_first + num_wps - 1

    return {
        "config_first_wp_row": cfg_wp_first,
        "config_last_wp_row": cfg_wp_last,
        "planned_header_row": planned_header_row,
        "planned_first_wp_row": planned_first_wp_row,
        "planned_last_wp_row": planned_last_wp_row,
        "planned_total_row": planned_total_row,
    }
