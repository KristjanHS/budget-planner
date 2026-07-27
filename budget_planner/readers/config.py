"""Read the Config sheet — project params, WP budget, roster, periods,
assignments, and external commitments."""

import os
import warnings
from collections.abc import Iterator

from openpyxl.utils import get_column_letter

from budget_planner.layout import (
    COL_ASSIGN_END,
    COL_ASSIGN_FTE,
    COL_ASSIGN_PERSON,
    COL_ASSIGN_START,
    COL_ASSIGN_WP,
    COL_EXT_END,
    COL_EXT_EU,
    COL_EXT_FTE,
    COL_EXT_PERSON,
    COL_EXT_PROJECT,
    COL_EXT_START,
    COL_GRADE_RATE,
    COL_ROSTER_FTE,
    COL_ROSTER_GRADE,
    COL_ROSTER_PRIORITY,
    COL_RP_END,
    COL_RP_START,
    COL_WP_BUDGET_MINUS_I,
    COL_WP_BUDGET_PLUS_I,
    COL_WP_END,
    COL_WP_FTE,
    COL_WP_RATE,
    COL_WP_START,
    COL_WP_TYPE,
    SECTION_HEADER_LABELS,
    SHEET_CONFIG,
)

from .layout_discovery import (
    _data_row_after,
    _find_header_row,
    _iter_named_range_rows,
    _parse_config_float,
    _parse_fte_rate,
    _safe_float,
)
from .timelog import _read_vacation_month_num

__all__ = [
    "read_config_sheet",
    "read_plan_config",
    "read_grade_rates",
    "read_team_from_config",
    "read_reporting_periods",
    "read_assignments_from_config",
    "read_externals_from_config",
]


def _iter_config_section_rows(
    ws,
    *,
    header_row: int,
    named_range: str | None,
    key_col: int,
    numeric_col: int,
) -> Iterator[tuple[int, float]]:
    """Yield ``(row, numeric_parsed)`` for each data row of a Config section.

    Hybrid named-range-or-sentinel scan shared by roster/assignments/externals:
    range-bounded when the named range exists (tolerates mid-range blanks from
    Excel inserts), sentinel-terminated past ``range_end`` or when no name is
    defined. ``_parse_config_float`` break-on-invalid semantics: ``None`` from
    bool/formula/blank/non-numeric stops the scan outside the range; inside,
    it is skipped as a mid-insert blank. Callers extract row payloads.
    """
    wb = ws.parent
    span = _iter_named_range_rows(wb, named_range) if wb is not None and named_range else None
    expected_first = _data_row_after(header_row)
    if span is not None and span[0] < expected_first:
        span = None
    if span is not None:
        first_data, range_end = span
    else:
        first_data = expected_first
        range_end = None

    # `yielded_any` stands in for "collected any rows" — load-bearing for the
    # past-range-with-no-yields break below. Equivalent only while every caller
    # appends each yielded row; don't add conditional-skip logic in callers
    # without revisiting that break.
    yielded_any = False
    r = first_data
    while r <= ws.max_row:
        key_val = ws.cell(r, key_col).value
        numeric_parsed = _parse_config_float(ws.cell(r, numeric_col).value)
        key_blank = key_val is None or str(key_val).strip() == ""
        inside_range = range_end is not None and r <= range_end
        # Past range_end with nothing collected — section is empty; break
        # rather than sentinel-scan into the next section's header.
        if range_end is not None and r > range_end and not yielded_any:
            break
        if key_blank or numeric_parsed is None:
            if inside_range:
                r += 1
                continue
            break
        yielded_any = True
        yield r, numeric_parsed
        r += 1


def _read_wp_budget_section(ws, project_months: int) -> list[dict]:
    """Parse the WP BUDGET section -> list of WP dicts (id/name/fte/rate/
    rate_raw/budget_plus_i/start/end/type). Owns layout detection (legacy
    Ceiling vs post-restructure Budget -i/+i), per-row validation, and the
    legacy Ceiling→Budget +i ×1.25 migration + >€3M sanity guard.
    Pass ``project_months=0`` to skip the end<=months check.
    """
    # WP BUDGET row carries id/name/fte/rate + start/end/type (cols I/J/K post-
    # indirect-view; H/I/J pre-restructure). Layout detected via col-5 header.
    wp_section = _find_header_row(ws, SECTION_HEADER_LABELS["wp_budget"])
    work_packages: list[dict] = []
    seen_ids: set[str] = set()
    valid_types = {"flat", "front-loaded", "bell", "back-loaded"}
    header_row = wp_section + 1
    col5_header = str(ws.cell(header_row, COL_WP_BUDGET_MINUS_I).value or "").strip()
    is_legacy_ceiling_layout = "Ceiling" in col5_header and "Budget -i" not in col5_header
    start_col = COL_WP_START - 1 if is_legacy_ceiling_layout else COL_WP_START
    end_col = COL_WP_END - 1 if is_legacy_ceiling_layout else COL_WP_END
    type_col = COL_WP_TYPE - 1 if is_legacy_ceiling_layout else COL_WP_TYPE
    r = _data_row_after(wp_section)
    legacy_rows: list[tuple[str, float]] = []  # for post-loop >€3M sanity check
    while r <= ws.max_row:
        wp_id = ws.cell(r, 1).value
        col_b = ws.cell(r, 2).value
        if wp_id is None or str(wp_id).strip() == "":
            break
        # Stop at TOTAL row
        if col_b and str(col_b).upper() == "TOTAL":
            break
        rate_cell = ws.cell(r, COL_WP_RATE).value
        id_str = str(wp_id).strip()
        if id_str in seen_ids:
            raise ValueError(
                f"Duplicate WP id '{id_str}' in WP BUDGET section. Multi-segment "
                "schedules are no longer supported — each WP must appear exactly "
                "once. Collapse to a single (Start, End) range."
            )
        seen_ids.add(id_str)
        start_val = ws.cell(r, start_col).value
        end_val = ws.cell(r, end_col).value
        type_val = ws.cell(r, type_col).value
        if start_val is None or end_val is None:
            raise ValueError(
                f"WP '{id_str}' is missing Start or End in WP BUDGET. Blank "
                "cells are not a valid placeholder under the single-range schema."
            )
        start_int = int(start_val)
        end_int = int(end_val)
        type_str = str(type_val).strip() if type_val else "flat"
        if not (1 <= start_int <= end_int):
            raise ValueError(
                f"WP '{id_str}' has invalid schedule range: Start={start_int}, "
                f"End={end_int}. Required: 1 <= Start <= End."
            )
        if project_months and end_int > project_months:
            raise ValueError(
                f"WP '{id_str}' End={end_int} exceeds project_months="
                f"{project_months}. Shorten the WP or extend the project."
            )
        if type_str not in valid_types:
            raise ValueError(f"WP '{id_str}' has invalid Type {type_str!r}. Must be one of: {sorted(valid_types)}.")
        fte_val = float(ws.cell(r, COL_WP_FTE).value or 0)
        rate_numeric = _safe_float(rate_cell, 6500)
        if is_legacy_ceiling_layout:
            # Legacy Ceiling = FTE×Rate; migrate to Budget +i with ×1.25 uplift.
            budget_plus_i = round(fte_val * rate_numeric * 1.25)
            legacy_rows.append((id_str, budget_plus_i))
        else:
            # Post-restructure: Budget +i cell. _safe_float falls back to
            # fte*rate*1.25 when the structured-ref hasn't been recalc'd.
            plus_i_raw = ws.cell(r, COL_WP_BUDGET_PLUS_I).value
            budget_plus_i = round(_safe_float(plus_i_raw, fte_val * rate_numeric * 1.25))
        # ``rate``: numeric coercion (6500 fallback on formula cells).
        # ``rate_raw``: literal cell — preserves user formula on round-trip.
        work_packages.append(
            {
                "id": id_str,
                "name": str(col_b or ""),
                "fte": fte_val,
                "rate": rate_numeric,
                "rate_raw": rate_cell,
                "budget_plus_i": budget_plus_i,
                "start": start_int,
                "end": end_int,
                "type": type_str,
            }
        )
        r += 1

    # Sanity guard: double-applied ×1.25 produces >€3M per WP — require
    # GRANT_ACCEPT_CEILING_MIGRATION=1 opt-in.
    if is_legacy_ceiling_layout:
        over_cap = [(wp_id, amount) for wp_id, amount in legacy_rows if amount > 3_000_000]
        if over_cap and os.environ.get("GRANT_ACCEPT_CEILING_MIGRATION") != "1":
            offenders = ", ".join(f"{wp}={amount:,.0f}" for wp, amount in over_cap)
            raise RuntimeError(
                "Legacy Ceiling→Budget migration would seed Budget +i values "
                f"above €3,000,000 per WP: {offenders}. Likely double-apply "
                "(source already holds with-indirect values). Verify, then "
                "set GRANT_ACCEPT_CEILING_MIGRATION=1 to override."
            )
        warnings.warn(
            f"Migrated {len(legacy_rows)} WP Ceiling value(s) to Budget +i "
            "(×1.25 uplift). Col 5 now derives Budget -i; col 6 is authoritative "
            "with-indirect Budget +i.",
            RuntimeWarning,
            stacklevel=2,
        )

    return work_packages


def _read_project_params(ws) -> dict:
    """Scan Config col A -> {title, start, months, rate, eu_ceiling_pm,
    vacation_month_num}. Rejects legacy two-section ("WP SCHEDULE") layouts."""
    # Legacy two-section layout — pre-merge "WP SCHEDULE" header within ~200 rows.
    for r in range(1, min(ws.max_row, 200) + 1):
        val = ws.cell(row=r, column=1).value
        if isinstance(val, str) and val.strip() == "WP SCHEDULE":
            raise ValueError(
                f"Legacy Config layout detected: 'WP SCHEDULE' section at row "
                f"{r}. Multi-segment WP schedules are no longer supported. "
                "Migration: delete the WP SCHEDULE section; add Start/End/Type "
                "columns (H/I/J) to WP BUDGET, one row per WP — collapse "
                "multi-range WPs to single (start, end)."
            )

    # Scan column A for known param labels (position-independent). "Vacation
    # Month" is optional — resolved via named range below (default 7).
    _param_map = {
        "Project Title": ("title", None),
        "Project Start": ("start", None),
        "Project Months": ("months", lambda v: int(v or 60)),
        "FTE Rate": ("rate", _parse_fte_rate),
        "EU Annual Ceiling": ("eu_ceiling_pm", lambda v: float(v or 12.0)),
    }
    project_params: dict = {}
    matched_keys: set[str] = set()
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        for prefix, (key, converter) in _param_map.items():
            if prefix.lower() in label.lower():
                raw = ws.cell(r, 2).value
                project_params[key] = converter(raw) if converter else raw
                matched_keys.add(key)
                break
        if len(matched_keys) == len(_param_map):
            break
    for prefix, (key, _) in _param_map.items():
        if key not in project_params:
            raise ValueError(f"Config sheet missing param row matching '{prefix}'")

    project_params["vacation_month_num"] = _read_vacation_month_num(ws)
    return project_params


def read_config_sheet(ws) -> tuple[dict, list[dict]]:
    """Read Config sheet -> (project_params, work_packages). See
    ``_read_project_params`` and ``_read_wp_budget_section`` for shapes and
    raise conditions. ``rate`` in params is a fallback only — the builder
    regenerates B4 from the live grade/roster layout."""
    project_params = _read_project_params(ws)
    project_months = int(project_params.get("months") or 0)
    work_packages = _read_wp_budget_section(ws, project_months)
    return project_params, work_packages


def read_plan_config(ws, layout) -> list[dict]:
    """Read WP definitions from Config (via Plan layout cross-refs). Uses
    ``layout["config_first_wp_row"..]`` since the Plan config table was removed."""
    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None
    if config_ws is None:
        raise ValueError("Workbook has no Config sheet")

    wps = []
    for row in range(layout["config_first_wp_row"], layout["config_last_wp_row"] + 1):
        wp_id = config_ws.cell(row, 1).value
        wp_name = config_ws.cell(row, 2).value
        fte = config_ws.cell(row, 3).value
        if wp_id is None:
            continue
        fte_parsed = _parse_config_float(fte)
        wps.append(
            {
                "id": str(wp_id).strip(),
                "name": str(wp_name or "").strip(),
                "fte": fte_parsed if fte_parsed is not None else 0.0,
            }
        )
    if not wps:
        raise ValueError("Config table has zero WP rows. Regenerate from defaults with: python -m budget_planner")
    return wps


def read_grade_rates(ws) -> dict[str, int | float]:
    """Grade -> EUR/PM rate (e.g. {"Senior": 12500, ...}). Falls back to
    ``budget_planner.config.GRADE_RATES`` when the RATE BY GRADE section is missing
    (legacy workbooks).

    A row with a grade name but a non-numeric rate is a data error, not a
    section terminator: raising here names the offending cell instead of
    silently truncating the dict and surfacing as a ``KeyError`` deep inside
    the Timelog solver."""
    section_row = _find_header_row(ws, SECTION_HEADER_LABELS["rate_by_grade"], required=False)
    if section_row is None:
        from budget_planner.config import GRADE_RATES

        return dict(GRADE_RATES)

    r = _data_row_after(section_row)
    rates: dict[str, int | float] = {}
    while r <= ws.max_row:
        grade_name = ws.cell(r, 1).value
        if grade_name is None or str(grade_name).strip() == "":
            break
        grade_name = str(grade_name).strip()
        rate_val = ws.cell(r, COL_GRADE_RATE).value
        # bool is an int subclass — TRUE would otherwise read as a rate of 1.
        if rate_val is None or isinstance(rate_val, bool) or not isinstance(rate_val, (int, float)):
            raise ValueError(_grade_rate_error(ws, r, grade_name, rate_val))
        rates[grade_name] = rate_val
        r += 1

    return rates


def _grade_rate_error(ws, row: int, grade_name: str, rate_val) -> str:
    """Human-actionable message for a RATE BY GRADE cell that is not a literal number."""
    cell_ref = f"{ws.title}!{get_column_letter(COL_GRADE_RATE)}{row}"
    if isinstance(rate_val, str) and rate_val.startswith("="):
        found = f"the formula {rate_val!r}"
    elif rate_val is None or (isinstance(rate_val, str) and rate_val.strip() == ""):
        found = "an empty cell"
    else:
        found = f"{rate_val!r}"
    return (
        f"RATE BY GRADE row {row} (grade {grade_name!r}): rate must be a literal number "
        f"(blue input cell), found {found} in {cell_ref}. Open the workbook in Excel, replace "
        f"that cell with its computed value, and re-run."
    )


def read_team_from_config(ws) -> list[dict]:
    """Read TEAM ROSTER -> list of {name, max_fte, grade, priority}. ``priority``
    is int 1-3, default 2 when the column is missing/blank (pre-Timelog-Plan
    workbooks)."""
    from budget_planner.config import DEFAULT_GRADE

    roster_section = _find_header_row(ws, SECTION_HEADER_LABELS["team_roster"], required=True)

    # Detect Grade / Priority columns by header inspection (legacy compat).
    person_header = roster_section + 1
    grade_hdr = ws.cell(person_header, COL_ROSTER_GRADE).value
    has_grade_col = bool(grade_hdr) and "grade" in str(grade_hdr).lower()
    prio_hdr = ws.cell(person_header, COL_ROSTER_PRIORITY).value
    has_priority_col = bool(prio_hdr) and "priority" in str(prio_hdr).lower()

    roster: list[dict] = []
    for r, max_fte_parsed in _iter_config_section_rows(
        ws,
        header_row=roster_section,
        named_range="roster_ftes",
        key_col=1,
        numeric_col=COL_ROSTER_FTE,
    ):
        name = ws.cell(r, 1).value
        grade = DEFAULT_GRADE
        if has_grade_col:
            grade_val = ws.cell(r, COL_ROSTER_GRADE).value
            if grade_val is not None and str(grade_val).strip():
                grade = str(grade_val).strip()
        priority = 2
        if has_priority_col:
            prio_val = ws.cell(r, COL_ROSTER_PRIORITY).value
            if prio_val is not None and str(prio_val).strip():
                try:
                    priority = int(float(prio_val))
                except (TypeError, ValueError):
                    priority = 2
                # Clamp to [1,3] — solver iterates prio in (1,2,3); out-of-band
                # values would silently exclude the person from all WP passes.
                priority = max(1, min(3, priority))
        roster.append(
            {
                "name": str(name).strip(),
                "max_fte": max_fte_parsed,
                "grade": grade,
                "priority": priority,
            }
        )

    return roster


def read_reporting_periods(ws, fallback_months=None):
    """Read REPORTING PERIODS -> list of {id, start, end}. When the section
    is missing: generate defaults for ``fallback_months`` if provided, else []."""
    rp_section = _find_header_row(ws, SECTION_HEADER_LABELS["reporting_periods"], required=False)
    if rp_section is None:
        if fallback_months is not None:
            from budget_planner.config import default_reporting_periods

            return default_reporting_periods(fallback_months)
        return []

    r = _data_row_after(rp_section)
    periods: list[dict] = []
    while r <= ws.max_row:
        period_id = ws.cell(r, 1).value
        if period_id is None or str(period_id).strip() == "":
            break
        start = ws.cell(r, COL_RP_START).value
        end = ws.cell(r, COL_RP_END).value
        periods.append(
            {
                "id": str(period_id),
                "start": int(start or 1),
                "end": int(end or 1),
            }
        )
        r += 1

    return periods


def read_assignments_from_config(ws) -> list[dict]:
    """Read PERSON-WP ASSIGNMENTS -> list of {person, wp, total_fte, start, end}.
    Empty list if the section is missing."""
    assignments_header = _find_header_row(ws, SECTION_HEADER_LABELS["assignments"], required=False)
    if assignments_header is None:
        return []

    assignments: list[dict] = []
    for r, total_fte_parsed in _iter_config_section_rows(
        ws,
        header_row=assignments_header,
        named_range="assign_persons",
        key_col=COL_ASSIGN_PERSON,
        numeric_col=COL_ASSIGN_FTE,
    ):
        person = ws.cell(r, COL_ASSIGN_PERSON).value
        wp = ws.cell(r, COL_ASSIGN_WP).value
        start = ws.cell(r, COL_ASSIGN_START).value
        end = ws.cell(r, COL_ASSIGN_END).value
        assignments.append(
            {
                "person": str(person).strip(),
                "wp": str(wp or "").strip(),
                "total_fte": total_fte_parsed,
                "start": int(start or 1),
                "end": int(end or 1),
            }
        )

    return assignments


def read_externals_from_config(ws) -> list[dict]:
    """Read EXTERNAL COMMITMENTS -> list of {person, project, fte, start, end,
    is_eu}. Empty list if the section is missing."""
    ext_config_header = _find_header_row(ws, SECTION_HEADER_LABELS["externals"], required=False)
    if ext_config_header is None:
        return []

    external_commitments: list[dict] = []
    for r, fte_parsed in _iter_config_section_rows(
        ws,
        header_row=ext_config_header,
        named_range="ext_persons",
        key_col=COL_EXT_PERSON,
        numeric_col=COL_EXT_FTE,
    ):
        person = ws.cell(r, COL_EXT_PERSON).value
        project = ws.cell(r, COL_EXT_PROJECT).value
        start = ws.cell(r, COL_EXT_START).value
        end = ws.cell(r, COL_EXT_END).value
        is_eu_val = ws.cell(r, COL_EXT_EU).value
        external_commitments.append(
            {
                "person": str(person).strip(),
                "project": str(project or "").strip(),
                "fte": fte_parsed,
                "start": int(start or 1),
                "end": int(end or 1),
                "is_eu": str(is_eu_val).strip().upper() in ("YES", "TRUE", "1"),
            }
        )

    return external_commitments
