"""Synthesize a workbook from a shape fingerprint.

Given a fingerprint dict produced by ``scripts.shape_fingerprint.extract_fingerprint``,
build a fresh ``.xlsx`` whose Config sections match the fingerprint's row counts.
All cell values are deterministic placeholders (``Person_01``, ``WP1``, ``6500``,
fixed dates) — safe to share as a bug reproducer.

The real builder entry point is ``budget_planner.workbook._build_all_sheets`` — we
seed module-shaped data (lists of dicts matching ``budget_planner/config.py``) and
dispatch directly into it, bypassing the config-module defaults that
``init_workbook`` uses.
"""
from __future__ import annotations

import os
from datetime import datetime
from typing import Any

from openpyxl import Workbook

from budget_planner.config import PROJECT_MONTHS
from budget_planner.workbook import (
    SHEET_CONFIG,
    _apply_all_year_groups,
    _build_all_sheets,
)


def _section_count(fp: dict, key: str) -> int:
    section = fp.get("config_sections", {}).get(key)
    if not section:
        return 0
    return int(section.get("row_count", 0))


def build_from_fingerprint(fp: dict, out_path: str) -> str:
    """Build a synthetic workbook whose Config section row counts match ``fp``.

    Only row counts are reproduced; cell values are deterministic placeholders.
    """
    n_wps = _section_count(fp, "work_packages")
    n_sched = _section_count(fp, "wp_schedule")
    n_team = _section_count(fp, "team_roster")
    n_assign = _section_count(fp, "assignments")
    n_ext = _section_count(fp, "external_commitments")
    n_rp = _section_count(fp, "reporting_periods")
    n_grades = _section_count(fp, "grade_rates")

    months = max(PROJECT_MONTHS, n_sched if n_sched > 0 else 1)

    # Grade rates — realistic programme bookkeeping bands. "Mid" is the DEFAULT_GRADE
    # and must exist whenever the fingerprint has any grade rows so team
    # members can reference it. Extra grades beyond the core three get
    # deterministic spread rates so each band differs (realistic recalc).
    _CORE_GRADE_RATES = [("Senior", 12500), ("Mid", 10000), ("Junior", 7500)]
    grade_rates: dict[str, int] = {}
    if n_grades >= 1:
        grade_rates["Mid"] = 10000
        core_extras = [g for g in _CORE_GRADE_RATES if g[0] != "Mid"]
        for i in range(1, n_grades):
            if i - 1 < len(core_extras):
                name, rate = core_extras[i - 1]
                grade_rates[name] = rate
            else:
                extra = i - 1 - len(core_extras)
                grade_rates[f"Grade_{i + 1}"] = 4500 + extra * 900

    grade_cycle = list(grade_rates.keys()) or ["Mid"]

    # Work packages — vary FTE budgets so weighted averages and WP totals
    # differ across rows instead of collapsing to a single repeated number.
    _WP_FTE_CYCLE = [4.0, 6.0, 8.0, 10.0, 12.0, 6.0, 8.0, 4.0, 10.0]
    work_packages: list[dict[str, Any]] = [
        {
            "id": f"WP{i + 1}",
            "name": f"Work package {i + 1}",
            "fte": _WP_FTE_CYCLE[i % len(_WP_FTE_CYCLE)],
        }
        for i in range(n_wps)
    ]

    # Schedule entries point at existing WPs; if there are no WPs, schedule
    # cannot be meaningfully populated — skip seeding to avoid a foreign-key
    # mismatch that would fail _validate_config. The fingerprint's schedule
    # count is meaningless without any WPs anyway.
    #
    # Vary distribution type AND active window per schedule row so capacity
    # distribution curves and inactive-month greying exercise real code paths
    # (flat vs front/back-loaded vs bell) during recalc.
    _DIST_TYPES = ["flat", "front-loaded", "bell", "back-loaded"]
    wp_schedule: list[dict[str, Any]] = []
    if n_wps > 0:
        for i in range(n_sched):
            dist = _DIST_TYPES[i % len(_DIST_TYPES)]
            if dist == "flat":
                start, end = 1, months
            elif dist == "front-loaded":
                start, end = 1, max(1, months // 2)
            elif dist == "bell":
                start = max(1, months // 4)
                end = max(start, (3 * months) // 4)
            else:  # back-loaded
                start = max(1, months // 2)
                end = months
            wp_schedule.append(
                {
                    "id": work_packages[i % n_wps]["id"],
                    "start": start,
                    "end": end,
                    "type": dist,
                }
            )

    # Inject a WP-gap case so synthetic workbooks exercise the gap-aware
    # Capacity distribution path (config-sheet helpers + Capacity grey-out
    # must both respect WP SCHEDULE gaps). Requires >=2 schedule slots: we
    # rewrite slot 0 as WP1 phase-1 and the last slot as WP1 phase-2, with
    # a deliberate gap between them.
    gap_wp_id: str | None = None
    gap_window: tuple[int, int] | None = None
    if n_wps > 0 and n_sched >= 2:
        phase1_end = max(1, months // 4)
        phase2_start = phase1_end + max(2, months // 6)
        phase2_end = min(months, phase2_start + max(1, months // 4))
        if phase2_start <= months and phase2_end >= phase2_start:
            gap_wp_id = work_packages[0]["id"]
            wp_schedule[0] = {
                "id": gap_wp_id, "start": 1, "end": phase1_end,
                "type": "front-loaded",
            }
            wp_schedule[-1] = {
                "id": gap_wp_id, "start": phase2_start, "end": phase2_end,
                "type": "front-loaded",
            }
            gap_window = (1, phase2_end)

    # Team — rotate through available grades + vary max_fte (some part-time).
    _MAX_FTE_CYCLE = [1.0, 1.0, 0.8, 1.0, 0.9, 1.0, 0.8, 1.0]
    team: list[dict[str, Any]] = [
        {
            "name": f"Person_{i + 1:02d}",
            "max_fte": _MAX_FTE_CYCLE[i % len(_MAX_FTE_CYCLE)],
            "grade": grade_cycle[i % len(grade_cycle)],
        }
        for i in range(n_team)
    ]

    # External commitments — vary FTE load, window, and EU-vs-non-EU flag so
    # the EU annual ceiling check has both contributing and non-contributing
    # rows to aggregate.
    _EXT_FTE_CYCLE = [3.0, 6.0, 4.0, 2.0, 5.0]
    external_commitments: list[dict[str, Any]] = []
    if n_team > 0:
        for i in range(n_ext):
            ext_end = max(1, months - (i % 3) * (max(1, months // 6)))
            external_commitments.append(
                {
                    "person": team[i % n_team]["name"],
                    "project": f"ExtProj_{i + 1}",
                    "fte": _EXT_FTE_CYCLE[i % len(_EXT_FTE_CYCLE)],
                    "start": 1 + (i % 3),
                    "end": ext_end,
                    "is_eu": (i % 2 == 0),
                }
            )

    # Reporting periods — equal slices of the project months; single-period fallback.
    reporting_periods: list[dict[str, Any]] = []
    if n_rp == 1:
        reporting_periods = [{"id": "RP1", "start": 1, "end": max(1, months)}]
    elif n_rp > 1:
        slice_len = max(1, months // n_rp)
        cur = 1
        for i in range(n_rp):
            if i == n_rp - 1:
                end = max(cur, months)
            else:
                end = min(months, cur + slice_len - 1)
            # Guard against start > end at the tail of very small month counts.
            if cur > months:
                cur = months
            if end < cur:
                end = cur
            reporting_periods.append({"id": f"RP{i + 1}", "start": cur, "end": end})
            cur = end + 1

    # Assignments — vary total_fte so planned monthly EUR differs per row, and
    # align each assignment's active window with its WP's schedule entry so the
    # rebuild-round-trip sees realistic inactive-month patterns per person/WP.
    _ASSIGN_FTE_CYCLE = [2.0, 4.0, 3.0, 6.0, 5.0, 1.0, 8.0, 2.5, 1.5, 3.5]
    _sched_by_wp: dict[str, dict[str, Any]] = {}
    for entry in wp_schedule:
        _sched_by_wp.setdefault(entry["id"], entry)
    assignments: list[dict[str, Any]] = []
    if n_team > 0 and n_wps > 0:
        for i in range(n_assign):
            wp = work_packages[i % n_wps]
            sched = _sched_by_wp.get(wp["id"])
            if gap_wp_id is not None and wp["id"] == gap_wp_id and gap_window is not None:
                # Assignment on the gap WP spans both phases so Capacity
                # exercises the gap-aware distribution path on recalc.
                a_start, a_end = gap_window
            elif sched is not None:
                a_start, a_end = int(sched["start"]), int(sched["end"])
            else:
                a_start, a_end = 1, max(1, months)
            assignments.append(
                {
                    "person": team[i % n_team]["name"],
                    "wp": wp["id"],
                    "start": a_start,
                    "end": a_end,
                    "total_fte": _ASSIGN_FTE_CYCLE[i % len(_ASSIGN_FTE_CYCLE)],
                }
            )

    project_params = {
        "title": "Synthetic",
        "start": datetime(2025, 1, 1),
        "months": months,
        "rate": 10000,
        "eu_ceiling_pm": 12.0,
    }

    wb = Workbook()

    _build_all_sheets(
        wb,
        project_params,
        work_packages,
        team,
        external_commitments,
        wp_schedule=wp_schedule,
        assignments=assignments,
        reporting_periods=reporting_periods,
        grade_rates=grade_rates,
    )

    _apply_all_year_groups(wb, project_params, current_month=1)

    wb.active = wb.sheetnames.index(SHEET_CONFIG)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)
    return out_path
