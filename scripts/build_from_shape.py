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
    SHEET_META,
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

    # Grade rates — "Mid" is the DEFAULT_GRADE used by builders; ensure it exists
    # whenever the fingerprint includes at least one grade row.
    grade_rates: dict[str, int] = {}
    if n_grades >= 1:
        grade_rates["Mid"] = 6500
        for i in range(1, n_grades):
            grade_rates[f"Grade_{i + 1}"] = 6500 + i * 500

    work_packages: list[dict[str, Any]] = [
        {"id": f"WP{i + 1}", "name": f"Work package {i + 1}", "fte": 10.0}
        for i in range(n_wps)
    ]

    # Schedule entries point at existing WPs; if there are no WPs, schedule
    # cannot be meaningfully populated — skip seeding to avoid a foreign-key
    # mismatch that would fail _validate_config. The fingerprint's schedule
    # count is meaningless without any WPs anyway.
    wp_schedule: list[dict[str, Any]] = []
    if n_wps > 0:
        wp_schedule = [
            {
                "id": work_packages[i % n_wps]["id"],
                "start": 1,
                "end": max(1, months),
                "type": "flat",
            }
            for i in range(n_sched)
        ]

    team: list[dict[str, Any]] = [
        {"name": f"Person_{i + 1:02d}", "max_fte": 1.0, "grade": "Mid"}
        for i in range(n_team)
    ]

    external_commitments: list[dict[str, Any]] = []
    if n_team > 0:
        external_commitments = [
            {
                "person": team[i % n_team]["name"],
                "project": f"ExtProj_{i + 1}",
                "fte": 1.0,
                "start": 1,
                "end": max(1, months),
                "is_eu": False,
            }
            for i in range(n_ext)
        ]

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

    assignments: list[dict[str, Any]] = []
    if n_team > 0 and n_wps > 0:
        assignments = [
            {
                "person": team[i % n_team]["name"],
                "wp": work_packages[i % n_wps]["id"],
                "start": 1,
                "end": max(1, months),
                "total_fte": 1.0,
            }
            for i in range(n_assign)
        ]

    project_params = {
        "title": "Synthetic",
        "start": datetime(2025, 1, 1),
        "months": months,
        "rate": 10000,
        "eu_ceiling_pm": 12.0,
    }

    wb = Workbook()
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = SHEET_META

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
