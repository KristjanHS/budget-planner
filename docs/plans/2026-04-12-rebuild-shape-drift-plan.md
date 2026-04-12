# Rebuild Shape-Drift Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Make `rebuild` robust to Excel-side row insertions in Config-tab tables (Team, Assignments, Externals) so no person/assignment is silently dropped and cross-sheet formulas resolve to live positions.

**Architecture:** Replace "first-empty-row terminates" table readers with **named-range-bounded** readers that use `wb.defined_names` as the authoritative row span (Excel auto-extends these on insertion). Fall back to scan-with-empty-tolerance only where no named range exists. Keep the discovered shape as the single source of truth; `_META` text stays documentation-only (already rewritten every build from live layout).

**Tech Stack:** openpyxl, pytest, Python 3.12+.

**Design doc:** `docs/plans/2026-04-12-rebuild-shape-drift-design.md`

**Pre-read before starting:**
- `budget_planner/reader.py` — the three reader functions at lines 328, 425, 469
- `budget_planner/config_sheet.py` — to locate where named ranges for roster/assignments/externals are registered
- `tests/test_rebuild.py` and `tests/test_rebuild_e2e.py` — existing rebuild test patterns
- Project rule `.claude/rules/dev-workflow.md` — round-trip test discipline (init → openpyxl edit → rebuild → reopen → assert)

**Conventions reminder (CLAUDE.md Critical Rules):**
- Never hardcode calculated values — formulas only.
- Font color: blue=input / black=formula / green=cross-sheet.
- After every generation, run the full suite (`pytest tests/ -q`).
- `_META` must reflect current live structure (already handled — verify).

---

## Task Order (quick overview)

| # | Task | Purpose |
|---|------|---------|
| T1 | Reproduction test (RED) | Pin the exact bug with openpyxl mid-insert |
| T2 | Audit named ranges in builder | Know which tables have usable ranges |
| T3 | Introduce `_iter_named_range_rows` helper | Shared row-span resolver from `defined_names` |
| T4 | Fix `read_team_from_config` | Use roster_ftes span, tolerate internal blanks |
| T5 | Fix `read_assignments_from_config` | Use assign_persons span |
| T6 | Fix `read_externals_from_config` | Use ext_persons span |
| T7 | T1 turns GREEN + make permanent | Bug is fixed; test stays as regression guard |
| T8 | Shape-parameterized rebuild tests (lean) | default + one variant per table = 5 cases |
| T9 | META-divergence invariant helper | Catches future META-trusting regressions |
| T10 | Fingerprint schema + extractor | Machine-readable structural snapshot (no cell values) |
| T11 | Synthetic workbook generator | Deterministic reproducer from a fingerprint |
| T12 | One-shot wrapper `make_reproducer.py` | Fully automated: real.xlsx → safe synthetic.xlsx |
| T13 | Privacy canary test | Fail loud if any real cell value ever leaks into synthetic |
| T14 | Debugging runbook doc | User-facing "when you hit a bug, do this" |
| T15 | Full-suite sweep + trigger-table updates | Ensure no regressions |

Each task ends with a commit. Run `pytest tests/ -q` after each code-changing task; run `pyright` before the final commit of each task.

---

## Task 1: Reproduction test (RED)

**Files:**
- Create: `tests/test_rebuild_shape_drift.py`

**Purpose:** Pin the exact bug the user reported — inserting rows in the middle of the Team roster causes rebuild to drop persons. This test must FAIL on current code; T7 will flip it green.

**Step 1: Write the failing test**

```python
"""Round-trip tests for Excel-side row insertions in Config tables.

Mirrors the user's workflow: generate a workbook, hand-edit Config in Excel
(simulated via openpyxl), run rebuild, assert no data loss.
"""
from pathlib import Path

import openpyxl
import pytest

from budget_planner.workbook import build_workbook, rebuild_workbook
from budget_planner.reader import (
    read_team_from_config,
    read_assignments_from_config,
)
from budget_planner import config as cfg


def _find_roster_first_data_row(ws):
    """Locate the first data row of the TEAM ROSTER section."""
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == "TEAM ROSTER":
            # Section header row, then column header row, then first data row.
            return r + 2
    raise AssertionError("TEAM ROSTER section not found")


def test_mid_insert_in_team_roster_preserved_through_rebuild(tmp_path: Path):
    """User inserts a new person row in the middle of Team Roster; rebuild must keep them."""
    path = tmp_path / "wb.xlsx"
    build_workbook(str(path))  # init from config.py defaults

    # Simulate the user opening Excel and inserting 1 row mid-roster.
    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    first_data = _find_roster_first_data_row(config_ws)
    insert_at = first_data + 1  # middle of the roster
    config_ws.insert_rows(insert_at, amount=1)
    config_ws.cell(insert_at, 1, "Midway Mary")
    config_ws.cell(insert_at, 2, 0.8)
    # Grade column — pick any existing grade from config
    config_ws.cell(insert_at, 3, next(iter(cfg.GRADE_RATES.keys())))
    wb.save(path)
    wb.close()

    # Sanity: reader should see the inserted person BEFORE rebuild.
    wb = openpyxl.load_workbook(path)
    roster = read_team_from_config(wb["Config"])
    wb.close()
    names = [p["name"] for p in roster]
    assert "Midway Mary" in names, (
        f"Pre-rebuild read already loses inserted person — reader bug. "
        f"Roster read: {names}"
    )

    # Run rebuild — the code path that currently drops people.
    rebuild_workbook(str(path))

    # After rebuild, the inserted person must still be present.
    wb = openpyxl.load_workbook(path)
    roster_after = read_team_from_config(wb["Config"])
    assignments_after = read_assignments_from_config(wb["Config"])
    wb.close()
    names_after = [p["name"] for p in roster_after]
    assert "Midway Mary" in names_after, (
        f"Rebuild dropped the mid-inserted person. Roster after rebuild: {names_after}"
    )
    # Any pre-existing assignment for post-insertion persons must survive too.
    dropped_assignees = {a["person"] for a in assignments_after} - set(names_after)
    assert not dropped_assignees, (
        f"Rebuild has assignments pointing at persons no longer in roster: {dropped_assignees}"
    )
```

**Step 2: Run test to verify it fails**

```
pytest tests/test_rebuild_shape_drift.py::test_mid_insert_in_team_roster_preserved_through_rebuild -v
```

Expected: FAIL. The exact failure may be at either the pre-rebuild assertion (reader drops on internal blank) or post-rebuild assertion (rebuild round-trip drops). Either is acceptable — both paths will be fixed. If the test *passes* on current code, **stop and investigate** — the bug hypothesis is wrong and this plan needs revision.

**Step 3: Commit the RED test**

```bash
git add tests/test_rebuild_shape_drift.py
git commit -m "test: add reproduction for mid-insert row drop on rebuild (RED)"
```

---

## Task 2: Audit named ranges in the builder

**Files:**
- Read-only: `budget_planner/config_sheet.py`
- Read-only: `budget_planner/workbook.py`

**Purpose:** Confirm which Config tables actually have named ranges and what each span is keyed on. This informs T3's API. No code written yet.

**Step 1: Enumerate named ranges**

```
grep -n "defined_names\|DefinedName\|create_named_range\|named_range" budget_planner/config_sheet.py budget_planner/workbook.py budget_planner/layout.py
```

Expected findings (from meta.py documentation):

- `roster_ftes` — `Config!$B$<first>:$B$<last>` (team Max FTE column)
- `roster_grades` — `Config!$C$<first>:$C$<last>`
- `assign_persons`, `assign_wps`, `assign_ftes` — assignment table columns
- `ext_persons`, `ext_ftes` — externals table columns
- `sched_wp_ids`, `sched_starts`, `sched_ends` — WP schedule
- `tbl_GradeRates` — `Config!$A$<first>:$B$<last>`

**Step 2: Record findings in a short note inside this plan file**

Append a **"Named-range audit results"** section to the bottom of *this plan file* listing the actual names/spans found, so T3–T6 reference real names. This is a plan update, not a code commit.

**Step 3: Commit plan update**

```bash
git add docs/plans/2026-04-12-rebuild-shape-drift-plan.md
git commit -m "docs(plan): record named-range audit for shape-drift fix"
```

---

## Task 3: Shared `_iter_named_range_rows` helper

**Files:**
- Modify: `budget_planner/reader.py` (add a helper near `_find_header_row`, before line 173)
- Create test in: `tests/test_reader_new.py` (append a new test function; this file already groups reader unit tests)

**Step 1: Write the failing test**

```python
def test_iter_named_range_rows_returns_live_row_span_after_middle_insert(tmp_path):
    """Helper must report the post-insertion row span via Excel's defined_names."""
    import openpyxl
    from openpyxl.workbook.defined_name import DefinedName
    from budget_planner.reader import _iter_named_range_rows

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Config"
    for i, val in enumerate(["A", "B", "C"], start=5):
        ws.cell(i, 1, val)
    # Named range B5:B7 (1-based, inclusive), but registered globally on wb.
    wb.defined_names["roster_ftes"] = DefinedName(
        name="roster_ftes", attr_text="Config!$B$5:$B$7"
    )

    # Excel-style middle insert: shift rows 6+ down by 1.
    ws.insert_rows(6, amount=1)
    ws.cell(6, 1, "MID")
    # openpyxl does NOT update DefinedName targets on insert — simulate Excel
    # by explicitly rewriting the range. The helper under test must tolerate
    # that the range given in defined_names is the SOURCE OF TRUTH.
    wb.defined_names["roster_ftes"].attr_text = "Config!$B$5:$B$8"

    first, last = _iter_named_range_rows(wb, "roster_ftes")
    assert (first, last) == (5, 8)
```

**Step 2: Run test to verify it fails**

```
pytest tests/test_reader_new.py::test_iter_named_range_rows_returns_live_row_span_after_middle_insert -v
```

Expected: FAIL with `AttributeError` / `ImportError` (helper does not exist).

**Step 3: Implement the helper**

Insert into `budget_planner/reader.py` just above `_find_header_row` (~line 115):

```python
def _iter_named_range_rows(wb, name: str) -> tuple[int, int] | None:
    """Resolve a defined name to its (first_row, last_row) on the first target range.

    Returns None if the name is not defined or cannot be parsed. Callers fall back
    to sentinel-scan reading when this returns None.
    """
    dn = wb.defined_names.get(name) if hasattr(wb.defined_names, "get") else None
    if dn is None:
        # openpyxl <3.1 style
        try:
            dn = wb.defined_names[name]
        except (KeyError, TypeError):
            return None
    try:
        destinations = list(dn.destinations)
    except Exception:
        return None
    if not destinations:
        return None
    # destinations = [(sheet_title, "$B$5:$B$8"), ...]. Take the first.
    _sheet, ref = destinations[0]
    # Parse absolute ref "$B$5:$B$8" or "$B$5".
    import re
    m = re.match(r"\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?$", ref)
    if not m:
        return None
    first = int(m.group(1))
    last = int(m.group(2) or m.group(1))
    return (first, last)
```

**Step 4: Run test to verify it passes**

```
pytest tests/test_reader_new.py::test_iter_named_range_rows_returns_live_row_span_after_middle_insert -v
```

Expected: PASS.

**Step 5: Commit**

```bash
git add budget_planner/reader.py tests/test_reader_new.py
git commit -m "feat(reader): add _iter_named_range_rows helper for live row spans"
```

---

## Task 4: Fix `read_team_from_config` — named-range primary, scan fallback

**Files:**
- Modify: `budget_planner/reader.py` — function `read_team_from_config` (line 328)
- Tests: extend `tests/test_reader_new.py` with a gap-tolerance test

**Step 1: Write failing tests**

```python
def test_read_team_tolerates_blank_row_inside_named_range(tmp_path):
    """A blank row inside roster_ftes span must not truncate the roster."""
    import openpyxl
    from openpyxl.workbook.defined_name import DefinedName
    from budget_planner.reader import read_team_from_config

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Config"
    # Minimum section header prelude expected by the reader.
    ws.cell(10, 1, "TEAM ROSTER")
    ws.cell(11, 1, "Person"); ws.cell(11, 2, "Max FTE"); ws.cell(11, 3, "Grade")
    ws.cell(12, 1, "Alice"); ws.cell(12, 2, 1.0); ws.cell(12, 3, "Senior")
    # Row 13 intentionally blank.
    ws.cell(14, 1, "Bob");   ws.cell(14, 2, 0.5); ws.cell(14, 3, "Junior")
    wb.defined_names["roster_ftes"] = DefinedName(
        name="roster_ftes", attr_text="Config!$B$12:$B$14"
    )

    roster = read_team_from_config(ws)

    names = [p["name"] for p in roster]
    assert "Alice" in names
    assert "Bob" in names, (
        "Blank row inside named range must not terminate reader. "
        f"Got: {names}"
    )
```

**Step 2: Run test to verify it fails**

```
pytest tests/test_reader_new.py::test_read_team_tolerates_blank_row_inside_named_range -v
```

Expected: FAIL — "Bob" missing.

**Step 3: Implement the fix**

Modify `read_team_from_config` in `budget_planner/reader.py` (around line 328). Replace the body after `person_header` detection with:

```python
    # Prefer named-range-bounded scan; fall back to sentinel scan when
    # roster_ftes is absent (e.g., legacy workbooks from before named ranges).
    wb = ws.parent
    span = _iter_named_range_rows(wb, "roster_ftes")
    if span is not None:
        first_data, last_data = span
    else:
        first_data = person_header + 1
        last_data = None  # sentinel-scan

    roster: list[dict] = []
    r = first_data
    end = last_data if last_data is not None else ws.max_row
    while r <= end:
        name = ws.cell(r, 1).value
        if (name is None or str(name).strip() == "") and last_data is None:
            break  # sentinel scan: stop at first blank
        max_fte = ws.cell(r, COL_ROSTER_FTE).value
        max_fte_parsed = _parse_config_float(max_fte)
        # Inside a named range, tolerate blanks/non-numeric: skip the row.
        if last_data is not None and (
            name is None or str(name).strip() == "" or max_fte_parsed is None
        ):
            r += 1
            continue
        if max_fte_parsed is None:
            break
        grade = DEFAULT_GRADE
        if has_grade_col:
            grade_val = ws.cell(r, COL_ROSTER_GRADE).value
            if grade_val is not None and str(grade_val).strip():
                grade = str(grade_val).strip()
        roster.append(
            {"name": str(name).strip(), "max_fte": max_fte_parsed, "grade": grade}
        )
        r += 1
    return roster
```

**Step 4: Run tests**

```
pytest tests/test_reader_new.py -v
pytest tests/ -q
```

Expected: new test PASS; existing 260 tests still PASS.

**Step 5: Commit**

```bash
git add budget_planner/reader.py tests/test_reader_new.py
git commit -m "fix(reader): use roster_ftes named range + tolerate internal blanks"
```

---

## Task 5: Fix `read_assignments_from_config`

**Files:**
- Modify: `budget_planner/reader.py` — function at line 425
- Modify: `budget_planner/reader.py` — `read_capacity_assignments` at line 515 (same pattern)
- Tests: `tests/test_reader_new.py`

**Step 1: Write failing test**

```python
def test_read_assignments_tolerates_blank_row_inside_named_range():
    """Blank inside assign_persons range must not drop later assignments."""
    import openpyxl
    from openpyxl.workbook.defined_name import DefinedName
    from budget_planner.reader import read_assignments_from_config

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Config"
    ws.cell(20, 1, "PERSON-WP ASSIGNMENTS")
    ws.cell(21, 1, "Person"); ws.cell(21, 2, "WP"); ws.cell(21, 3, "Total FTE")
    ws.cell(21, 4, "Start"); ws.cell(21, 5, "End")
    ws.cell(22, 1, "Alice"); ws.cell(22, 2, "WP1"); ws.cell(22, 3, 0.5)
    ws.cell(22, 4, 1); ws.cell(22, 5, 12)
    # Row 23 blank.
    ws.cell(24, 1, "Bob");   ws.cell(24, 2, "WP2"); ws.cell(24, 3, 0.3)
    ws.cell(24, 4, 6); ws.cell(24, 5, 18)
    wb.defined_names["assign_persons"] = DefinedName(
        name="assign_persons", attr_text="Config!$A$22:$A$24"
    )

    out = read_assignments_from_config(ws)
    persons = [a["person"] for a in out]
    assert persons == ["Alice", "Bob"], f"Got {persons}"
```

**Step 2: Run, see fail**

```
pytest tests/test_reader_new.py::test_read_assignments_tolerates_blank_row_inside_named_range -v
```

Expected: FAIL — "Bob" missing.

**Step 3: Implement**

Apply the same pattern as Task 4 to `read_assignments_from_config` (line 425) and `read_capacity_assignments` (line 515). Named range to check: `assign_persons`. The range gives first/last rows; iterate within, skip blanks, break only when outside a named range.

**Step 4: Run suite**

```
pytest tests/ -q
```

**Step 5: Commit**

```bash
git add budget_planner/reader.py tests/test_reader_new.py
git commit -m "fix(reader): use assign_persons named range for assignment reads"
```

---

## Task 6: Fix `read_externals_from_config`

**Files:**
- Modify: `budget_planner/reader.py` — function at line 469 and the paired `read_capacity_externals` (search for it near line 560 with `grep -n 'def read_capacity_externals\|ext_config_header' budget_planner/reader.py`)
- Tests: `tests/test_reader_new.py`

**Step 1: Write failing test** — same structure as T5, using `ext_persons` named range and an "EXTERNAL COMMITMENTS" section header.

**Step 2-4: As in T5.**

**Step 5: Commit**

```bash
git add budget_planner/reader.py tests/test_reader_new.py
git commit -m "fix(reader): use ext_persons named range for external reads"
```

---

## Task 7: T1 turns GREEN

**Files:**
- Verify: `tests/test_rebuild_shape_drift.py` (no edits expected)

**Step 1: Run T1's test**

```
pytest tests/test_rebuild_shape_drift.py -v
```

Expected: PASS.

**Step 2: Run the whole suite**

```
pytest tests/ -q
```

Expected: all green.

**Step 3: If T1 still fails**, diagnose whether it's a reader gap (unlikely — tasks 4–6 fixed all three) or a builder issue (the rebuilder re-emits formulas based on reader output — if the reader is now complete, builders should be correct). Fix forward before proceeding.

**Step 4: Commit** (no code change expected, but mark the milestone)

```bash
git commit --allow-empty -m "test: shape-drift reproduction now passes (T1 GREEN)"
```

---

## Task 8: Shape-parameterized rebuild tests (lean matrix)

**Files:**
- Create: `tests/test_rebuild_shapes.py`

**Scope (per design §2a):** default + one variant per table = 5 cases. Vary **one** dimension at a time.

**Step 1: Write the parameterized test**

```python
"""Parameterized rebuild round-trip tests varying one Config-table shape at a time."""
from pathlib import Path

import openpyxl
import pytest

from budget_planner.workbook import build_workbook, rebuild_workbook
from budget_planner.reader import (
    read_team_from_config,
    read_assignments_from_config,
    read_externals_from_config,
    read_config_sheet,
)


@pytest.mark.parametrize(
    "variant",
    [
        "default",
        "extra_wp",
        "extra_person",
        "extra_external",
        "extra_grade",
    ],
)
def test_rebuild_roundtrip_shape_variant(tmp_path: Path, variant: str, monkeypatch):
    """For each variant, generate → hand-edit one extra row mid-table → rebuild → assert preserved."""
    path = tmp_path / f"wb_{variant}.xlsx"
    build_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]

    # Insertion helpers keyed on section header — mirrors user's Excel workflow.
    def _insert_in_section(section_upper: str, values: list):
        for r in range(1, config_ws.max_row + 1):
            if str(config_ws.cell(r, 1).value or "").strip().upper() == section_upper:
                data_start = r + 2
                config_ws.insert_rows(data_start + 1, amount=1)
                for col, v in enumerate(values, start=1):
                    config_ws.cell(data_start + 1, col, v)
                return
        pytest.fail(f"Section {section_upper!r} not found")

    if variant == "extra_wp":
        _insert_in_section("WP BUDGET", ["WPX", "Extra WP", 1.0, 6500])
    elif variant == "extra_person":
        _insert_in_section("TEAM ROSTER", ["Midway Mary", 0.8, "Senior"])
    elif variant == "extra_external":
        _insert_in_section("EXTERNAL COMMITMENTS",
                           ["Alice", "Other Grant", 0.2, 1, 24, "Yes"])
    elif variant == "extra_grade":
        _insert_in_section("RATE BY GRADE", ["Principal", 9500])

    wb.save(path)
    wb.close()

    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]

    if variant == "extra_wp":
        _, wps, _ = read_config_sheet(config_ws)
        assert any(w["id"] == "WPX" for w in wps), [w["id"] for w in wps]
    elif variant == "extra_person":
        assert any(p["name"] == "Midway Mary" for p in read_team_from_config(config_ws))
    elif variant == "extra_external":
        assert any(e["project"] == "Other Grant"
                   for e in read_externals_from_config(config_ws))
    elif variant == "extra_grade":
        from budget_planner.reader import read_grade_rates
        assert "Principal" in read_grade_rates(config_ws)

    # Universal invariant: every assignment's person is in the roster.
    roster_names = {p["name"] for p in read_team_from_config(config_ws)}
    orphaned = {a["person"] for a in read_assignments_from_config(config_ws)} - roster_names
    assert not orphaned, f"Orphaned assignments after rebuild: {orphaned}"
    wb.close()
```

**Step 2: Run**

```
pytest tests/test_rebuild_shapes.py -v
```

Expected: 5/5 PASS (T4–T6 fixes covered the reader; builders re-emit from reader output).

**Step 3: Commit**

```bash
git add tests/test_rebuild_shapes.py
git commit -m "test: parameterized rebuild shape-variant round-trips"
```

---

## Task 9: META-divergence invariant helper

**Files:**
- Create: `tests/_shape_invariants.py` (shared helpers — underscore prefix keeps pytest from collecting)
- Modify: `tests/test_rebuild_shape_drift.py` and `tests/test_rebuild_shapes.py` to call the helper at end of each test

**Step 1: Write the helper**

```python
"""Shared test invariants for rebuilt workbooks."""
from budget_planner.reader import (
    read_team_from_config,
    read_assignments_from_config,
    read_externals_from_config,
)


def assert_meta_matches_live(wb):
    """META text must mention team/WP counts matching the live Config.

    _META is a documentation blob — this is a coarse sanity check that the
    blob was regenerated on the latest build and not carried over stale.
    """
    meta_ws = wb["_META"]
    meta_text = meta_ws["A1"].value or ""
    live_team = len(read_team_from_config(wb["Config"]))
    live_ext = len(read_externals_from_config(wb["Config"]))
    assert f"{live_team} people" in meta_text or f"{live_team}" in meta_text, (
        f"META does not reflect live team count ({live_team}). "
        "Was _META rewritten on last build?"
    )
    # Assignments: just ensure section exists; counts are not printed verbatim in meta.
    _ = live_ext  # reserved for future tightening
```

**Step 2: Add one call** at the end of the mid-insert test in T1 and inside the parameterized test in T8. Run the suite.

**Step 3: Commit**

```bash
git add tests/_shape_invariants.py tests/test_rebuild_shape_drift.py tests/test_rebuild_shapes.py
git commit -m "test: add assert_meta_matches_live invariant to shape tests"
```

---

## Task 10: Fingerprint schema + extractor

**Files:**
- Create: `scripts/shape_fingerprint.py` (schema + extractor)
- Create: `tests/test_shape_fingerprint.py`

**Purpose:** Produce a machine-readable, **value-free** structural snapshot of any workbook. This is the privacy boundary: the fingerprint is a JSON blob that by construction contains **only counts, row/column numbers, sheet names, and named-range names**. No cell values, no person names, no salaries, no WP names — ever.

**Fingerprint schema (JSON):**

```json
{
  "schema_version": 1,
  "sheet_names": ["Config", "Plan", "Capacity", "Actuals", "Summary", "Trends", "_META"],
  "named_ranges": {
    "roster_ftes":     {"sheet": "Config", "first_row": 62, "last_row": 74, "cols": ["B"]},
    "assign_persons":  {"sheet": "Config", "first_row": 26, "last_row": 58, "cols": ["A"]}
  },
  "config_sections": {
    "work_packages":        {"row_count": 9},
    "wp_schedule":          {"row_count": 12},
    "team_roster":          {"row_count": 13},
    "assignments":          {"row_count": 33},
    "external_commitments": {"row_count": 4},
    "reporting_periods":    {"row_count": 3},
    "grade_rates":          {"row_count": 4}
  },
  "cross_sheet_formula_targets": {
    "Capacity": {"config_refs": 412, "distinct_config_cells": 87},
    "Plan":     {"config_refs": 54,  "distinct_config_cells": 22}
  }
}
```

**Explicit forbidden fields:** no `cell_values`, no `sample_rows`, no strings or numbers that originated from any cell other than structural counts.

**Step 1: Write failing test — schema contract**

```python
"""Fingerprint must be value-free by construction."""
from pathlib import Path

import openpyxl

from scripts.shape_fingerprint import extract_fingerprint


def test_fingerprint_contains_only_structural_fields(tmp_path: Path):
    from budget_planner.workbook import build_workbook
    path = tmp_path / "wb.xlsx"
    build_workbook(str(path))

    fp = extract_fingerprint(str(path))

    # Top-level keys are the whitelist — anything else fails.
    assert set(fp.keys()) == {
        "schema_version", "sheet_names", "named_ranges",
        "config_sections", "cross_sheet_formula_targets",
    }
    # Named ranges: only structural fields.
    for nr in fp["named_ranges"].values():
        assert set(nr.keys()) <= {"sheet", "first_row", "last_row", "cols"}
    # Config sections: only counts.
    for sec in fp["config_sections"].values():
        assert set(sec.keys()) == {"row_count"}
```

**Step 2: Run, see fail**

```
pytest tests/test_shape_fingerprint.py::test_fingerprint_contains_only_structural_fields -v
```

Expected: FAIL (module does not exist).

**Step 3: Implement `scripts/shape_fingerprint.py`**

```python
"""Shape fingerprint — value-free structural snapshot of a workbook.

Privacy contract: the returned dict contains ONLY counts, row/column numbers,
sheet names, and named-range names. No cell values ever. If you extend this
module, do not add fields that store string or numeric data read from cells.
A canary test guards this contract — see tests/test_shape_fingerprint.py.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import openpyxl

from budget_planner.reader import (
    read_team_from_config,
    read_assignments_from_config,
    read_externals_from_config,
    read_reporting_periods,
    read_grade_rates,
    read_config_sheet,
)

SAFE_NAMED_RANGES = [
    "roster_ftes", "roster_grades",
    "assign_persons", "assign_wps", "assign_ftes",
    "ext_persons", "ext_ftes",
    "sched_wp_ids", "sched_starts", "sched_ends",
    "tbl_GradeRates", "project_start", "fte_rate",
    "eu_annual_ceiling_pm", "partner_ceiling", "current_month",
]

_REF_RE = re.compile(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?")


def _named_range_entry(wb, name):
    try:
        dn = wb.defined_names[name]
    except (KeyError, TypeError):
        return None
    try:
        dests = list(dn.destinations)
    except Exception:
        return None
    if not dests:
        return None
    sheet, ref = dests[0]
    m = _REF_RE.search(ref)
    if not m:
        return None
    col1, row1, col2, row2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    first_row = row1
    last_row = int(row2) if row2 else row1
    cols = sorted({col1, col2}) if col2 else [col1]
    return {"sheet": sheet, "first_row": first_row, "last_row": last_row, "cols": cols}


def extract_fingerprint(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=False)
    config_ws = wb["Config"]

    named_ranges = {}
    for nm in SAFE_NAMED_RANGES:
        entry = _named_range_entry(wb, nm)
        if entry is not None:
            named_ranges[nm] = entry

    _, wps, sched = read_config_sheet(config_ws)
    sections = {
        "work_packages":        {"row_count": len(wps)},
        "wp_schedule":          {"row_count": len(sched)},
        "team_roster":          {"row_count": len(read_team_from_config(config_ws))},
        "assignments":          {"row_count": len(read_assignments_from_config(config_ws))},
        "external_commitments": {"row_count": len(read_externals_from_config(config_ws))},
        "reporting_periods":    {"row_count": len(read_reporting_periods(config_ws))},
        "grade_rates":          {"row_count": len(read_grade_rates(config_ws))},
    }

    cross = {}
    for sheet_name in ["Capacity", "Plan", "Actuals", "Summary", "Trends"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        cells: set[tuple[str, int]] = set()
        total = 0
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("=") and "Config!" in v:
                    for m in _REF_RE.finditer(v):
                        total += 1
                        cells.add((m.group(1), int(m.group(2))))
        cross[sheet_name] = {"config_refs": total, "distinct_config_cells": len(cells)}

    return {
        "schema_version": 1,
        "sheet_names": list(wb.sheetnames),
        "named_ranges": named_ranges,
        "config_sections": sections,
        "cross_sheet_formula_targets": cross,
    }


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("usage: shape_fingerprint.py <workbook.xlsx> [-o fingerprint.json]",
              file=sys.stderr)
        return 2
    in_path = argv[1]
    out_path = None
    if len(argv) >= 4 and argv[2] == "-o":
        out_path = argv[3]
    fp = extract_fingerprint(in_path)
    js = json.dumps(fp, indent=2, sort_keys=True)
    if out_path:
        Path(out_path).write_text(js, encoding="utf-8")
        print(f"Fingerprint written to {out_path}. Safe to share — value-free by construction.")
    else:
        print(js)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
```

**Step 4: Run, see pass**

```
pytest tests/test_shape_fingerprint.py -v
```

**Step 5: Commit**

```bash
git add scripts/shape_fingerprint.py tests/test_shape_fingerprint.py
git commit -m "feat(scripts): shape fingerprint extractor (value-free by construction)"
```

---

## Task 11: Synthetic workbook generator

**Files:**
- Create: `scripts/build_from_shape.py`
- Extend: `tests/test_shape_fingerprint.py` with a generator round-trip test

**Purpose:** Given a fingerprint, generate a synthetic `.xlsx` matching the exact row counts and named-range spans — populated **only** with deterministic fake data (`Person_01`, `WP_A`, `1000.0`, fixed dates). This is the reproducer I debug against.

**Step 1: Write failing test**

```python
def test_generator_produces_workbook_matching_fingerprint(tmp_path):
    """Round trip: fingerprint → synthetic → re-fingerprint = original fingerprint."""
    from budget_planner.workbook import build_workbook
    from scripts.shape_fingerprint import extract_fingerprint
    from scripts.build_from_shape import build_from_fingerprint

    src = tmp_path / "src.xlsx"
    build_workbook(str(src))
    fp = extract_fingerprint(str(src))

    synth = tmp_path / "synth.xlsx"
    build_from_fingerprint(fp, str(synth))

    fp2 = extract_fingerprint(str(synth))
    assert fp2["config_sections"] == fp["config_sections"], (
        "Synthetic must preserve row counts per Config section"
    )
    assert set(fp2["named_ranges"].keys()) == set(fp["named_ranges"].keys())
```

**Step 2: Implement** — `build_from_fingerprint(fp: dict, out_path: str)`. Strategy:

- Call `budget_planner.config.default_seed()` (or equivalent) to get a baseline project.
- Override: `n_wps` = `fp["config_sections"]["work_packages"]["row_count"]`; `n_team`, `n_assignments`, `n_externals`, `n_grades`, `n_reporting_periods` likewise.
- Synthesise seed data in the shape: `work_packages = [{"id": f"WP_{chr(65+i)}", "name": f"Work package {i+1}", "fte": 1.0, "rate": 10000} for i in range(n_wps)]` etc. Use deterministic generators — no randomness.
- Call `build_workbook(out_path, seed=synthetic_seed)`.

The key invariant the test enforces: re-fingerprinting the generated workbook yields the same counts as the input fingerprint.

**Step 3: Run, commit**

```bash
pytest tests/test_shape_fingerprint.py -v
git add scripts/build_from_shape.py tests/test_shape_fingerprint.py
git commit -m "feat(scripts): generate synthetic workbook from a shape fingerprint"
```

---

## Task 12: One-shot wrapper `make_reproducer.py`

**Files:**
- Create: `scripts/make_reproducer.py`
- Create: `tests/test_make_reproducer.py`

**Purpose:** A single command that turns the user's real file into a shareable synthetic reproducer, with the fingerprint as an in-memory intermediate (not persisted unless requested).

**Step 1: Test first**

```python
def test_make_reproducer_end_to_end(tmp_path):
    from budget_planner.workbook import build_workbook
    from scripts.make_reproducer import make_reproducer
    from scripts.shape_fingerprint import extract_fingerprint

    real = tmp_path / "real.xlsx"
    build_workbook(str(real))
    synth = tmp_path / "synth.xlsx"

    make_reproducer(str(real), str(synth))

    assert synth.exists()
    assert extract_fingerprint(str(synth))["config_sections"] == \
           extract_fingerprint(str(real))["config_sections"]
```

**Step 2: Implement**

```python
"""Fully automated: real.xlsx -> synthetic.xlsx (safe to share, same shape).

Usage: python scripts/make_reproducer.py <real.xlsx> <synthetic.xlsx>
       python scripts/make_reproducer.py <real.xlsx> <synthetic.xlsx> --keep-fingerprint fp.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from scripts.shape_fingerprint import extract_fingerprint
from scripts.build_from_shape import build_from_fingerprint


def make_reproducer(real_path: str, out_path: str, keep_fingerprint: str | None = None) -> None:
    fp = extract_fingerprint(real_path)
    if keep_fingerprint:
        Path(keep_fingerprint).write_text(json.dumps(fp, indent=2, sort_keys=True))
    build_from_fingerprint(fp, out_path)


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(__doc__, file=sys.stderr)
        return 2
    real, out = argv[1], argv[2]
    keep = None
    if "--keep-fingerprint" in argv:
        keep = argv[argv.index("--keep-fingerprint") + 1]
    make_reproducer(real, out, keep_fingerprint=keep)
    print(f"Synthetic reproducer written to {out}. Safe to share.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
```

**Step 3: Commit**

```bash
git add scripts/make_reproducer.py tests/test_make_reproducer.py
git commit -m "feat(scripts): one-shot make_reproducer wrapper (real -> synthetic)"
```

---

## Task 13: Privacy canary test

**Files:**
- Create: `tests/test_privacy_canary.py`

**Purpose:** The hard privacy guarantee. Stuff every Config cell with a distinctive marker string; assert the marker never appears in the fingerprint JSON nor in the generated synthetic workbook's cell values. Any future code change that accidentally copies a cell value out will fail this test loudly.

**Step 1: Write the test**

```python
"""Privacy canary: confirm no real cell value ever leaves the real file.

Seeds a workbook with an unforgeable marker in every input cell, then runs the
reproducer pipeline end-to-end. The marker must not appear in:
  - the extracted fingerprint JSON
  - any cell value of the generated synthetic workbook

If either assertion fails, a cell value is leaking — investigate immediately.
"""
import json
from pathlib import Path

import openpyxl

from budget_planner.workbook import build_workbook
from scripts.make_reproducer import make_reproducer
from scripts.shape_fingerprint import extract_fingerprint


MARKER = "CONFIDENTIAL-CANARY-9f2a1c8e"


def _seed_with_marker(path: Path) -> None:
    build_workbook(str(path))
    wb = openpyxl.load_workbook(path)
    ws = wb["Config"]
    # Overwrite every text/number-looking cell in Config with the marker or
    # marker-tagged value. Formulas (start with "=") are left intact so the
    # workbook stays valid.
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and not v.startswith("="):
                cell.value = f"{MARKER}-{cell.coordinate}"
            elif isinstance(v, (int, float)):
                # Keep numeric; tag via a sibling cell would pollute structure.
                # Use a large distinctive magic number instead.
                cell.value = 987654321.0
    wb.save(path)


def test_marker_never_leaks_through_reproducer_pipeline(tmp_path: Path):
    real = tmp_path / "real_with_marker.xlsx"
    _seed_with_marker(real)

    # Fingerprint must be marker-free.
    fp = extract_fingerprint(str(real))
    fp_json = json.dumps(fp)
    assert MARKER not in fp_json, "Fingerprint leaked a cell string value"
    assert "987654321" not in fp_json, "Fingerprint leaked the magic numeric"

    # Synthetic must be marker-free.
    synth = tmp_path / "synth.xlsx"
    make_reproducer(str(real), str(synth))
    wb = openpyxl.load_workbook(synth)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    assert MARKER not in cell.value, (
                        f"Marker leaked into {ws.title}!{cell.coordinate}"
                    )
                elif isinstance(cell.value, (int, float)):
                    assert cell.value != 987654321.0, (
                        f"Magic number leaked into {ws.title}!{cell.coordinate}"
                    )
```

**Step 2: Run**

```
pytest tests/test_privacy_canary.py -v
```

Expected: PASS on first run. If FAIL, do not weaken the test — find the leak and stop it at the source.

**Step 3: Commit**

```bash
git add tests/test_privacy_canary.py
git commit -m "test: privacy canary — cell values never leak through reproducer"
```

---

## Task 14: Debugging runbook (user-facing)

**Files:**
- Create: `docs/debugging-your-real-file.md`

**Content outline:**

```markdown
# Debugging shape bugs in your confidential workbook

**When to use this:** You hit a bug in `rebuild` (or any other code path)
on your real `output/<your_confidential>.xlsx`. You want Claude's help
without Claude reading the file.

## One-shot workflow

```bash
python scripts/make_reproducer.py \
    output/<your_confidential>.xlsx \
    tmp/reproducer.xlsx
```

This produces `tmp/reproducer.xlsx` — a synthetic workbook with the same
table shapes as your real file, populated with fake data (`Person_01`,
`WP_A`, etc.). Safe to share.

## What to send Claude

1. A description of the bug ("After rebuild, Alice is missing from Config").
2. The path to the reproducer (`tmp/reproducer.xlsx`).
3. Optionally, the fingerprint JSON (`--keep-fingerprint fp.json`) for extra context.

Do NOT send the real workbook. Claude will not open it.

## Privacy guarantees

- `shape_fingerprint.py` reads only row counts, named-range spans, and
  formula coordinates. No cell values.
- `build_from_shape.py` synthesises all names and numbers deterministically
  from the fingerprint. It never receives cell values from your real file.
- The pipeline is guarded by `tests/test_privacy_canary.py` — a marker
  test that fails if any cell value ever leaks through.

## Re-running after a fix

Rerun `make_reproducer.py` whenever you edit your real file — the
reproducer is cheap to regenerate.
```

**Step 1: Write the doc.**

**Step 2: Commit**

```bash
git add docs/debugging-your-real-file.md
git commit -m "docs: runbook for shape-only debugging on confidential workbooks"
```

---

## Task 15: Full-suite sweep + trigger-table updates

**Step 1: Full test + type check**

```
pytest tests/ -q
pyright
```

Expected: all green. Fix anything that regressed.

**Step 2: Update CLAUDE.md trigger table**

Add two rows to the trigger table in `CLAUDE.md`:

```markdown
| Rebuild drops rows, stale Config refs | `docs/plans/2026-04-12-rebuild-shape-drift-design.md` | Readers use named ranges; tolerate internal blanks |
| Debugging against user's confidential workbook | `docs/debugging-your-real-file.md` | Run `make_reproducer.py`; never send the real file |
```

Also update `docs/internals.md` § Reader and Rebuild if that section talks about sentinel-based termination — it is now named-range-primary.

**Step 3: Commit docs**

```bash
git add CLAUDE.md docs/
git commit -m "docs: record shape-drift fix in trigger table + internals"
```

**Step 4: Stop here.** Do not push — user pushes from WSL (per global rule).

---

## Named-range audit results

*(T2 writes its findings here.)*
