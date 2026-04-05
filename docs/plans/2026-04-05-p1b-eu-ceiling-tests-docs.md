# P1b: 215-Day EU Ceiling Check — Tests & Docs (Session 2 of 2)

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Prerequisite:** Session 1 (`docs/plans/2026-04-05-p1a-eu-ceiling-implementation.md`) must be merged first.

**Goal:** Add rebuild safety-net tests for manual row insertion, update _META documentation, and run final integration verification.

**Tech Stack:** Python, openpyxl, pytest

---

## Task 1: Add missing "insert new row" rebuild tests

**Why:** No tests exist for the scenario where a user manually adds a new row
to PERSON-WP ASSIGNMENTS, EXTERNAL COMMITMENTS, or TEAM ROSTER in Excel and
then runs rebuild. Only modify/rename scenarios are covered today. These tests
validate existing behavior — they should pass without code changes.

**Files:**
- Test: `tests/test_rebuild_e2e.py`

**Step 1: Write tests for all three insert scenarios**

Add to `tests/test_rebuild_e2e.py`:
```python
class TestRebuildWithInsertedRows:
    """User adds new rows in Excel config blocks, then rebuilds."""

    def test_added_assignment_row_creates_wp_row_after_rebuild(self, assigned_init_path):
        """New assignment row added in Capacity config block → person gets new WP row."""
        from budget_planner.reader import read_capacity_assignments, read_capacity_sheet
        from budget_planner.config import TEAM, WORK_PACKAGES

        wb = openpyxl.load_workbook(assigned_init_path)
        cap_ws = wb["Capacity"]
        original = read_capacity_assignments(cap_ws)

        # Find the last assignment row and add a new one below it
        last_asgn_row = None
        for r in range(1, cap_ws.max_row + 1):
            person = cap_ws.cell(r, 1).value
            wp = cap_ws.cell(r, 2).value
            if person and wp and any(a["person"] == str(person) and a["wp"] == str(wp)
                                     for a in original):
                last_asgn_row = r

        assert last_asgn_row is not None

        # Pick a person and WP that aren't already assigned together
        target_person = TEAM[0]["name"]
        assigned_wps = {a["wp"] for a in original if a["person"] == target_person}
        new_wp = next((wp["id"] for wp in WORK_PACKAGES
                       if wp["id"] not in assigned_wps), None)
        if new_wp is None:
            pytest.skip("All WPs already assigned to first person")

        new_row = last_asgn_row + 1
        _set(cap_ws, new_row, 1, target_person)
        _set(cap_ws, new_row, 2, new_wp)
        _set(cap_ws, new_row, 3, 2.0)   # total FTE
        _set(cap_ws, new_row, 4, 1)     # start month
        _set(cap_ws, new_row, 5, 12)    # end month
        wb.save(assigned_init_path)
        wb.close()

        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        _, wps = read_config_sheet(wb["Config"])
        cap_data = read_capacity_sheet(wb["Capacity"], wps)
        person_wps = cap_data["person_groups"][target_person]["wp_values"]
        assert new_wp in person_wps, (
            f"{target_person} should have WP row for {new_wp} after rebuild"
        )
        wb.close()

    def test_added_external_commitment_row_appears_after_rebuild(self, assigned_init_path):
        """New external commitment added in Capacity config block → person gets external row."""
        wb = openpyxl.load_workbook(assigned_init_path)
        cap_ws = wb["Capacity"]

        # Find EXTERNAL COMMITMENTS config block end
        ext_header = None
        for r in range(1, cap_ws.max_row + 1):
            val = cap_ws.cell(r, 1).value
            if val and "EXTERNAL COMMITMENTS" in str(val).upper():
                ext_header = r
                break
        assert ext_header is not None

        # Find last data row in the block
        r = ext_header + 2  # skip section header + column headers
        while r <= cap_ws.max_row and cap_ws.cell(r, 1).value:
            r += 1
        insert_row = r  # first empty row after existing data

        new_person = TEAM[-1]["name"]  # last team member
        new_project = "NewExtProject"
        _set(cap_ws, insert_row, 1, new_person)
        _set(cap_ws, insert_row, 2, new_project)
        _set(cap_ws, insert_row, 3, 2.0)
        _set(cap_ws, insert_row, 4, 1)
        _set(cap_ws, insert_row, 5, 12)
        wb.save(assigned_init_path)
        wb.close()

        rebuild_workbook(input_path=assigned_init_path)

        wb = openpyxl.load_workbook(assigned_init_path)
        _, wps = read_config_sheet(wb["Config"])
        cap_data = read_capacity_sheet(wb["Capacity"], wps)
        person_ext = cap_data["person_groups"][new_person]["external_values"]
        assert new_project in person_ext, (
            f"{new_person} should have external row for {new_project} after rebuild"
        )
        wb.close()

    def test_added_team_member_gets_capacity_group_after_rebuild(self, init_path):
        """New person added to Config TEAM ROSTER → gets Capacity person group."""
        wb = openpyxl.load_workbook(init_path)
        config_ws = wb["Config"]

        # Find last roster data row
        roster_section = None
        for r in range(1, config_ws.max_row + 1):
            if config_ws.cell(r, 1).value and "TEAM ROSTER" in str(config_ws.cell(r, 1).value).upper():
                roster_section = r
                break
        assert roster_section is not None

        # Scan to end of roster data
        r = roster_section + 2  # skip section header + column headers
        while r <= config_ws.max_row and config_ws.cell(r, 1).value:
            r += 1
        insert_row = r

        _set(config_ws, insert_row, 1, "Eve")
        _set(config_ws, insert_row, 2, 0.5)
        wb.save(init_path)
        wb.close()

        rebuild_workbook(input_path=init_path)

        wb = openpyxl.load_workbook(init_path)
        _, wps = read_config_sheet(wb["Config"])
        cap_data = read_capacity_sheet(wb["Capacity"], wps)
        assert "Eve" in cap_data["person_groups"], (
            f"'Eve' should have a Capacity person group after rebuild"
        )
        # Verify Eve appears in rebuilt roster
        roster = read_team_from_config(wb["Config"])
        names = {p["name"] for p in roster}
        assert "Eve" in names
        wb.close()
```

**Step 2: Run tests — expect them to PASS (testing existing behavior)**

Run: `pytest tests/test_rebuild_e2e.py::TestRebuildWithInsertedRows -v`
Expected: All 3 PASS. If any fail, that's a pre-existing bug to fix before proceeding.

**Step 3: Commit**

```bash
git add tests/test_rebuild_e2e.py
git commit -m "test(rebuild): add insert-new-row scenarios for assignments, externals, roster"
```

---

## Task 2: Update _META sheet documentation

**Files:**
- Modify: `budget_planner/meta.py` (Capacity section, named ranges section)

**Step 1: Find the Capacity documentation section in meta.py**

Read `budget_planner/meta.py` and find the section documenting the Capacity sheet structure.

**Step 2: Add EU TOTAL documentation**

Add to the Capacity sheet section:
- Document the EU TOTAL row: purpose, formula structure, placement after REMAINING
- Document the EU ceiling CF rules: per-calendar-year, red >ceiling, amber >ceiling-1
- Document the `is_eu` column in the external commitments config block

Add to the named ranges section:
- `eu_annual_ceiling_pm` → Config!$B$5 (default 12.0, ≈215 working days)

**Step 3: Run tests**

Run: `pytest tests/ -q`
Expected: All pass (meta changes are documentation only)

**Step 4: Commit**

```bash
git add budget_planner/meta.py
git commit -m "docs(meta): document EU TOTAL row, ceiling param, is_eu flag"
```

---

## Task 3: Final integration verification

**Step 1: Generate demo workbook and inspect**

```bash
python -m budget_planner init --demo
python scripts/recalc.py output/budget_plan.xlsx
```

Open in Excel/LibreOffice and verify:
- Config sheet row 5: "EU Annual Ceiling (PM)" = 12.00
- Capacity external config block: "EU Grant" column with Yes/No values
- Each person has EU TOTAL row (light blue) after REMAINING
- Bob's EU TOTAL shows programme WPs + HorizonEU-Alpha combined monthly FTE
- Calendar-year CF highlighting appears when EU total approaches ceiling

**Step 2: Run full test suite**

```bash
pytest tests/ -v
```

Expected: All pass

**Step 3: Run rebuild round-trip**

```bash
python -m budget_planner rebuild
pytest tests/ -q
```

Expected: All pass, no data loss warnings

**Step 4: Final commit (if any fixups needed)**

```bash
git add -A
git commit -m "feat(P1): 215-day EU ceiling check complete"
```

---

## Summary of changes per file

| File | Change |
|------|--------|
| `meta.py` | Document EU TOTAL row, ceiling param, is_eu flag, CF rules |
| `test_rebuild_e2e.py` | Insert-new-row scenarios: assignment, external, roster |
