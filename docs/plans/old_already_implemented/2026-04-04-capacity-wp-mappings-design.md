# Capacity Person-WP Mappings & Auto-Distribution

**Date:** 2026-04-04
**Status:** Validated design, ready for implementation

## Goal

Reduce Capacity sheet clutter by showing only assigned WPs per person (instead of all 9), and auto-fill monthly FTE values using distribution curves based on WP type.

## Design Decisions

- Monthly cells are **static values written by Python** (not Excel formulas)
- Rebuild **always recalculates** from config block — manual month overrides are lost
- Users can override individual month cells post-generation for fine-tuning
- Distribution type lives on **Config sheet** (per WP), not repeated in Capacity config block

## Changes

### 1. Config Sheet — WP Type Column

Add a `Type` column (column F) to `tbl_WorkPackages`:

| WP ID | WP Name | FTE Budget | Start | End | Type |
|-------|---------|------------|-------|-----|------|
| WP1 | Project Management | 12.0 | 1 | 14 | flat |
| WP2 | Requirements and concept | 8.0 | 1 | 6 | front-loaded |
| WP3 | Research and analysis | 10.0 | 1 | 7 | front-loaded |
| WP4 | Architecture design | 15.0 | 2 | 11 | bell |
| WP5 | Component design | 12.0 | 4 | 13 | bell |
| WP6 | Subsystem design | 12.0 | 4 | 13 | bell |
| WP7 | Software development | 10.0 | 6 | 14 | bell |
| WP8 | Prototyping | 8.0 | 10 | 14 | back-loaded |
| WP9 | Validation and demo | 6.0 | 12 | 14 | back-loaded |

- Blue input cell — user can change type in Excel
- Valid values: `flat`, `front-loaded`, `bell`, `back-loaded`
- Reader picks it up on rebuild

### 2. Capacity Config Blocks

Two config blocks at the top of the Capacity sheet, both in a **collapsed row group** by default:

**Block 1: PERSON-WP ASSIGNMENTS** (new)

| Person | WP | Total FTE | Start | End |
|--------|-----|-----------|-------|-----|
| Alice | WP1 | 4.5 | 1 | 14 |
| Alice | WP2 | 4.0 | 1 | 6 |
| Bob | WP4 | 7.5 | 2 | 11 |

- Blue input cells
- Start/End default to WP active period, can be narrowed
- Total FTE = person's allocation to that WP
- Seed: `DEMO_CAPACITY_ASSIGNMENTS` for `--demo`; empty for plain `init`

**Block 2: EXTERNAL COMMITMENTS** (existing, unchanged layout)

### 3. Capacity Per-Person Grid

- Only assigned WPs shown per person (from config block), not all WPs
- External project rows for persons with external commitments
- Monthly values auto-filled using distribution curves (see below)
- Inactive months: grey and empty
- Validation columns unchanged: Sum, Budget (ref config block), Variance

### 4. Distribution Curves

All distributions sum exactly to the person's Total FTE allocation.

| Type | Shape | Calculation |
|------|-------|-------------|
| `flat` | Equal every month | `total / months` |
| `front-loaded` | Linearly decreasing | Month 1 highest, last month lowest |
| `bell` | Triangular | Linear ramp to midpoint, symmetric ramp down |
| `back-loaded` | Linearly increasing | Month 1 lowest, last month highest |

External project rows always use `flat` distribution.

### 5. Reader & Rebuild

**Reader additions:**
- `read_config_sheet()` gains `type` field per WP
- `read_capacity_config()` reads both Person-WP Assignments and External Commitments blocks
- Block headers used as landmarks

**Rebuild flow:**
1. Read Config sheet (WPs now include `type`)
2. Read Capacity config blocks (person-WP + external)
3. Rebuild Config sheet (with Type column)
4. Rebuild Capacity sheet:
   - Write Person-WP Assignments config block
   - Write External Commitments config block
   - Collapse both blocks
   - Per person: rows only for assigned WPs + external projects
   - Calculate distribution per WP type, write static values
   - Write validation formulas + TOTAL/REMAINING rows

**Preserved on rebuild:** Config block values (Person, WP, Total FTE, Start, End), WP types
**Recalculated on rebuild:** All monthly cell values

### 6. Plan Sheet

No changes needed — Plan cross-sheet refs read cell values regardless of whether they're formulas or static numbers.

## Dependency Graph

```
Step 1: config.py — add type field to WORK_PACKAGES, update DEMO seeds
Step 2: config_sheet.py — add Type column to tbl_WorkPackages
Step 3: reader.py — read Type from Config, read Person-WP config block
  (depends on: Step 2)
Step 4: capacity.py — Person-WP config block + collapsed grouping + distribution logic + filtered per-person grids
  (depends on: Steps 1, 2, 3)
Step 5: meta.py — update _META sheet documentation
  (depends on: Steps 2, 4)
Step 6: tests — update all affected tests
  (depends on: Steps 1-5)

Parallelizable: Steps 1+2 can run together. Step 3 after 2. Step 4 after 1+2+3.
```
