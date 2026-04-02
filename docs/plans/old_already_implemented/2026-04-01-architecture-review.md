# Architecture Review: Capacity Planning Design Spec

Reviewed: 2026-04-01
Spec: `docs/plans/2026-04-01-capacity-planning-design.md`
Codebase: `budget_planner/` (config, styles, plan, summary, meta, reader, workbook)

## 1. Internal Consistency

### PASS

- Single planning surface (Capacity) → Plan formulas → Summary formulas: data flow is unidirectional, no circular dependencies
- Font color convention (blue=input, black=formula, green=xref) consistently applied across all sheet designs
- Round-trip pattern (reader → preserved_data → builder) is consistent with existing architecture
- Snapshot-based scenario planning avoids the toggle complexity that would have created bidirectional formula paths

### ISSUES

**I1: Duplicate section numbering (two "Section 7" headers)**
The spec has two sections numbered 7: "Scenario Planning Workflow" and "Generation Modes". Minor doc issue.

**I2: External commitment seeding references "Total FTE ÷ active months" but no config columns defined**
The Capacity sheet description shows config columns (Person, Project, Total FTE, Start Month, End Month) for external projects, but these appear in the Capacity sheet layout description (Section 2). However, the Capacity sheet layout example only shows 2 left columns (A=ID, B=Name) for all rows. The external project config columns (Total FTE, Start/End) need a home — either:
- Additional columns on the Capacity sheet (but then programme WP rows have unused columns), or
- These are generation-time inputs from `config.py` only, not stored in Excel

This needs clarification. If the user wants to re-seed an external project (e.g., change total FTE and redistribute), those values need to be somewhere in Excel.

**I3: "WP share × WP planned FTE / active months" seeding formula is circular**
The spec says Capacity cells are seeded from `WP share × WP planned FTE / active months`. But after the redesign, the Plan sheet's Planned FTE is a SUM formula reading FROM Capacity. On first generation this is fine (seed values exist before formulas), but the language is confusing — it should reference config table's Total FTE, not "planned FTE."

## 2. Feasibility Against Current Architecture

### Builder Pattern: FEASIBLE

The current pattern is clean: each sheet has a `build_X_sheet()` function that receives a worksheet, workbook, and layout/preserved data. New sheets (Team, Capacity) fit naturally:

```
build_team_sheet(ws, wb, layout)         → returns team_layout
build_capacity_sheet(ws, wb, layout, team_layout, preserved_data) → returns capacity_layout
```

The orchestrator (`workbook.py`) calls builders in sequence and threads layout dicts through. This scales to 5 sheets without architectural change.

### Styles: FEASIBLE

All needed styles already exist in `styles.py`: BLUE_INPUT, BLACK_FORMULA, GREEN_XREF, GREY_FILL, SECTION_FILL, etc. New conditional formatting (amber fill for >0.8) needs one new constant: `AMBER_FILL = PatternFill("solid", fgColor="FFD966")`.

### Reader Pattern: NEEDS SIGNIFICANT EXTENSION

**Current reader** (`reader.py`): scans Plan sheet only, uses heuristic markers ("WP ID", "PLANNED", "ACTUAL") to find blocks, reads cell values by position. Returns a flat dict keyed by WP index.

**New reader needs**:
1. Read Plan config table → extract WP definitions (ID, name, FTE, start, end)
2. Read Team sheet → extract roster + WP share matrix
3. Read Capacity sheet → discover person groups (variable row counts per person), classify rows as programme WP or external, extract monthly values
4. Handle variable-length structures (different people have different numbers of external projects)

The Capacity reader is the hardest part. Current reader assumes fixed WP count → fixed row offsets. The Capacity sheet has **variable geometry**: each person block has N_wp + N_external + 2 (TOTAL, REMAINING) rows, and N_external varies per person.

**Approach**: scan for person header rows (section fill + bold font, or match against Team roster names), then within each person block, scan until TOTAL row. This is more complex than the current reader but follows the same heuristic pattern.

### Formula Complexity: MANAGEABLE BUT FRAGILE

**Plan Planned FTE formulas** become:
```
=SUM(Capacity!C15, Capacity!C25, Capacity!C35, Capacity!C42)
```
(summing each person's WP1 row for month M1)

These are **non-contiguous cell references** — not a simple `SUM(range)` because each person's WP1 row is at a different offset within the Capacity sheet. This means:

- Formulas must be constructed at build time with exact row numbers from the capacity layout
- Adding/removing a person or external project row shifts ALL subsequent row numbers
- `--update` must rebuild these formulas every time (can't preserve them)

This is feasible but is the single most fragile part of the design. The builder must always regenerate Plan formulas from the current Capacity layout, never try to preserve them.

## 3. Excel-First Risks

### R1: WP Addition/Removal Cascades (MEDIUM risk)

Adding a WP in Excel means:
1. New row in Plan config table
2. New row in each person's group on Capacity sheet (if they work on that WP)
3. New column in Team share matrix
4. Updated SUM formulas on Plan sheet
5. Updated Summary sheet

The `--update` tool must detect the config table change and propagate. This is the equivalent of a database schema migration — doable but needs careful ordering. The current reader handles fixed WP counts; the new one must diff against previous state.

**Mitigation**: On `--update`, always regenerate formula sheets (Plan planned, Summary) from scratch. Only preserve input cells (Capacity monthly values, Actuals, Team roster).

### R2: Row Count Sensitivity (HIGH risk)

The Capacity sheet has variable row counts per person. Every cross-sheet formula from Plan to Capacity depends on exact row numbers. If the user inserts/deletes a row manually in Excel (e.g., adds an external project), the Plan formulas won't update (openpyxl doesn't adjust formulas like Excel does).

**Mitigation**: `--update` must always rebuild Plan formulas from discovered Capacity layout. Document in _META: "After manually adding rows in Capacity, run `--update` to rebuild formulas."

### R3: Dual Source of Truth During Transition (LOW risk)

During `--update`, the tool reads WP definitions from Excel but still imports `MONTH_COL_START` and `PROJECT_MONTHS` from `config.py`. If these diverge (e.g., user changed project duration in their head but not in config.py), the tool would generate wrong column counts.

**Mitigation**: Consider reading `PROJECT_MONTHS` from the Excel file too (count month columns in existing Plan sheet), or validate that config.py and Excel agree.

## 4. Cross-Sheet Formula Map

```
Team!Roster
  ↓ (name refs)
Capacity!Person headers ←→ Capacity!Monthly grid (input)
  ↓ (SUM per WP across people)         ↓ (TOTAL/REMAINING per person)
Plan!Planned FTE (formulas)           Capacity!Load overview
  ↓                                     ↑
Plan!Config table (input)             Team!Max FTE
  ↓
Summary!Annual aggregation (formulas)
```

Cross-sheet refs: 
- Capacity → Team (person names, Max FTE): simple cell refs
- Capacity → Plan config (WP ID, WP Name): simple cell refs  
- Plan Planned FTE → Capacity (non-contiguous SUM): complex, layout-dependent
- Summary → Plan (unchanged from today): existing pattern

The Capacity → Plan config refs are the same pattern as the existing WP rows referencing config. The new complexity is Plan → Capacity (reverse direction, non-contiguous).

## 5. Breaking Changes

### Test Suite Impact — Detailed Mapping

All 58 existing tests will break or need modification. Each finding from this review maps to specific test cases in the design spec's Test Strategy (Steps 0–8):

| Test area | Impact | Covers finding |
|-----------|--------|----------------|
| `test_workbook.py::test_sheets_exist` | Add `"Team"`, `"Capacity"` to expected list | — |
| `test_workbook.py::test_planned_fte_has_formulas` | Still passes — SUM formulas change shape but assertion holds | — |
| `test_workbook.py::test_planned_fte_variance` | **Move to e2e** — can't read cross-sheet SUM results without recalc | F1 |
| `test_workbook.py::test_actual_fte_zeros` | Unchanged | — |
| `test_workbook.py::test_inactive_months_empty` | Extend to Capacity sheet programme WP rows (Step 3) | — |
| `test_roundtrip.py::test_discover_layout` | Extend for Team + Capacity layouts (Step 5) | F2 |
| `test_roundtrip.py::test_preserves_actuals` | Unchanged — regression guard for migration | B2 |
| `test_roundtrip.py::test_preserves_planned_edits` | **Delete** → replaced by `test_roundtrip_capacity_edits_preserved` (Step 7) | B1 |
| `test_roundtrip.py::test_fresh_generation_has_even_distribution` | Move to Capacity — distribution is per-person now (Step 3) | — |
| `test_recalc.py::test_formula_count_positive` | Will increase significantly — add `test_formula_count_increased` (Step 8) | — |

**New tests directly addressing review findings:**

| Finding | Risk | Test (from design spec) | Step |
|---------|------|------------------------|------|
| R1: WP add/remove cascades | Medium | `test_roundtrip_team_roster_edits_preserved` — add person, verify cascade | 7 |
| R2: Row count sensitivity | **High** | `test_planned_fte_formula_refs_match_capacity_layout` — verify row refs after layout change | 4 |
| R2: Row count sensitivity | **High** | `test_roundtrip_capacity_edit_reflected_in_plan` — formula rebuilt after edit | 7 |
| R3: config.py vs Excel divergence | Low | `test_read_plan_config_extracts_wp_definitions` — reads from Excel, not config.py | 5 |
| F1: Non-contiguous Plan SUMs fragile | **High** | `test_capacity_to_plan_formula_linkage` — parse formula, match against layout | 7 |
| F1: Non-contiguous Plan SUMs fragile | **High** | `test_demo_mode_plan_planned_equals_capacity_sums` — recalc verifies arithmetic | 8 |
| F2: Variable-geometry Capacity reader | Medium | `test_read_capacity_handles_variable_geometry` — 0 vs 2 external projects | 5 |
| B1: Planned edits test breaks | Medium | `test_roundtrip_capacity_edits_preserved` replaces old test | 7 |
| B2: Old workbook incompatibility | Medium | `test_v1_update_creates_team_and_capacity_sheets` | 7 |
| B2: Old workbook incompatibility | Medium | `test_v1_update_produces_valid_v2` — structural checks pass | 7 |

**Critical path for test confidence:** Steps 3→4→5→7 cover the highest-risk findings (R2, F1, F2). These must pass before any implementation is considered complete.

### Migration Path

Existing workbooks (generated before this change) cannot be `--update`d by the new tool — they have no Team or Capacity sheet, and Plan has value cells where formulas are expected.

**Decision (AD2):** Auto-detect v1 format on `--update`. No separate `--migrate` flag.

**Test coverage for migration (Step 7):**
- `test_v1_update_creates_team_and_capacity_sheets` — new sheets appear
- `test_v1_update_assigns_to_unassigned_person` — placeholder person holds all FTE
- `test_v1_update_preserves_actuals` — actuals survive migration
- `test_v1_update_produces_valid_v2` — output passes all v2 structural checks

## 6. Summary of Findings

| # | Finding | Severity | Action |
|---|---------|----------|--------|
| I1 | Duplicate section 7 numbering | Low | Fix in spec |
| I2 | External project config columns undefined in layout | Medium | Clarify where Total FTE / Start / End live in Excel |
| I3 | Seeding formula references "planned FTE" (circular) | Low | Fix wording to "config Total FTE" |
| R1 | WP add/remove cascades across 4 sheets | Medium | Always regenerate formula sheets on --update |
| R2 | Row count sensitivity breaks Plan formulas | High | Always rebuild Plan formulas from Capacity layout on --update |
| R3 | config.py vs Excel divergence for structural constants | Low | Validate or read PROJECT_MONTHS from Excel |
| B1 | test_update_preserves_manual_planned_edits breaks | Medium | Rewrite as Capacity round-trip test |
| B2 | Old workbooks incompatible with --update | Medium | Accept as v2 or add --migrate |
| F1 | Non-contiguous Plan SUM formulas are fragile | High | Document rebuild requirement; never preserve Plan formulas |
| F2 | Variable-geometry Capacity reader is complex | Medium | Use person header detection heuristic |

## Decisions Made

| # | Finding | Decision |
|---|---------|----------|
| I2 | External project config columns | Config block at top of Capacity sheet (mirrors Plan pattern) |
| I3 | Seeding formula wording | Fixed: "config Total FTE" not "planned FTE" |
| R2/F1 | Plan formula fragility | Always rebuild Plan + Summary formulas on every --update |
| B1 | Planned edits test breaks | Both: Capacity round-trip test + backward compat test |
| B2 | Old workbook incompatibility | Auto-detect v1 format on --update, assign to placeholder person |
