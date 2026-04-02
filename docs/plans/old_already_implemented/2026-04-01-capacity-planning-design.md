# Capacity Planning & Draft Mode — Design Spec

## Goals

1. **Cross-project per-person visibility**: See each team member's total FTE load across all projects, per month, in one place. Spot overcommitment early.
2. **Safe replanning**: Explore FTE reallocations without destroying the current valid plan. Promote draft to baseline when satisfied.
3. **Excel-first data**: All project, team, and commitment data lives in the workbook. `config.py` is only the initial seed.
4. **Two generation modes**: Skeleton (production) and demo (exploration).

## Non-Goals

- Automated Redmine integration
- Multi-partner / consortium-wide view
- Per-deliverable tracking (WP-level granularity is sufficient)
- Actuals import from timesheets

## Key Design Decision: Single Planning Surface

All per-person monthly FTE allocation — both programme WP contributions and external project commitments — is entered in **one place**: the Capacity sheet. The Plan sheet's Planned FTE block becomes **formula-driven** (sums across people from Capacity), not directly editable.

This eliminates duplicate monthly grids and gives per-person cross-project visibility as a natural consequence of the data entry workflow.

## Sheet Overview

| Sheet | Purpose | New? |
|-------|---------|------|
| `_META` | AI-agent-readable documentation | Existing (updated) |
| `Plan` | Config table + Planned FTE (formulas) + Actual FTE (input) | Existing (modified) |
| `Team` | Person roster + WP share seeds | **New** |
| `Capacity` | **The** planning surface — per-person monthly allocation | **New** |
| `Summary` | Annual budget aggregation | Existing (unchanged) |

Scenario planning is handled via file-level snapshots, not a dedicated sheet.

### Monthly Grids Inventory

| Grid | Sheet | Editable? | Purpose |
|------|-------|-----------|---------|
| Per-person allocation | Capacity | **Yes** | Primary planning input — programme WPs + external projects |
| WP Planned FTE | Plan | No (formulas) | Compliance view — sums people per WP |
| WP Actual FTE | Plan | **Yes** | Actuals entry (unchanged from today) |

## Detailed Design

### 1. Team Sheet

Tab color: Orange.

**Block 1 — Team Roster**

| Column | Content | Type |
|--------|---------|------|
| A | Person name | Blue input |
| B | Max FTE | Blue input (default 1.0) |

Max FTE handles part-time contracts. Used by Capacity sheet for remaining capacity calculation.

**Block 2 — WP Share Seeds** (below roster, separated by gap row)

|        | WP1  | WP2  | WP3  | ... |
|--------|------|------|------|-----|
| Alice  | 0.25 | 0.50 | 0.25 | ... |
| Bob    | 0.25 |      | 0.50 | ... |

- Row headers = person names (formula refs to roster)
- Column headers = WP IDs (formula refs to Plan config table)
- Cell values = blue input, fraction of WP effort assigned to this person
- **Validation row**: SUM per column. Yellow fill if ≠ 1.0 (WP effort not fully assigned)
- Empty cells treated as 0
- These shares are used **only** to seed the Capacity sheet's monthly grid on first generation. After that, the Capacity sheet is the source of truth.

### 2. Capacity Sheet

Tab color: Purple. **This is the primary planning surface.**

**Config block — External Commitments** (top of sheet, above person groups)

| Person | Project | Total FTE | Start Month | End Month |
|--------|---------|-----------|-------------|-----------|
| Alice  | ProjectX | 3.0      | 1           | 11        |
| Alice  | ProjectZ | 2.0      | 13          | 24        |
| Bob    | ProjectX | 6.0      | 1           | 60        |

- Blue input cells. Used to seed external project rows in the monthly grid below.
- Mirrors the Plan sheet's config table pattern (config above grid).
- On `--update`, these values are preserved. The monthly grid is the source of truth for actual allocation; this table is reference/seed data.

**Per-Person Allocation** (grouped by person, below config block)

For each person, rows are grouped together:

```
ALICE                                    M1     M2     M3    ...   M60
  WP1 - Project Management             0.05   0.05   0.05  ...   0.05
  WP2 - Requirements & Design          0.10   0.10   0.15  ...   0.00
  WP3 - Core Development               0.05   0.05   0.05  ...   0.00
  ProjectX                              0.30   0.30   0.50  ...   0.00
  ProjectZ                              0.00   0.00   0.20  ...   0.00
  TOTAL                                 0.50   0.50   0.95  ...   0.05
  REMAINING                             0.50   0.50   0.05  ...   0.95

BOB
  WP1 - Project Management             0.05   0.05   0.05  ...   0.05
  ProjectX                              0.50   0.50   0.80  ...   0.00
  TOTAL                                 0.55   0.55   0.85  ...   0.05
  REMAINING                             0.45   0.45   0.15  ...   0.95
```

**Row types:**

| Row type | Content | Type |
|----------|---------|------|
| Person header | Person name (bold, section fill) | Formula ref to Team roster |
| Programme WP row | WP ID + name, monthly FTE | Blue input |
| External project row | Project name, monthly FTE | Blue input |
| TOTAL row | SUM of all rows for this person | Formula (black) |
| REMAINING row | Max FTE (from Team) - TOTAL | Formula (black) |

**Config columns for each row** (left side):

| Column | programme WP rows | External project rows |
|--------|------------|----------------------|
| A | WP ID (formula ref to Plan config) | Project name (blue input) |
| B | WP Name (formula ref to Plan config) | — |

**Row type detection** (for `--update` reader): column A is matched against the config table's WP IDs. Match = programme WP row. No match = external project row. No extra marker column needed.

**Monthly grid** (M1..M60):

- All blue input cells
- Programme WP rows: first generation seeded from `WP share × config Total FTE ÷ active months` (even distribution of this person's portion)
- External project rows: first generation seeded from `Total FTE ÷ active months` (from config block above)
- Inactive months (outside WP/project active period): greyed out, empty
- User edits freely to shape the actual allocation curve
- `--update` preserves all manual edits

**Validation columns** (rightmost, per row):

| Column | Content | Type |
|--------|---------|------|
| Sum | SUM of monthly values | Formula |
| Budget | Person's share of WP budget FTE (programme rows) or Total FTE (external rows) | Formula / blue input |
| Variance | Budget - Sum | Formula, yellow if ≠ 0 |

**Conditional formatting:**

- TOTAL cell red fill if > Max FTE (overcommitted)
- TOTAL cell amber fill if > 0.8 × Max FTE (approaching limit)
- REMAINING cell red fill if < 0

**Double-confirmation on regeneration:**

Full regeneration without `--update`: CLI prompts twice before overwriting:
1. "Capacity data exists in the current file. Overwrite with defaults? [y/N]"
2. "Are you sure? This cannot be undone. [y/N]"

Defaults to No on both prompts.

### 3. Plan Sheet (modified)

The Plan sheet's **Planned FTE block changes from input to formulas**:

- Each WP × month cell becomes a SUM formula across all people's corresponding Capacity rows
- e.g., `Plan!WP1_M3 = SUM(Capacity!Alice_WP1_M3, Capacity!Bob_WP1_M3, ...)`
- Font color: black (formula), not blue (input)
- Validation columns (Sum, Budget FTE, Variance) remain as today

The **Actual FTE block is unchanged** — direct blue input cells for logged time.

The **Config table is unchanged** — WP definitions (ID, name, total FTE, start/end month).

### 4. Scenario Planning via Snapshots

Instead of a dedicated Draft sheet, scenario planning uses file-level snapshots (leveraging the existing timestamped backup mechanism):

**`--snapshot`** CLI flag: copies the current workbook to `budget_plan.snapshot.xlsx`. The user then:
1. Edits Capacity directly to explore a reallocation
2. Opens the snapshot side-by-side in Excel to compare
3. If the new plan is bad, restores from snapshot

This eliminates a duplicated sheet, toggle formula complexity, and extra reader logic — while providing the same safe-replanning capability through a simpler mechanism.

**CLI**:
- `--snapshot file.xlsx`: copies file to `<stem>.snapshot.xlsx`
- `--restore file.xlsx`: copies `<stem>.snapshot.xlsx` back to file (with timestamped backup of current state first)

### 5. Excel-First WP Configuration

WP definitions move from `config.py` to Excel as primary source.

- **First generation**: `config.py` `WORK_PACKAGES` seeds the Plan sheet's config table (as today)
- **After first generation**: user edits WPs directly in Excel — add/remove rows, change FTE, adjust periods
- **`--update` mode**: `reader.py` extracts WP definitions from Excel config table. Builders receive WP list from reader, not from `config.py` imports.
- **`config.py` retains**: `PROJECT_MONTHS`, `MONTH_COL_START`, `DEFAULT_RATE`, `PROJECT_START`, `PROJECT_TITLE` as structural constants. `WORK_PACKAGES`, `TEAM` (new), and `EXTERNAL_COMMITMENTS` (new) become initial seeds only.
- **Guard rail**: if `--update` finds zero WPs in Excel, it refuses with error message suggesting regeneration from defaults.

### 6. Team Sheet — WP Share Seeds (generation-time only)

On first generation, the Capacity sheet's programme WP rows are seeded using:

```
cell value = (WP share from Team sheet) × (WP total FTE ÷ WP active months)
```

After that, the Team sheet shares are **reference only** — the Capacity sheet monthly values are the source of truth. Editing a share on the Team sheet does NOT automatically redistribute Capacity values (that would destroy manual edits).

A user can manually re-seed by deleting Capacity rows and running `--update` — the tool detects missing rows and re-generates from current shares.

### 7. Scenario Planning Workflow (using snapshots)

1. `python -m budget_planner --snapshot output/budget_plan.xlsx` — saves baseline
2. Edit Capacity sheet directly in Excel — explore reallocation
3. Open snapshot side-by-side to compare visually
4. Happy? Keep the edits. Not happy? `python -m budget_planner --restore output/budget_plan.xlsx` — restores baseline (current state backed up with timestamp first)

### 8. Generation Modes

**`--demo`** (default):
- Plan: config table + Planned FTE as formulas summing Capacity
- Team: example roster (4 people, realistic Max FTE) + WP shares summing to 1.0
- Capacity: all person-WP rows with even distribution + example external commitments with some months showing overload (red/amber highlights visible)
- Actual FTE: zeros
- Summary: fully calculated via formulas

**`--skeleton`** (production use):
- Plan: config table filled, Planned FTE formulas in place (will show values once Capacity has data)
- Team: person names from config seed, Max FTE = 1.0, WP Shares empty
- Capacity: person headers and row structure in place, monthly values **empty**
- Actual FTE: zeros
- Summary: formulas in place
- First cell of each empty block gets a cell comment explaining what to enter

**CLI**:
```
python -m budget_planner                            # --demo (default)
python -m budget_planner --skeleton                 # production skeleton
python -m budget_planner --update file.xlsx         # reads from Excel, preserves edits
python -m budget_planner --snapshot file.xlsx       # save snapshot for scenario comparison
python -m budget_planner --restore file.xlsx        # restore from snapshot (backs up current first)
```

### 9. config.py Changes

```python
# Existing (unchanged — structural constants)
PROJECT_TITLE = "Programme Project Budget Plan"
PROJECT_START = datetime(2026, 7, 1)
PROJECT_MONTHS = 60
DEFAULT_RATE = 7500
MONTH_COL_START = 3

# Existing (becomes seed only — Excel config table is primary after first generation)
WORK_PACKAGES = [
    {"id": "WP1", "name": "Project Management", "fte": 4.0, "start": 1, "end": 60},
    ...
]

# New (seed only — Excel Team sheet is primary after first generation)
TEAM = [
    {"name": "Alice", "max_fte": 1.0, "wps": {"WP1": 0.25, "WP2": 0.50, "WP3": 0.25}},
    {"name": "Bob",   "max_fte": 1.0, "wps": {"WP1": 0.25, "WP3": 0.50, "WP7": 0.25}},
    {"name": "Carol", "max_fte": 1.0, "wps": {"WP4": 0.40, "WP5": 0.40, "WP6": 0.20}},
    {"name": "Dave",  "max_fte": 1.0, "wps": {"WP3": 0.25, "WP4": 0.60, "WP8": 0.15}},
]

# New (seed only — Excel Capacity sheet is primary after first generation)
EXTERNAL_COMMITMENTS = [
    {"person": "Alice", "project": "ProjectX", "fte": 2.0, "start": 1, "end": 11},
    {"person": "Bob",   "project": "ProjectX", "fte": 2.0, "start": 1, "end": 60},
]
```

## Architecture Decisions (from review)

### AD1: Always rebuild formula sheets on --update

On every `--update` run, preserve input sheets (Capacity monthly values, Actuals, Team roster/shares, Plan config table). Regenerate formula sheets (Plan Planned FTE block, Summary) from scratch using the current Capacity layout. This eliminates stale formula risk from row count changes. Regeneration takes milliseconds.

### AD2: Auto-detect old workbook format on --update

If `--update` finds no Capacity sheet, treat it as a v1 workbook:
1. Read config table, planned values, and actual values from old format
2. Generate new-format workbook with Team + Capacity sheets
3. Assign all planned FTE to a placeholder person ("Unassigned") that the user redistributes
4. Preserve actual values in the Plan Actual FTE block

No separate `--migrate` flag needed — detection is automatic.

### AD3: Test strategy for breaking changes

- Rewrite `test_update_preserves_manual_planned_edits` as Capacity round-trip test (preserves Capacity edits, Plan formulas reflect them)
- Add backward compatibility test: `--update` on v1 workbook produces valid v2 workbook (error-free, not a crash)

## Mitigations (from pre-mortem)

| Risk | Mitigation |
|------|------------|
| Workbook complexity (6 sheets, cross-refs) | Clear _META docs per sheet, consistent color conventions, freeze panes on every sheet |
| WP shares drift from Capacity actuals | Shares are seed-only; Capacity is source of truth. _META documents this clearly |
| Round-trip breaks with new sheets | Extend test suite: round-trip tests for every new block (Capacity, Team) |
| Skeleton mode confusing for new users | Cell comments on first cell of each empty block explaining what to enter |
| Excel WP config corrupted/empty | Guard rail: refuse --update with zero WPs, suggest regeneration |
| Plan sheet formulas break when people/WPs added | Reader discovers Capacity layout dynamically; builders adapt to found rows |
| Capacity overwrite destroys manual planning | Double-confirmation prompt on regen without --update |
| Reader can't distinguish programme WP rows from external project rows | Match row ID against config table WP IDs — match = programme, no match = external. Zero user effort, negligible collision risk |

## Test Strategy (TDD — write tests before implementation)

Tests are organized in three layers. Each implementation step writes failing tests first, then code to make them pass. Steps within a layer are ordered so earlier tests enable later ones.

### Step 0: Existing Test Adaptations

Before writing new code, update existing tests to accommodate the v2 design:

| Existing test | Change needed |
|---------------|---------------|
| `test_sheets_exist` | Add `"Team"`, `"Capacity"` to expected list |
| `test_planned_fte_has_formulas` | Keep — formulas change from direct SUM to cross-sheet SUM, but assertion still holds |
| `test_planned_fte_variance` | **Cannot read formula results without recalc** — move to e2e/recalc layer |
| `test_update_preserves_manual_planned_edits` | **Delete** — Planned FTE is formulas now, not input. Replaced by Capacity round-trip test |
| `test_fresh_generation_has_even_distribution` | Move to Capacity layer — distribution is per-person now |
| `test_inactive_months_empty` | Extend to Capacity sheet programme WP rows |

### Step 1: Unit Tests — Config & Data Structures

**File: `tests/test_config.py`** (new)

```
test_team_data_has_required_fields
    Each entry in TEAM has: name (str), max_fte (float > 0), wps (dict str→float)

test_team_wp_shares_reference_valid_wps
    Every WP ID in TEAM[*].wps exists in WORK_PACKAGES[*].id

test_external_commitments_has_required_fields
    Each entry in EXTERNAL_COMMITMENTS has: person (str), project (str),
    fte (float > 0), start (int), end (int), start <= end

test_external_commitments_persons_in_team
    Every person in EXTERNAL_COMMITMENTS exists in TEAM[*].name
```

### Step 2: Unit Tests — Team Sheet Builder

**File: `tests/test_team.py`** (new)

```
test_team_sheet_exists_in_workbook
    build_team_sheet creates a worksheet named "Team"

test_team_roster_has_all_persons
    Column A contains all names from TEAM config

test_team_roster_max_fte_values
    Column B contains max_fte values matching TEAM config

test_team_roster_cells_are_blue_input
    Person name and max_fte cells have blue font color

test_team_share_matrix_wp_columns
    Share matrix has one column per WP, headers matching WP IDs from config

test_team_share_matrix_person_rows
    Share matrix has one row per team member, labels referencing roster

test_team_share_matrix_values_match_config
    Cell values match TEAM[*].wps shares; missing WPs are empty

test_team_share_validation_row_has_sum_formulas
    Row below share matrix has SUM formula per WP column

test_team_tab_color_orange
    Team sheet tab color is orange
```

### Step 3: Unit Tests — Capacity Sheet Builder

**File: `tests/test_capacity.py`** (new)

```
# ── Config block (external commitments) ──

test_capacity_external_config_block_exists
    Config block at top has headers: Person, Project, Total FTE, Start Month, End Month

test_capacity_external_config_values_match_seed
    Config block rows match EXTERNAL_COMMITMENTS entries

test_capacity_external_config_cells_are_blue_input
    All config block value cells have blue font

# ── Person group structure ──

test_capacity_has_person_group_per_team_member
    One header row per person in TEAM, bold with section fill

test_capacity_person_header_refs_team_roster
    Person header cells are formula refs to Team sheet roster cells

test_capacity_person_has_grant_wp_rows
    Each person has rows for every WP in their TEAM[*].wps dict

test_capacity_person_has_external_project_rows
    Each person has rows for their external projects from EXTERNAL_COMMITMENTS

test_capacity_person_has_total_row
    Each person group ends with TOTAL row

test_capacity_person_has_remaining_row
    Each person group ends with REMAINING row after TOTAL

# ── Row content ──

test_capacity_wp_row_id_refs_plan_config
    Programme WP row column A is a formula ref to Plan config table WP ID cell

test_capacity_wp_row_name_refs_plan_config
    Programme WP row column B is a formula ref to Plan config table WP Name cell

test_capacity_wp_rows_are_blue_input
    Monthly cells in programme WP rows have blue font

test_capacity_external_rows_are_blue_input
    Monthly cells in external project rows have blue font

test_capacity_total_row_is_sum_formula
    TOTAL row monthly cells are SUM formulas over all rows in the person group

test_capacity_remaining_row_formula
    REMAINING = Team!MaxFTE - TOTAL (formula, black font)

# ── Seeding ──

test_capacity_grant_wp_seed_value(person, wp)     [parametrized]
    Programme cell = WP share × WP total FTE ÷ WP active months

test_capacity_external_seed_value(person, project)  [parametrized]
    External cell = Total FTE ÷ active months

test_capacity_inactive_months_empty(person, wp)  [parametrized]
    Months outside WP active period are empty with grey fill

# ── Validation columns ──

test_capacity_row_sum_column_has_formula
    Rightmost validation column has SUM of monthly cells

test_capacity_row_budget_column
    Programme rows: person share × WP budget FTE; external rows: Total FTE from config

test_capacity_row_variance_column
    Variance = Budget - Sum, formula

# ── Layout return ──

test_capacity_returns_layout_dict
    build_capacity_sheet returns dict with person_groups keyed by name,
    each containing row numbers for header, wp_rows (dict wp_id→row),
    external_rows (dict project→row), total_row, remaining_row

test_capacity_tab_color_purple
    Capacity sheet tab color is purple
```

### Step 4: Unit Tests — Plan Sheet Modifications

**File: `tests/test_plan_v2.py`** (new, supplements existing test_workbook.py)

```
test_planned_fte_cells_are_sum_formulas_from_capacity
    Each Planned FTE monthly cell is a SUM formula referencing
    Capacity sheet cells (one per person who has that WP)

test_planned_fte_formula_refs_match_capacity_layout
    The cell refs inside each SUM formula match the actual row numbers
    in the capacity_layout dict for that WP × month

test_planned_fte_cells_have_black_font
    Planned FTE cells have black font (formula), not blue (input)

test_planned_fte_inactive_months_empty
    Months outside WP active period still empty (unchanged from v1)

test_plan_config_table_unchanged
    Config table structure and values identical to v1

test_plan_actual_block_unchanged
    Actual FTE block structure, styling, and values identical to v1
```

### Step 5: Unit Tests — Reader Extensions

**File: `tests/test_reader_v2.py`** (new, supplements existing test_roundtrip.py)

```
# ── Version detection ──

test_detect_v2_workbook
    Workbook with Team + Capacity sheets detected as v2

test_detect_v1_workbook
    Workbook without Capacity sheet detected as v1

# ── Team reader ──

test_read_team_roster_names_and_fte
    Extracts list of {name, max_fte} matching what was written

test_read_team_share_matrix
    Extracts dict[person][wp_id] = share matching written values

# ── Capacity reader ──

test_read_capacity_discovers_person_groups
    Finds all person header rows by section fill + bold heuristic

test_read_capacity_classifies_grant_vs_external_rows
    Programme rows matched by WP ID against config table; others = external

test_read_capacity_extracts_monthly_values
    Per-person, per-row monthly values match what was written

test_read_capacity_handles_variable_geometry
    Person with 2 external projects + person with 0 external projects
    both discovered correctly (different row counts per group)

# ── Plan config reader ──

test_read_plan_config_extracts_wp_definitions
    Reads WP ID, name, FTE, start, end from Excel config table
    (not from config.py import)

# ── Guard rails ──

test_read_empty_config_table_raises_error
    Config table with zero WP rows raises ValueError with helpful message
```

### Step 6: Unit Tests — CLI & Snapshot

**File: `tests/test_cli.py`** (new)

```
test_cli_default_is_demo_mode
    Running without flags produces demo workbook (Capacity has values)

test_cli_skeleton_mode
    --skeleton produces workbook with empty Capacity monthly cells

test_cli_skeleton_has_cell_comments
    First empty cell in each block has a comment explaining what to enter

test_cli_snapshot_creates_copy
    --snapshot file.xlsx creates file.snapshot.xlsx

test_cli_snapshot_is_exact_copy
    Snapshot file is byte-identical to original

test_cli_restore_overwrites_from_snapshot
    --restore file.xlsx copies snapshot back to file

test_cli_restore_creates_timestamped_backup
    Before restoring, current file backed up with timestamp

test_cli_restore_no_snapshot_raises_error
    --restore without existing snapshot gives clear error message
```

### Step 7: Integration Tests — Cross-Module Interactions

**File: `tests/test_integration.py`** (new)

```
# ── Data flow: Team → Capacity → Plan → Summary ──

test_capacity_to_plan_formula_linkage
    Generate full workbook. For each WP, each month: Plan Planned FTE cell's
    SUM formula references exactly the Capacity rows of people assigned to
    that WP. Verify by parsing the formula string and matching row numbers
    against capacity_layout.

test_team_shares_to_capacity_seed_values
    Generate full workbook. For each person-WP pair:
    Capacity monthly value == Team share × config FTE ÷ active months

test_plan_formulas_sum_to_config_budget(wp)  [parametrized]
    After LibreOffice recalc: Plan Planned FTE SUM column value equals
    config table Total FTE for that WP (verifies no allocation is lost
    during person-level distribution)

# ── Round-trip preservation ──

test_roundtrip_capacity_edits_preserved
    Generate → manually edit Capacity cell → --update → edited value survives

test_roundtrip_team_roster_edits_preserved
    Generate → add person to Team roster → --update → new person present

test_roundtrip_capacity_edit_reflected_in_plan
    Generate → edit Capacity WP1 M3 for Alice to 0.99 → --update →
    Plan WP1 M3 formula still references Alice's row (formula rebuilt)

test_roundtrip_actuals_still_preserved
    Generate → inject actuals → --update → actuals survive (regression guard)

# ── v1 → v2 migration ──

test_v1_update_creates_team_and_capacity_sheets
    Generate v1 workbook (current code) → new code --update →
    output has Team + Capacity sheets

test_v1_update_assigns_to_unassigned_person
    v1 → v2 migration creates "Unassigned" person holding all planned FTE

test_v1_update_preserves_actuals
    v1 workbook with actuals → --update → actuals survive migration

test_v1_update_produces_valid_v2
    v1 → v2 output passes all structural checks (sheets exist, formulas valid)

# ── Guard rails ──

test_update_refuses_empty_config_table
    Workbook with config table rows deleted → --update → error, not crash

test_double_confirmation_on_overwrite(monkeypatch)
    Fresh generation when Capacity data exists: prompts twice, defaults to No
```

### Step 8: E2E Tests — Full Workflow with Recalc

**File: `tests/test_e2e.py`** (new)

```
# ── These tests require LibreOffice for formula recalculation ──

test_demo_mode_recalc_zero_errors
    Generate demo workbook → recalc → 0 formula errors across all 5 sheets

test_skeleton_mode_recalc_zero_errors
    Generate skeleton workbook → recalc → 0 formula errors
    (formulas reference empty cells but must not produce #REF or #NAME)

test_demo_mode_plan_planned_equals_capacity_sums
    After recalc: for each WP, Plan Planned FTE monthly values equal
    sum of all person allocations from Capacity (data_only=True read)

test_demo_mode_overcommitment_detection
    Generate demo with known overload month → recalc →
    verify TOTAL cell has red fill (conditional formatting)

test_demo_mode_approaching_limit_detection
    TOTAL > 0.8 × Max FTE → amber fill

test_demo_mode_negative_remaining
    REMAINING < 0 → red fill

test_full_roundtrip_with_recalc
    Generate → edit Capacity → --update → recalc → Plan values reflect edits
    (end-to-end data integrity through Excel formula engine)

test_snapshot_restore_roundtrip
    Generate → --snapshot → edit Capacity → --restore →
    workbook matches pre-edit state (binary compare or value compare)

test_summary_reflects_capacity_changes
    Generate → edit Capacity → --update → recalc →
    Summary annual totals reflect the Capacity edit

test_formula_count_increased
    v2 workbook has more formulas than v1 (Capacity TOTAL/REMAINING +
    Plan cross-sheet SUMs). Verify recalc reports higher formula count.
```

### Implementation Order (dependency graph)

```
Step 0 (adapt existing tests)
  ↓
Step 1 (config) ──────────────────────────┐
  ↓                                       │
Step 2 (Team builder) ───┐                │  ← can parallelize 2 & 5a
  ↓                      │                │
Step 3 (Capacity builder)│                │
  ↓                      ↓                ↓
Step 4 (Plan v2) ← needs capacity_layout
  ↓
Step 5 (Reader) ← needs all sheet formats defined
  ↓
Step 6 (CLI/Snapshot) ← needs workbook orchestrator updated
  ↓
Step 7 (Integration) ← needs all modules working together
  ↓
Step 8 (E2E/Recalc) ← needs full pipeline + LibreOffice
```

**Parallelizable groups:**
- Steps 1 + 2 can run in parallel (no file overlap)
- Steps 5 + 6 can partially overlap (reader and CLI are separate files)
- Steps 7 + 8 are sequential (integration first, then e2e validates with recalc)

**File overlap constraints:**
- Steps 3 and 4 both affect `workbook.py` (orchestrator) → sequential
- Step 5 (reader.py) and Step 7 (test_integration.py) don't overlap on source files but Step 7 depends on Step 5 output

## Out of Scope

- Automated Redmine integration
- Per-deliverable / milestone tracking within WPs
- Multi-partner consortium view
- Actuals import from timesheet systems
- Chart/dashboard visualizations
- Sheet protection / cell locking
- Audit trail / replan log (future consideration)
