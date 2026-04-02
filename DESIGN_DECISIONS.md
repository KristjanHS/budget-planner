# DESIGN_DECISIONS.md — Brainstorming Session Record

Captured from Claude.ai brainstorming session, March 2026.

## Problem Statement

Plan FTE budget per work package per month across a 5+ year EU-funded project. Track actuals against plan to prevent budget overruns. Single organisation share only.

## Key Decisions Made

### 1. Distribution Logic

**Decision**: Even spread across active months as default, with manual override per cell.

**Rationale**: Funder only checks total FTE per WP and that time isn't logged outside active periods. No need for complex ramp-up/down profiles. Even spread gives a baseline the user adjusts as needed.

### 2. Sheet Layout — Mix of Approach A + B

**Decision**: Single flat Plan sheet (Approach A) plus a Summary sheet (Approach B).

**Alternatives considered**:
- A: Single flat sheet only — everything visible but no annual summary
- B: Year-per-sheet — compact per year but loses full timeline view
- C: Input/output separation — too complex for the use case

**Rationale**: Full timeline visibility for day-to-day work, plus annual aggregation for reporting.

### 3. Planned vs Actual Layout — Separate Blocks

**Decision**: All Planned rows in one block, all Actual rows in a separate block below.

**Alternative**: Stacked rows (Planned/Actual/Remaining per WP grouped together).

**Rationale**: Easier to scan; easier to paste actuals from timesheet exports.

### 4. Audit Protection

**Decision**: Grey out cells outside each WP's active period (start month to end month).

**Rationale**: Funder auditors flag time logged when a WP is not supposed to be active. Visual and structural prevention.

### 5. Euro Derivation

**Decision**: Single FTE rate in one cell (Plan!$E$2), all EUR calculated as FTE × rate.

**Rationale**: Official reporting unit is FTE. EUR is internal only. Different team member rates not needed at this stage.

### 6. Actuals Tracking

**Added during design**: User requested ability to track monthly actuals and see remaining budget per WP to prevent overruns.

**Implementation**: Second monthly grid block (identical structure to Planned), with validation showing Remaining = Budget - Cumulative Actual, red fill when negative.

### 7. AI-Agent Documentation

**Decision**: Dedicated `_META` sheet inside the workbook with full structural documentation. Plus external markdown docs (SPEC.md, CONVENTIONS.md, CLAUDE.md) for Claude Code.

**Rationale**: User explicitly requested the spreadsheet be self-documenting for AI agents as primary audience, to enable iterative improvement.

### 8. Per-Person Capacity Planning (v2)

**Decision**: Team sheet for roster + WP share seeds, Capacity sheet for per-person monthly FTE allocation with cross-sheet aggregation back to Plan.

**Rationale**: The planner needs to see individual overcommitment before it happens. WP-level totals hide the constraint that each person has a fixed monthly FTE capacity across all projects (programme + external).

**Key design choices**:
- Team roster is a separate sheet so shares can be edited without touching allocation data
- Capacity groups WP rows + external project rows per person, with TOTAL and REMAINING summary rows
- Conditional formatting: amber at >80% utilization, red at >100%
- Plan Planned FTE is now derived via cross-sheet SUM formulas from Capacity (not manually entered)
- Demo mode seeds values from `share × WP FTE ÷ active months`; skeleton mode leaves empty

### 9. CLI Modes — Skeleton, Snapshot, Restore

**Decision**: `--skeleton` for production use (empty Capacity), `--snapshot`/`--restore` for scenario comparison.

**Rationale**: Demo mode with pre-filled values is useful for first-time exploration, but production users need a blank slate. Snapshot/restore enables "what-if" scenario comparison without manual file copying.

## Open Items for Future Iterations

1. Non-personnel cost categories (travel, equipment, subcontracting)
2. Auto-correction of rounding variance in last active month
3. Timesheet import automation
4. Dashboard/charts sheet
5. Sheet protection with unlocked input cells
6. Data validation on start/end month fields
7. Reporting period alignment (if not calendar-year based)
