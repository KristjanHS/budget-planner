# Budget Planner — Architecture Diagram Design

**Date:** 2026-04-03
**Status:** Design complete, not yet implemented
**Output:** `docs/grant-architecture.html` (self-contained, no external dependencies)

## Purpose

A two-layer HTML architecture diagram covering both the user-facing project management flow and the technical build pipeline. Primary audiences: project managers (how to use it) and developers/AI agents (how it works).

## Page Structure

1. Header with gradient banner
2. Six SVG diagram sections (details below)
3. Footer with generation date and version

**Color palette:**
- Blue `#4299e1` — data, config, input
- Orange `#ed8936` — processing, builders
- Purple `#9f7aea` — capacity, planning
- Green `#48bb78` — outputs, success, OK status
- Red `#e53e3e` — alerts, critical
- Gold `#d69e2e` — trends, warnings
- Grey `#a0aec0` — inactive, infrastructure

**Layout:** max-width 1200px, system-ui font, responsive. All SVGs use `viewBox` for scaling.

---

## Layer 1 — User-Facing Flow

### Diagram 1: Project Lifecycle Overview

Horizontal swim-lane with 4 phases connected by thick arrows. A feedback loop arrow curves from Monitor back to Plan & Allocate.

| Phase | Primary Sheets | User Actions | Outputs |
|-------|---------------|--------------|---------|
| **Configure** | Config, Plan config | Set params, define WPs, set FTE rate | WP table, named ranges |
| **Plan & Allocate** | Team, Capacity, Plan FTE | Assign FTE per person/month, check load, balance | Planned FTE grid |
| **Track Actuals** | Actuals | Enter EUR per WP per month | Cumulative spend |
| **Monitor & Escalate** | Trends, Summary | Update current_month, review RAG | Burn rate alerts, chart |

Below the swim-lane: a thin bar showing CLI commands that trigger each phase:
- `python -m budget_planner init` → Configure + Plan
- `python -m budget_planner rebuild` → iteration on any phase
- Manual Excel editing → Track
- Update `current_month` cell → Monitor

### Diagram 2: Excel Sheet Relationships

Central hub diagram. Config sits top-center as origin. Sheets radiate in build order.

**Nodes (8 sheets):**
- `_META` — floats to the side, dashed arrows from all sheets (documentation, not data flow)
- `Config` (blue tab `#4472C4`) — project params + tbl_WorkPackages
- `Plan` (no tab color) — global params + config table + Planned FTE block
- `Team` (orange tab `#FFA500`) — roster + WP share matrix
- `Capacity` (purple tab `#7030A0`) — per-person monthly FTE, PRIMARY SURFACE (emphasized)
- `Actuals` (blue tab `#0000FF`) — EUR per WP per month
- `Summary` (green tab `#008000`) — annual aggregation
- `Trends` (gold tab `#FFD700`) — burn rate, exhaustion, person load, chart

**Arrows with labels:**
- Config → Plan: `project_start`, `fte_rate` (named ranges)
- Plan config → Capacity: green cross-sheet refs (WP ID/Name)
- Plan config → Actuals: green cross-sheet refs (WP ID/Name)
- Team → Capacity: green cross-sheet refs (person names, max FTE)
- **Capacity → Plan** (reverse, highlighted): SUM formulas (`_apply_planned_formulas`)
- Plan → Trends: `partner_ceiling` named range
- Actuals → Trends: cumulative EUR refs
- Actuals → Summary: annual aggregation formulas
- Plan → Summary: annual FTE aggregation
- Trends ← `current_month` named range (Trends!B1)

### Diagram 3: RAG Escalation Model

Traffic-light table with three color columns (Green / Amber / Red) and rows grouped by sheet:

**Capacity (person load):**
| Green | Amber | Red |
|-------|-------|-----|
| TOTAL ≤ 0.8 × Max FTE | TOTAL > 0.8 × Max FTE | TOTAL > Max FTE |
| REMAINING ≥ 0 | | REMAINING < 0 |

**Actuals (WP spend):**
| Green | Amber | Red |
|-------|-------|-----|
| Variance ≥ 0 | % Ceiling > 85% | Variance < 0 (overspend) |
| % Ceiling ≤ 85% | Breach In < 12 months | % Ceiling > 100% |
| Breach In ≥ 12 months | | Breach In < 6 months |

**Trends — burn rate:**
| Green | Amber | Red |
|-------|-------|-----|
| \|Burn Delta\| ≤ 5% | Delta 5–15% (overspend) | Delta > 15% |
| Status = "OK" | Delta < -10% ("SLOW") | Status = "OVER" |
| Exhaustion ≥ WP end | Exhaustion within 3 months of end | Exhaustion > 3 months early |

**Trends — person load:**
| Green | Amber | Red |
|-------|-------|-----|
| Remaining avg ≤ 0.8 × Max | 0.8 × Max < avg ≤ Max | Remaining avg > Max FTE |
| Status = "OK" | Status = "WATCH" | Status = "OVER" |

**Annotation:** Lower-priority rules (amber) applied first in openpyxl; higher-priority (red) added second and overrides. Sidebar: font color convention (Blue = input, Black = formula, Green = cross-sheet ref).

---

## Layer 2 — Technical Pipeline

### Diagram 4: Build Pipeline

Left-to-right pipeline with an init/rebuild fork at the start converging into a 14-step common pipeline.

**Fork — Init path:**
- Source: `config.py` seed data (WPs, team, external commitments)
- Modes: `--skeleton` (empty Capacity) vs demo (pre-filled)

**Fork — Rebuild path:**
- Source: existing `.xlsx` via `reader.py`
- Steps: read Config → read Team → read Capacity FTE → read Actuals EUR → validate cross-refs
- Gate: validation must pass before entering common pipeline

**Common pipeline (14 steps):**
1. Config sheet (`config_sheet.py`) — params + tbl_WorkPackages + named ranges
2. Plan sheet (`plan.py`) — global params, config table, seed Planned FTE
3. Team sheet (`team.py`) — roster + WP share matrix
4. Capacity sheet (`capacity.py`) — per-person FTE grid, external commitments, CF
5. **Patch Plan ← Capacity** (`plan.py:_apply_planned_formulas`) — overwrite Plan FTE with SUM formulas from Capacity. _Highlighted differently — this is the critical cross-sheet wiring._
6. Trends skeleton (`trends.py`) — create `current_month` cell + named range early
7. Actuals sheet (`actuals.py`) — EUR grid, cumulative formulas, CF
8. Summary sheet (`summary.py`) — annual aggregation
9. Trends content (`trends.py`) — burn rate table, person load, chart, CF
10. _META sheet (`meta.py`) — AI-readable documentation from all layout dicts
11. Sheet protection — all sheets, password "grant", blue cells unlocked
12. Column grouping (`styles.py:apply_year_groups`) — 12-month year groups, expand current year
13. Save `.xlsx` — timestamped backup of existing file first
14. Post-save: `scripts/recalc.py` (external, LibreOffice headless) — mandatory formula recalculation

**Annotations:**
- Step 5 highlighted with a distinct border/color — the formula patch is the key architectural decision
- Steps 6 and 9 connected by a dashed line showing the Trends split (skeleton early, content late)
- Skeleton/demo toggle icon on step 4
- `layout` dict shown as a return value flowing from builders back to `workbook.py`

### Diagram 5: Round-Trip / Rebuild Flow

Circular diagram with 4 nodes forming a loop:

**Node 1 — Live Workbook** (top, emphasized)
- User edits in Excel: Capacity FTE (blue), Actuals EUR (blue), current_month
- This is the operational state between rebuilds

**Node 2 — Extract** (right)
- `reader.py` reads: Config params, Team roster, Capacity FTE values, Actuals EUR values, current_month from Trends!B1
- Marks what's preserved: Capacity FTE ✓ (normal) / ✗ (skeleton), Actuals EUR ✓ (always)

**Node 3 — Validate** (bottom)
- Cross-reference checks: WP IDs unique, persons exist in team, external commitment persons valid
- Orphan detection: removed persons or deleted WPs → data silently dropped + warning to console and _META

**Node 4 — Regenerate** (left)
- Full build pipeline with preserved values injected at appropriate steps
- Timestamped backup created before overwriting
- Output: updated `.xlsx` → flows back to Live Workbook

**Sidebar — Snapshot/Restore:**
- Branch off Live Workbook node
- `--snapshot` saves `.snapshot.xlsx` copy
- `--restore` reverts to snapshot (current file backed up first)
- Arrows show both directions

**Preservation legend:**
- ✓ Always preserved: Actuals EUR
- ✓ Preserved (normal rebuild): Capacity FTE, current_month
- ✗ Cleared (skeleton rebuild): Capacity FTE
- Dropped with warning: orphaned data (removed persons/WPs)

### Diagram 6: Module Dependency Map

Three horizontal bands with call arrows flowing downward:

**Band 1 — Orchestration (top):**
- `__main__.py` — CLI argument parsing, subcommand dispatch
- `workbook.py` — `init_workbook()`, `rebuild_workbook()`, `snapshot_workbook()`, `restore_workbook()`
- Arrow: `__main__.py` → `workbook.py`

**Band 2 — Builders (middle):**
Eight modules called by `workbook.py` in build order:
- `config_sheet.py` → `plan.py` → `team.py` → `capacity.py` → `actuals.py` → `summary.py` → `trends.py` → `meta.py`
- Special arrow: `plan.py` ↔ `capacity.py` labeled "formula patch" (bidirectional dependency)
- All return `layout` dicts upward to `workbook.py`

**Band 3 — Infrastructure (bottom):**
- `styles.py` — font/fill/border constants, `style_cell()`, `apply_year_groups()`
- `reader.py` — round-trip extraction + validation (dashed arrow from `workbook.py`, rebuild-only)
- `config.py` — seed data (dashed arrow from `workbook.py`, init-only)

**External (outside bands, dashed border):**
- `scripts/recalc.py` — LibreOffice headless recalculation, required post-processing step

**Annotations:**
- Dashed arrows for conditional dependencies (rebuild-only, init-only)
- Solid arrows for always-used dependencies
- `layout` dict labeled on return arrows from builders to `workbook.py`

---

## Implementation Notes

- Single self-contained HTML file, no external CSS/JS dependencies
- All diagrams as inline SVGs with `viewBox` for responsive scaling
- Use `<defs>` block for shared arrow markers and gradient definitions
- Sheet nodes use actual Excel tab colors as header fills
- Interactive: none required (static document), but CSS hover tooltips on nodes welcome
- Target output: `docs/grant-architecture.html`
