# Documentation Refactoring Plan: Remove Redundancy & Over-Engineering

**Date**: 2026-04-07
**Status**: Draft — review before executing

---

## Problem

16 active doc files (2,876 lines) + 15 archived plans (4,920 lines) = ~7,800 lines total. A previous audit (2026-04-02) cleaned up SPEC.md and DESIGN_DECISIONS.md, but then smart-docs generation added a new layer of redundancy (Project Overview, Architecture Overview, Workflow Overview, deep-dive/ directory).

---

## Findings

### F1: `docs/1. Project Overview.md` is 90% redundant with `README.md`
Both contain: project description, features, tech stack, project structure tree, getting started commands. Only unique content: a 6-line codebase metrics table. **79 lines for 6 lines of unique value.**

### F2: `docs/2. Architecture Overview.md` over-engineers a simple project
C4 diagrams, build-order sequence diagram, 5 architectural patterns, 3 design decisions, per-module breakdowns — for a 3,600-LOC project with 12 files. "Module Breakdown" section duplicates `.claude/rules/architecture.md` and README.md. **The mermaid diagrams are the only unique value.**

### F3: `docs/3. Workflow Overview.md` partially duplicates `docs/deep-dive/Reader and Rebuild.md`
- "Rebuild Round-Trip: What Gets Preserved" table appears in both
- Both have rebuild flow diagrams (different detail levels)
- Named ranges table also in CONVENTIONS.md cross-sheet reference map

### F4: `CONVENTIONS.md` overlaps with `docs/deep-dive/Layout and Styling.md`
- Font colors: both files
- Fill colors: both files
- Column grouping: both files
- Unique to CONVENTIONS.md: CF rules, cross-sheet ref map
- Unique to Layout and Styling: column positioning, style_cell() doc

### F5: `docs/cloud-env-setup.md` is 300 lines of incident postmortem posing as setup docs
Useful content: "install LibreOffice" (3 lines), pre-flight script (40 lines), failure modes (50 lines). The rest: narrative about a cloud agent misdiagnosis, repeated rules from CLAUDE.md/dev-workflow.md, data_only=True mental model (also in Reader deep-dive).

### F6: Research docs too large for permanent reference (1,016 lines)
Research A (463 lines) and B (553 lines) were produced for a specific planning session. Operational insights already extracted into grant-domain.md + ideas-to-explore.md. Remaining bulk: source citations, methodology, detailed tables unlikely to be referenced.

### F7: `docs/plans/old_already_implemented/` is dead weight (4,920 lines)
15 archived plan files. Git history preserves them. Identified in 2026-04-02 audit but never deleted.

### F8: `.claude/rules/architecture.md` file tree duplicates README.md
Both have project structure tree. architecture.md is auto-loaded for Claude on `**/*.py` — right audience, but content also in README.md and Project Overview.

### F9: CLAUDE.md has two overlapping lookup tables
"Reference Index" (lines 39-49) and "Before Modifying Code" (lines 88-96) both map triggers to documents. Four entries appear in both.

---

## Target State: 9 files, ~1,070 lines

| File | Purpose | Lines (est.) |
|------|---------|-------------|
| `README.md` | Human-facing: what, why, quick start, structure, Windows setup | ~120 |
| `CLAUDE.md` | Agent: commands, critical rules, single lookup table | ~85 |
| `.claude/rules/dev-workflow.md` | Build/test, openpyxl patterns, invariants | ~65 (no change) |
| `.claude/rules/architecture.md` | File tree only | ~44 (no change) |
| `.claude/rules/common-tasks.md` | Modification recipes | ~28 (no change) |
| `docs/grant-domain.md` | programme cost model, funding, reallocation | ~300 (no change) |
| `docs/architecture.md` | Mermaid diagrams + build order (unique value from current 2. + 3.) | ~150 |
| `docs/internals.md` | Builders, reader, layout/styling merged (from deep-dive/ + CONVENTIONS.md unique content) | ~200 |
| `docs/cloud-env-setup.md` | Trimmed to setup script + failure modes only | ~80 |

---

## Files to Delete

| File | Reason |
|------|--------|
| `CONVENTIONS.md` | Unique content (CF rules, cross-sheet ref map) → `docs/internals.md`. Font/fill already in styles.py. |
| `docs/1. Project Overview.md` | 90% redundant with README.md. Metrics table → README.md. |
| `docs/2. Architecture Overview.md` | Replaced by slimmed `docs/architecture.md` (diagrams only). |
| `docs/3. Workflow Overview.md` | Diagrams → `docs/architecture.md`, preservation table → `docs/internals.md`. |
| `docs/deep-dive/` (3 files) | Merged into `docs/internals.md`. |
| `docs/research/` (2 files) | Archive or delete. Operational insights already in grant-domain.md + ideas-to-explore. |
| `docs/plans/old_already_implemented/` (15 files) | Git history is the archive. |

---

## Changes to Kept Files

**CLAUDE.md**: Merge "Reference Index" and "Before Modifying Code" into a single lookup table. Remove stale entries. Update doc paths.

**README.md**: Add codebase metrics from Project Overview. Update doc links table to match new file structure.

---

## Execution Steps

### Group A — Independent deletes (parallel)

| Step | Action |
|------|--------|
| A1 | Delete `docs/plans/old_already_implemented/` (15 files, 4,920 lines) |
| A2 | Delete `docs/1. Project Overview.md` after copying metrics table to README.md |
| A3 | Move `docs/research/` to `docs/research-archive/` (or delete — git has them) |

### Group B — Merge work (parallel, but each step is self-contained)

| Step | Action | Sources → Target |
|------|--------|-----------------|
| B1 | Create `docs/architecture.md` | Mermaid diagrams from `2. Architecture Overview` + sequence diagrams from `3. Workflow Overview` + cross-sheet formula graph from `3. Workflow Overview`. Drop all prose that restates what code shows. |
| B2 | Create `docs/internals.md` | `deep-dive/Sheet Builders.md` + `deep-dive/Reader and Rebuild.md` + `deep-dive/Layout and Styling.md` + CF rules and cross-sheet ref map from `CONVENTIONS.md`. Drop font/fill tables (already in styles.py). |
| B3 | Trim `docs/cloud-env-setup.md` | Keep: Required System Dependencies, Pre-Flight script, Common Failure Modes. Delete: narrative, repeated rules, data_only mental model, "Missing Context" section. |

### Group C — Cleanup (after A + B)

| Step | Action |
|------|--------|
| C1 | Delete source files: `CONVENTIONS.md`, `docs/2. Architecture Overview.md`, `docs/3. Workflow Overview.md`, `docs/deep-dive/` |
| C2 | Update CLAUDE.md: merge two lookup tables into one, update all doc paths |
| C3 | Update README.md: add metrics, update doc links table |

---

## Estimated Reduction

| Metric | Before | After | Reduction |
|--------|--------|-------|-----------|
| Active doc files | 16 | 9 | -44% |
| Active doc lines | ~2,876 | ~1,070 | -63% |
| Archived plan lines | 4,920 | 0 | -100% |
| Total lines | ~7,800 | ~1,070 | -86% |

---

## What Stays Unchanged

- `.claude/rules/dev-workflow.md` — operational, non-duplicated
- `.claude/rules/architecture.md` — auto-loaded for code work, concise
- `.claude/rules/common-tasks.md` — recipes, concise
- `docs/grant-domain.md` — standalone domain reference, well-structured
- `docs/plans/2026-04-05-research-ideas-to-explore.md` — active backlog
