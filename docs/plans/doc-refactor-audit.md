# Documentation Audit & Refactoring Plan

**Date**: 2026-04-02
**Status**: Proposed

---

## Inventory (16 doc files, ~900 lines active + 3,272 lines archived)

| File | Lines | Role |
|------|-------|------|
| `README.md` | 87 | Public-facing: quick start, features, structure |
| `CLAUDE.md` | 41 | Claude Code agent instructions |
| `SPEC.md` | 150 | Full specification |
| `CONVENTIONS.md` | 97 | Spreadsheet formatting/structural conventions |
| `DESIGN_DECISIONS.md` | 87 | Brainstorming session record |
| `.claude/rules/architecture.md` | 57 | File tree + test coverage list |
| `.claude/rules/dev-workflow.md` | 42 | Build/test/lint + openpyxl patterns + invariants |
| `.claude/rules/common-tasks.md` | 26 | Modification recipes |
| `docs/grant-budget-resource-scheduling-flow.md` | 170 | programme lifecycle flow |
| `docs/grant-actual-cost-budgeting-logic.md` | 143 | programme budgeting logic |
| `docs/plans/old_already_implemented/` (9 files) | 3,272 | Archived plans |
| `docs/diagrams/` (5 files) | — | Visual diagrams (jsx, html, mmd) |

---

## Finding 1: STALE — SPEC.md doesn't know about Actuals or Trends sheets

**Severity**: High

SPEC.md describes 5 sheets: `_META, Plan, Team, Capacity, Summary`. The actual workbook generates **7 sheets** including `Actuals` (EUR actuals per WP from accountant) and `Trends` (burn rate, person load, cumulative chart).

- SPEC.md §Architecture §Sheets table — missing Actuals and Trends
- SPEC.md §Plan Sheet — says "ACTUAL FTE block" lives in Plan. The Plan sheet no longer has any actuals content (confirmed: `grep -c 'actual' plan.py` = 0). Actuals is its own sheet tracking EUR, not FTE.
- SPEC.md §Validation Columns — describes the Plan actual block validation, which moved to Actuals sheet with different columns (Cumulative, % of Ceiling, Breach flag)
- SPEC.md §Known Limitations line 142 — says "Actuals import: Currently manual entry" referencing FTE timesheets, but the Actuals sheet tracks EUR costs, not FTE

**Action**: SPEC.md needs a rewrite to match current architecture. But see Finding 6 re: whether SPEC.md should exist at all.

---

## Finding 2: DUPLICATION — Font/fill/number conventions in 4 places

**Content duplicated across**:
1. `SPEC.md` lines 97–120 (font colors, fill colors, number formats, named ranges)
2. `CONVENTIONS.md` lines 1–35 (font colors, fill colors, number formats — identical content)
3. `CLAUDE.md` line 38 (font color rule, abbreviated)
4. `.claude/rules/dev-workflow.md` line 39 (font color rule, abbreviated)

CONVENTIONS.md and SPEC.md §Conventions are ~95% identical text. The brief mentions in CLAUDE.md and dev-workflow.md are summary references.

**Action**: Single source of truth should be CONVENTIONS.md. SPEC.md should not duplicate it. CLAUDE.md and dev-workflow.md references are fine as brief pointers.

---

## Finding 3: DUPLICATION — Quick start / key commands in 2 places

- `README.md` lines 7–22: install, generate, skeleton, recalc, test
- `CLAUDE.md` lines 9–31: same commands plus --update, --snapshot, --restore

**Action**: README.md targets humans, CLAUDE.md targets Claude. Different audiences — slight duplication is acceptable here. But CLAUDE.md has commands README.md doesn't (update, snapshot, restore). README.md should either list all modes or just point to `--help`.

---

## Finding 4: DUPLICATION — Project structure / file tree in 2 places

- `README.md` lines 36–59: directory tree with descriptions
- `.claude/rules/architecture.md` lines 6–37: same directory tree, slightly more detailed

**Action**: architecture.md is auto-injected for Claude on `**/*.py` paths — it serves a different audience. But it duplicates the file list. Since architecture.md is authoritative for code work and README.md is for humans, keep both but ensure they're consistent. Currently architecture.md is more up-to-date (includes actuals.py, trends.py); README.md is stale (missing both).

---

## Finding 5: DUPLICATION — Design decisions / open items in 3 places

- `DESIGN_DECISIONS.md` lines 79–87: "Open Items for Future Iterations" (7 items)
- `SPEC.md` lines 136–143: "Known Limitations / Future Work" (7 items, nearly identical)
- `README.md` lines 69–79: "Design Decisions" summary that just says "See SPEC.md"

**Action**: One place for decisions, one place for future work. README.md's pointer is fine. SPEC.md and DESIGN_DECISIONS.md should not both maintain the same list.

---

## Finding 6: YAGNI — Do we need both SPEC.md and CONVENTIONS.md as separate files?

SPEC.md (150 lines) contains:
- Business context → useful, not elsewhere
- Sheet architecture → duplicated/stale (see Finding 1)
- Conventions → duplicated from CONVENTIONS.md (see Finding 2)
- CLI modes → duplicated from README.md
- Dependencies → duplicated from README.md
- Known limitations → duplicated from DESIGN_DECISIONS.md

After removing duplication, SPEC.md's unique content is ~40 lines of business context and sheet architecture. Meanwhile CONVENTIONS.md has unique structural detail (cross-sheet reference map, conditional formatting rules, config-to-grid linkage) that SPEC.md doesn't.

**Question**: Should SPEC.md be collapsed into the other files, or retained as the single architecture document?

---

## Finding 7: YAGNI — Two overlapping programme domain docs

- `docs/grant-budget-resource-scheduling-flow.md` (170 lines) — full lifecycle from proposal to close-out
- `docs/grant-actual-cost-budgeting-logic.md` (143 lines) — cost categories, funding rates, budget hierarchy

Overlap:
- Both cover indirect costs (25% flat rate) — flow.md §Phase 1 step 5, logic.md §6
- Both cover funding rates — flow.md §Phase 1 step 6, logic.md §programme-Specific Funding Rates
- Both cover reporting — flow.md §Phase 5, logic.md §Reporting
- Both cover reallocation rules — flow.md §Phase 4+6, logic.md §Reallocation Rules
- Both have nearly identical formulas for `indirect_costs`

**Action**: Merge into one programme domain reference. The flow doc's lifecycle phases and the logic doc's cost category detail are complementary, not separate documents.

---

## Finding 8: STALE — README.md feature list is outdated

README.md line 27: "Plan sheet: WP config table, planned FTE grid, **actual FTE grid** with budget validation"

The Plan sheet no longer has an actual FTE grid. Actuals are in a separate sheet tracking EUR. README also doesn't mention the Actuals or Trends sheets at all.

**Action**: Update README.md features list.

---

## Finding 9: YAGNI — 3,272 lines of archived plans

`docs/plans/old_already_implemented/` contains 9 implemented plans totaling 3,272 lines. These are dead weight — git history preserves them. They serve no purpose for future development.

**Action**: Delete the archive directory. Git history is the archive.

---

## Finding 10: STALE — DESIGN_DECISIONS.md open items

Several "open items" from the brainstorming are now implemented:
- Item 3 "Timesheet import automation" — not applicable; Actuals tracks EUR, not FTE timesheets
- Item 4 "Dashboard/charts sheet" — partially addressed by Trends sheet

The rest may still be valid, but the list hasn't been maintained.

**Action**: Remove implemented/obsolete items. Keep only genuinely open items, or delete the file entirely if the remaining items are speculative.

---

## Finding 11: architecture.md test coverage list is a maintenance trap

`.claude/rules/architecture.md` lines 46–57 list 13 test coverage areas. This duplicates what `pytest tests/ -v` tells you and must be manually updated every time a test is added.

**Action**: Remove the test coverage list. `pytest --co -q` is the source of truth.

---

## Refactoring Plan

### Target state: 6 files, zero duplication

| File | Purpose | Lines (est.) |
|------|---------|-------------|
| `README.md` | Human-facing: what it does, quick start, how to modify | ~60 |
| `CLAUDE.md` | Claude agent: key commands, critical rules | ~40 (no change) |
| `CONVENTIONS.md` | Single source for spreadsheet conventions: font/fill/number, structural patterns, cross-sheet refs, CF rules | ~100 (no change) |
| `docs/grant-domain.md` | Merged programme domain reference: lifecycle + cost logic | ~200 |
| `.claude/rules/architecture.md` | File tree only (no test list, no references section) | ~40 |
| `.claude/rules/dev-workflow.md` | Build/test + openpyxl patterns + invariants | ~42 (no change) |
| `.claude/rules/common-tasks.md` | Modification recipes | ~26 (no change) |

### Files to delete

| File | Reason |
|------|--------|
| `SPEC.md` | Content split: business context → README.md, conventions → already in CONVENTIONS.md, architecture → architecture.md |
| `DESIGN_DECISIONS.md` | Historical record preserved in git. Remaining open items either obsolete or speculative. |
| `docs/grant-budget-resource-scheduling-flow.md` | Merged into `docs/grant-domain.md` |
| `docs/grant-actual-cost-budgeting-logic.md` | Merged into `docs/grant-domain.md` |
| `docs/plans/old_already_implemented/` | Dead archive, git history is sufficient |

### Dependency graph

```
Step 1: Merge programme domain docs              (independent)
Step 2: Update README.md                   (independent)
Step 3: Update architecture.md             (independent)
Step 4: Delete SPEC.md                     (after step 2 — README absorbs business context)
Step 5: Delete DESIGN_DECISIONS.md         (after step 2)
Step 6: Delete archived plans              (independent)
Step 7: Delete old programme docs                (after step 1)

Parallel groups:
  Group A: Steps 1, 2, 3, 6 (all independent)
  Group B: Steps 4, 5, 7 (after Group A)
```

### What stays unchanged

- `CLAUDE.md` — already minimal and correct
- `CONVENTIONS.md` — already the best-structured doc, single source of truth for formatting
- `.claude/rules/dev-workflow.md` — practical, non-duplicated
- `.claude/rules/common-tasks.md` — practical, non-duplicated
- `docs/diagrams/` — visual assets, no text duplication
