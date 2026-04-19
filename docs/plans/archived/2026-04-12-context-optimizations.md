# Plan: Top 5 Context Optimizations for budget-planner

## Context

Research into 2026 best practices for AI-agent-oriented documentation revealed that this project's progressive disclosure architecture is already well above average. However, the "would removing this cause mistakes?" acid test reveals ~51% of always-loaded documentation tokens are unnecessary for routine coding sessions. Five targeted changes address this.

**Current always-loaded docs budget** (when touching any .py file):
- Project CLAUDE.md: 4,601 bytes (~1,150 tokens)
- Global CLAUDE.md: 3,588 bytes (~897 tokens)
- dev-workflow.md: 4,998 bytes (~1,249 tokens)
- openpyxl-patterns.md: 3,253 bytes (~813 tokens)
- common-tasks.md: 1,467 bytes (~366 tokens)
- **Total: ~4,475 tokens of docs context per session**

---

## Improvement 1: Add `.claudeignore` (prevention: 500-22K tokens)

**Why**: No exclusion file exists. `docs/grant-architecture.html` alone is 89KB (~22K tokens). The `output/` dir has 4 xlsx files. `docs/plans/` has 5 plan files from prior sessions. `docs/analysis/` has reports. All discoverable by exploratory searches.

**Create** `~/projects/budget-planner/.claudeignore`:
```
# Binary / generated output
output/
*.xlsx

# Build artifacts
*.egg-info/
__pycache__/
.venv/
.ruff_cache/
.pytest_cache/

# Large generated HTML (89KB)
docs/grant-architecture.html

# Completed plans (active plans stay accessible)
docs/plans/old_already_implemented/

# Analysis artifacts from prior sessions
docs/analysis/

# Windows scripts (WSL interop disabled)
*.bat
```

Note: `docs/plans/` itself stays accessible (needed by `/impag`). Only the `old_already_implemented/` subfolder is excluded.

**Verify**: New session, run `Glob("docs/**")` — HTML file should not appear. Run `Glob("output/**")` — nothing should appear.

---

## Improvement 2: Remove meta-documentation from both CLAUDE.md files (~485 tokens)

**Why**: 27 lines of "Information Recording Principles" in project CLAUDE.md (lines 7-35) and 23 lines in global CLAUDE.md (lines 2-26) explain the documentation system's own architecture. This loads every session but is only relevant when asked to record/remember something. The existing file structure is self-documenting.

**Project CLAUDE.md** (`~/projects/budget-planner/CLAUDE.md`):
- Delete lines 7-35 (from `## Information Recording Principles` through blank line before `## Reference Index`)
- Add a 1-line HTML comment: `<!-- L1=this file (always), L2=docs/ (on demand via trigger table), L3=.claude/rules/ (path-gated) -->`
- Result: 93 → ~67 lines

**Global CLAUDE.md** (`~/.claude/CLAUDE.md`):
- Delete lines 2-26 (from `## Information Recording Principles` through blank line before `## Reference Index`)
- Add a 1-line HTML comment: `<!-- L1=this file, L2=~/.claude/rules/ (path-gated), L3=~/.claude/references/ (manual) -->`
- Result: 66 → ~44 lines

**Optionally preserve** by creating `~/.claude/references/recording-principles.md` with the removed content from both files, and adding a trigger row to both Reference Index tables: `| "record this", "remember this", organizing docs | references/recording-principles.md | L1/L2/L3 criteria, what belongs where |`

**Verify**: New session, confirm trigger table and critical rules still visible. Ask "where should I record X?" — Claude should either infer from structure or follow the trigger row to the reference doc.

---

## Improvement 3: Convert `common-tasks.md` from rule to reference doc (~366 tokens)

**Why**: This file is a set of SOPs ("add a WP", "add team member", etc.), not rules or invariants. It loads on every Python file touch but is only relevant when the user explicitly asks to perform one of these tasks. It fails the acid test: removing it from auto-load doesn't cause mistakes when writing code.

**Steps**:
1. Copy `.claude/rules/common-tasks.md` → `docs/common-tasks.md`, stripping the YAML frontmatter
2. Delete `.claude/rules/common-tasks.md`
3. Add trigger row to CLAUDE.md Reference Index:
   `| Add WP, add team member, add actuals, compare scenarios | docs/common-tasks.md | Step-by-step SOPs for common data tasks |`
4. Remove the `common-tasks.md` entry from the "Rules Index" table at bottom of CLAUDE.md

**Verify**: New session, touch a `.py` file — only 2 rules should load (dev-workflow + openpyxl-patterns). Ask "how do I add a WP?" — Claude should find it via trigger table.

---

## Improvement 4: Narrow `openpyxl-patterns.md` path gate (~813 tokens on non-builder sessions)

**Why**: 22 lines of hard-won openpyxl/LO gotchas gated to `**/*.py` load even when editing `config.py`, `__main__.py`, `reader.py`, `layout.py`, `meta.py`, or test files — none of which produce openpyxl output. Only the builder/style/workbook files need these rules.

**Edit** `~/projects/budget-planner/.claude/rules/openpyxl-patterns.md` frontmatter:

From:
```yaml
---
paths: "**/*.py"
---
```

To:
```yaml
---
# Add new sheet builders here when created
paths:
  - "budget_planner/plan.py"
  - "budget_planner/summary.py"
  - "budget_planner/capacity.py"
  - "budget_planner/actuals.py"
  - "budget_planner/trends.py"
  - "budget_planner/config_sheet.py"
  - "budget_planner/styles.py"
  - "budget_planner/workbook.py"
---
```

**Update** Rules Index in CLAUDE.md to note scope: `openpyxl-patterns.md` → `budget_planner/{builders,styles,workbook}.py`

**Verify**: New session, read `config.py` — openpyxl-patterns should NOT load. Read `plan.py` — it SHOULD load.

---

## Improvement 5: Split `dev-workflow.md` into core + situational (~600 tokens on routine sessions)

**Why**: At 4,998 bytes (~1,249 tokens), this is the largest rule file. It contains a mix of always-relevant coding rules and situational checklists (code review grep-cascades, /impag planning protocol, worktree sync, MCP tools) that only matter during specific workflows.

**Keep in `.claude/rules/dev-workflow.md`** (core, ~2,200 bytes, gated to `**/*.py`):
- Build + Test (lines 7-11)
- Linting (lines 14-16)
- Formatting (lines 40-41)
- Testing (lines 51-53)
- Invariants (lines 55-57)

**Move to `docs/dev-workflow-review.md`** (situational, ~2,800 bytes, on-demand):
- Code Review section (lines 18-24) — the grep-cascade safety checks
- Planning & Implementation (lines 26-31)
- Worktrees (lines 33-38)
- MCP Tools section (lines 43-49)

**Update CLAUDE.md**:
- Add trigger row: `| Column/row shift, code review, /impag, worktrees, MCP inspect | docs/dev-workflow-review.md | Grep-cascade checks, planning protocol, worktree sync, MCP tools |`
- Update existing trigger rows that reference `.claude/rules/dev-workflow.md § MCP Tools` to point to `docs/dev-workflow-review.md § MCP Tools`
- Update existing trigger row for "Config column layout" to point to `docs/dev-workflow-review.md`

**Verify**: New session, read any `.py` file — only the slim core rules load. Ask about column shifts or MCP tools — Claude should find the situational doc via trigger table.

---

## Combined Impact

| # | Change | Tokens saved | Effort |
|---|--------|-------------|--------|
| 1 | `.claudeignore` | 500-22K (prevention) | Trivial |
| 2 | Remove meta-docs | ~485 (every session) | Low |
| 3 | common-tasks → reference doc | ~366 (most sessions) | Low |
| 4 | Narrow openpyxl-patterns paths | ~813 (30-50% of sessions) | Low |
| 5 | Split dev-workflow | ~600 (routine sessions) | Medium |

**Routine coding session touching `config.py`**: 4,475 → ~2,212 tokens (~51% reduction)
**Builder session touching `plan.py`**: 4,475 → ~3,625 tokens (~19% reduction, openpyxl rules still load)

## Critical Files

- `~/projects/budget-planner/.claudeignore` (create)
- `~/projects/budget-planner/CLAUDE.md` (edit)
- `~/.claude/CLAUDE.md` (edit)
- `~/projects/budget-planner/.claude/rules/common-tasks.md` (move to docs/)
- `~/projects/budget-planner/.claude/rules/openpyxl-patterns.md` (edit frontmatter)
- `~/projects/budget-planner/.claude/rules/dev-workflow.md` (split)
- `~/projects/budget-planner/docs/common-tasks.md` (create from move)
- `~/projects/budget-planner/docs/dev-workflow-review.md` (create from split)

## Verification

After all 5 changes:
1. Start new session, run `Glob("output/**")` — should return nothing (.claudeignore working)
2. Check always-loaded context — should see only: project CLAUDE.md (~67 lines), global CLAUDE.md (~44 lines)
3. Touch `config.py` — should load only `dev-workflow.md` (core, ~30 lines)
4. Touch `plan.py` — should also load `openpyxl-patterns.md` (22 lines)
5. Ask "how do I add a WP?" — should find `docs/common-tasks.md` via trigger table
6. Ask "what MCP tools are available?" — should find `docs/dev-workflow-review.md` via trigger table
7. Run `pytest tests/ -q` — all tests should still pass (no code changes)
