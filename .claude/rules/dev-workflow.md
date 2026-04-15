---
paths: "**/*.py"
---
# Dev Workflow (core)

Situational workflows (code review grep-cascades, /impag planning, worktree sync, MCP tools) live in `docs/dev-workflow-review.md` — load on demand via CLAUDE.md trigger table.

## Build + Test
LibreOffice is a required dependency — install it, don't skip recalc tests. See `docs/cloud-env-setup.md` for cloud/container environments.
When generating output workbooks, check what already exists — only generate missing or explicitly requested files. Don't overwrite existing workbooks unnecessarily.
For Excel-specific bugs (implicit intersection, formula eval): generate file → user verifies in Excel → then run pytest. openpyxl/LibreOffice mask Excel-specific issues.
When designing formulas, probe a generated workbook with `python3 -c` to verify cell value types (text vs number), ref existence, and static-vs-cumulative semantics before committing the design.
When tests return None from `data_only=True`: verify recalc actually ran (`<v>` elements in XML), don't assume openpyxl/LO interop failure. No `<v>` = no calc engine ran.
**MCP-assisted verification**: Use `openpyxl-read` MCP only when you need to answer a specific question about the generated workbook (e.g., "what formula is in B4?", "is this cell blue?"). Don't routinely MCP-scan after every generation — pytest covers structural validation. See `docs/dev-workflow-review.md` § MCP Tools for tool details.
If a single test fails unexpectedly on an otherwise-passing suite, re-run after `uv pip install -e .` before investigating — stale editable installs cause phantom failures. A surprise `Uninstalled 1 package / Installed 1 package` line in the next `uv run` output is the rebuild that "fixed" it.
Pytest runs mutate `output/*.xlsx` (openpyxl/LO zip metadata). Always include these in commits — user treats them as tracked artifacts, not transient build output.
Always include editor-config dirs (`.codex/`, `.cursor/`, `.vscode/`) in commits when present — user wants these tracked.

## Linting
Run `pyright` (full project, not just changed files) before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Formatting
Preserve all existing cell formatting (fonts, fills, table styles, column widths) unless the user explicitly asks for a change.

## Testing
Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months).
Derive expected values from config imports or from the generated workbook. Tests must pass after user changes data to real values in Excel and runs rebuild.
When a refactor touches both the reader and a builder, the test must exercise the full round-trip: init → user-edit via openpyxl → `rebuild_workbook` → reopen and assert. A reader-only "does it parse?" assertion is a false green — the builder may still clobber the preserved value.

## Docs
Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context). The `docs/` root is for trigger-indexed reference docs listed in CLAUDE.md. Don't drop one-off session outputs into the reference tier.
Bucket-migration tasks (moving content out of a sheet/module): before committing, grep the old identifier across `docs/` (including `docs/plans/*.md` — non-archived plan files gate paused-work resumption and will mislead the next implementing agent), `CLAUDE.md`, and `.claude/rules/`. Additive migrations leave stale prose in downstream docs the migrator never opened.

## Invariants (extends CLAUDE.md Critical Rules)
- Before merge, record any new architectural invariant in `docs/design-invariants.md` (rationale for Rule 6: a new cross-sheet constraint, trust boundary, or derived-metric rule is a spec future agents need in loaded context — not an implementation detail discoverable from diff reading).
- On rebuild, all config data (including person-WP assignments and external commitments) comes from the existing workbook's Config sheet, never from config.py. Builders receive config as function parameters.
- When touching `workbook.py` (generation pipeline), a post-v2 follow-up requires stamping a commit identifier into a Config named cell (candidate name: `version_stamp`) so an older xlsx can be mapped back to the code that produced it — see `docs/plans/2026-04-14-hours-timelog-v2-plan.md` § Post-v2 follow-ups for the full spec. Flag if implementing.
- When applying a code-review fix that references an existing import or helper, grep the target file first — reviewer context can be stale and a confident-sounding "already imported" claim may not be true. Verify, then apply.
- CF rule-count churn is a design signal. When a single task's implementation bumps a CF-count assertion more than once, pause rather than re-fix the count — usually the task is duplicating a signal already rendered elsewhere on the sheet.
