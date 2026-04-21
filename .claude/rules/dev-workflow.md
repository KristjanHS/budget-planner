---
paths: "**/*.py"
---
# Dev Workflow (core)

Situational workflows (code review grep-cascades, /impag planning, worktree sync, MCP tools) live in `docs/dev-workflow-review.md` — load on demand via CLAUDE.md trigger table.

## Build + Test
LibreOffice is a required dependency — install it, don't skip recalc tests. See `docs/ops-quickref.md` (quick summary) or `docs/human/cloud-env-setup.md` (full runbook) for cloud/container environments.
When generating output workbooks, check what already exists — only generate missing or explicitly requested files. Don't overwrite existing workbooks unnecessarily.
For Excel-specific bugs (implicit intersection, formula eval): generate file → user verifies in Excel → then run pytest. openpyxl/LibreOffice mask Excel-specific issues.
When designing formulas, probe a generated workbook with `python3 -c` to verify cell value types (text vs number), ref existence, and static-vs-cumulative semantics before committing the design.
When tests return None from `data_only=True`: verify recalc actually ran (`<v>` elements in XML), don't assume openpyxl/LO interop failure. No `<v>` = no calc engine ran.
**MCP-assisted verification**: use `openpyxl-read` MCP only to answer a specific question about the generated workbook; don't routinely MCP-scan — pytest covers structural validation. See `docs/dev-workflow-review.md` § MCP Tools.
If a single test fails unexpectedly on an otherwise-passing suite, re-run after `uv pip install -e .` before investigating — stale editable installs cause phantom failures.
Pass `timeout: 300000` up front for full `pytest tests/ -q` runs; the suite exceeds Bash's default. Targeted sub-suites don't need the override.
Pytest runs mutate `output/*.xlsx` (openpyxl/LO zip metadata). Always include these in commits — user treats them as tracked artifacts.
Always include editor-config dirs (`.codex/`, `.cursor/`, `.vscode/`) in commits when present.

## Linting
Run `pyright` (full project) before committing. Fix all errors.
Subagents create dead code — verify after each task.
After adding any `# pyright: ignore[reportMissingImports]` suppressions, rerun `uv run pyright` before declaring clean — the next error tier is often hidden behind the import failure.

## Formatting
Preserve all existing cell formatting (fonts, fills, table styles, column widths) unless the user explicitly asks for a change.

## Invariants (extends CLAUDE.md Critical Rules)
- Rule 6 rationale: new cross-sheet constraints, trust boundaries, and derived-metric rules are specs future agents need in loaded context — not implementation details discoverable from diff reading. Structural facts (rows/columns/freeze panes) stay in `budget_planner/layout.py`.
- On rebuild, all config data (including person-WP assignments and external commitments) comes from the existing workbook's Config sheet, never from config.py. Builders receive config as function parameters.
- When applying a code-review fix that references an existing import or helper, grep the target file first — reviewer context can be stale and a confident-sounding "already imported" claim may not be true.
