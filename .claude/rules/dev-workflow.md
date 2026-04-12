---
paths: "**/*.py"
---
# Dev Workflow (core)

Situational workflows (code review grep-cascades, /impag planning, worktree sync, MCP tools) live in `docs/dev-workflow-review.md` — load on demand via CLAUDE.md trigger table.

## Build + Test
LibreOffice is a required dependency — install it, don't skip recalc tests. See `docs/cloud-env-setup.md` for cloud/container environments.
When generating output workbooks, check what already exists — only generate missing or explicitly requested files. Don't overwrite existing workbooks unnecessarily.
For Excel-specific bugs (implicit intersection, formula eval): generate file → user verifies in Excel → then run pytest. openpyxl/LibreOffice mask Excel-specific issues.
When tests return None from `data_only=True`: verify recalc actually ran (`<v>` elements in XML), don't assume openpyxl/LO interop failure. No `<v>` = no calc engine ran.
**MCP-assisted verification**: Use `openpyxl-read` MCP only when you need to answer a specific question about the generated workbook (e.g., "what formula is in B4?", "is this cell blue?"). Don't routinely MCP-scan after every generation — pytest covers structural validation. See `docs/dev-workflow-review.md` § MCP Tools for tool details.

## Linting
Run `pyright` (full project, not just changed files) before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Formatting
Preserve all existing cell formatting (fonts, fills, table styles, column widths) unless the user explicitly asks for a change.

## Testing
Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months).
Derive expected values from config imports or from the generated workbook. Tests must pass after user changes data to real values in Excel and runs rebuild.

## Invariants (extends CLAUDE.md Critical Rules)
- When updating meta.py, verify text against the actual builder (e.g., count summary columns in summary.py, not the plan)
- On rebuild, all config data (including person-WP assignments and external commitments) comes from the existing workbook's Config sheet, never from config.py. Builders receive config as function parameters.
