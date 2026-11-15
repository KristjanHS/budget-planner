---
paths: "**/*.py"
---
# Dev Workflow (core)

Situational workflows (code review grep-cascades, /impag planning, worktree sync, MCP tools) live in `docs/spec.md` § Dev Workflow — load on demand via CLAUDE.md trigger table.

## Build + Test
LibreOffice is a required dependency — install it, don't skip recalc tests. See `docs/spec.md` § Ops Quickref (quick summary) or `docs/human/cloud-env-setup.md` (full runbook) for cloud/container environments.
When generating output workbooks, check what already exists — only generate missing or explicitly requested files. Don't overwrite existing workbooks unnecessarily.
For Excel-specific bugs (implicit intersection, formula eval): generate file → user verifies in Excel → then run pytest. openpyxl/LibreOffice mask Excel-specific issues.
When designing formulas, probe a generated workbook with `python3 -c` to verify cell value types (text vs number), ref existence, and static-vs-cumulative semantics before committing the design.
When tests return None from `data_only=True`: verify recalc actually ran (`<v>` elements in XML), don't assume openpyxl/LO interop failure. No `<v>` = no calc engine ran.
**MCP-assisted verification**: use `openpyxl-read` MCP only to answer a specific question about the generated workbook; don't routinely MCP-scan — pytest covers structural validation. See `docs/spec.md` § MCP Tools for openpyxl-read / mcp-excel.
If a single test fails unexpectedly on an otherwise-passing suite, re-run after `uv pip install -e .` before investigating — stale editable installs cause phantom failures.
Pass `timeout: 300000` up front for full `pytest tests/ -q` runs; the suite exceeds Bash's default. Targeted sub-suites don't need the override.
Pytest runs mutate `output/*.xlsx` (openpyxl/LO zip metadata). `output/` is gitignored whole (workbooks embed seed data); never `git add -f` anything under it.
Always include editor-config dirs (`.codex/`, `.cursor/`, `.vscode/`) in commits when present.

## Linting
Run `make lint` (pyright + `ruff check` + `ruff format --check`, all full-project) before committing. Fix all errors.
**Do NOT upgrade the host ruff past 0.12.9 without re-testing this.** ruff 0.16.0's formatter strips the parentheses from `except (A, B):` under this project's `line-length = 120`, emitting invalid Python — `make fmt` under it left 6 files unparseable while its own `--check` called them formatted. Reproduced on a clean HEAD file (`uvx ruff@0.16.0 format --diff budget_planner/readers/config.py`); `--isolated` (width 88) does not trigger it. 0.16 also formats Python inside Markdown code blocks, hence the `**/*.md` exclude.
**`ruff format` needs two passes here.** One pass over the builder modules leaves 7 of them still needing a reformat (the `style_cell(...).value = (...)` long-assignment shape); reproduced on ruff 0.12.9 and 0.16.0, so it is not a stale-version bug. `make fmt` therefore formats twice and asserts convergence — never conclude "formatted" from a single pass.
A write pass records its own output as formatted, so a cached `--check` straight after a cached `ruff format` reports clean while a cold check finds the residue. `make lint`/`make fmt` pass `--no-cache` (costs 0.02s); the global PostToolUse ruff hook does not, and stays honest only because the file it checks has just changed (cache miss) — it can go blind on files a cached `ruff format` wrote and nothing has touched since.
Ruff config lives in `pyproject.toml` `[tool.ruff]` (line-length 120, `select = ["E","F","W"]`, `ignore = ["E501"]`, `scripts/office` excluded + `force-exclude = true` so the PostToolUse hook's single-file invocation honours the exclude). Line width is the formatter's job; keep our own modules under 120 by hand anyway — `make fmt` won't split long comments or string literals.
Subagents create dead code — verify after each task.
After adding any `# pyright: ignore[reportMissingImports]` suppressions, rerun `uv run pyright` before declaring clean — the next error tier is often hidden behind the import failure.

## Formatting
Preserve all existing cell formatting (fonts, fills, table styles, column widths) unless the user explicitly asks for a change.

## Invariants (extends CLAUDE.md Critical Rules)
- Rule 6 rationale: new cross-sheet constraints, trust boundaries, and derived-metric rules are specs future agents need in loaded context — not implementation details discoverable from diff reading. Structural facts (rows/columns/freeze panes) stay in `budget_planner/layout.py`.
- On rebuild, all config data (including person-WP assignments and external commitments) comes from the existing workbook's Config sheet, never from config.py. Builders receive config as function parameters.
- When applying a code-review fix that references an existing import or helper, grep the target file first — reviewer context can be stale and a confident-sounding "already imported" claim may not be true.
