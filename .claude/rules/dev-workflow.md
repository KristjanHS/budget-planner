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
Pyright import-resolution errors mask downstream type errors. After adding any `# pyright: ignore[reportMissingImports]` suppressions, rerun `uv run pyright` before declaring the project clean — the next error tier is often hidden behind the import failure.

## Formatting
Preserve all existing cell formatting (fonts, fills, table styles, column widths) unless the user explicitly asks for a change.

## Testing
Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months).
Derive expected values from config imports or from the generated workbook. Tests must pass after user changes data to real values in Excel and runs rebuild.
When a refactor touches both the reader and a builder, the test must exercise the full round-trip: init → user-edit via openpyxl → `rebuild_workbook` → reopen and assert. A reader-only "does it parse?" assertion is a false green — the builder may still clobber the preserved value.
Full `pytest tests/ -q` runs ~140s, exceeding Bash's 120s default and triggering auto-backgrounding. Pass `timeout: 300000` up front. If a pytest call auto-backgrounds, wait for the completion notification — do NOT re-issue it in the foreground. Targeted sub-suites (`tests/test_foo.py`) don't need the override.

## Plan hygiene
Shipped plans left in `docs/plans/` root auto-load into future sessions and masquerade as active work. When `/impag` (or any plan-survey step) finds a plan fully shipped, `git mv` it into `docs/plans/archived/` the same session — don't just note shipped status and move on.
Before `git mv`: grep the plan path across `budget_planner/**`, `tests/**`, `docs/**`, `CLAUDE.md` and update references to `docs/plans/archived/<filename>.md`. Verify `.claudeignore` names `archived/` (this project's convention, not `archive/`).
Honor the global "Superseded docs: consume before archive" rule — pull load-bearing content into living docs (`design-invariants.md`, `internals.md`) before archival. Do NOT archive plans marked "approved but not implemented" — those are still active even if dormant.

## Docs
Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context). The `docs/` root is for trigger-indexed reference docs listed in CLAUDE.md. Don't drop one-off session outputs into the reference tier.
Bucket-migration tasks (moving content out of a sheet/module): before committing, grep the old identifier across `docs/` (including `docs/plans/*.md` — non-archived plan files gate paused-work resumption and will mislead the next implementing agent), `CLAUDE.md`, and `.claude/rules/`. Additive migrations leave stale prose in downstream docs the migrator never opened.
When a doc row claims a specific font color or formula-vs-input style for a named cell ("blue input", "black formula", "green cross-ref"), grep the builder's actual `style_cell(...)` / `font=...` call to verify. Critical Rule 3 (blue=input, black=formula, green=cross-ref) is load-bearing; mislabelling is worse than silent. Do not generalize from sibling cells — e.g. Config trust-boundary cells being blue does not imply all Config-adjacent scalars are blue.
When adding a rule example that documents a *shipped* fix (openpyxl-patterns formulas, style conventions), grep the builder's emitted form in the committed code and copy it verbatim — not the design-doc version. Design docs often carry the pre-fix form; shipping that as the rule example re-introduces the original bug as a template for the next implementer. 2026-04-19: openpyxl-patterns Last log entry was initially added with the unguarded design-doc form; code-reviewer caught that shipping it would re-introduce the LO `INDEX(range, 0)` implicit-intersection bug the fix solved.
When shifting documented row positions on a sheet (Config param rows, WP section, etc.), grep `docs/internals.md` for the old row numbers — positions live in BOTH the Per-Sheet Notes paragraph AND the Extraction Functions table. Test docstrings that cite specific rows/counts also need touching; assertion-level edits don't reach prose.

## Invariants (extends CLAUDE.md Critical Rules)
- Rule 6 rationale: new cross-sheet constraints, trust boundaries, and derived-metric rules are specs future agents need in loaded context — not implementation details discoverable from diff reading. Structural facts (rows/columns/freeze panes) stay in `layout.py` + `docs/internals.md`.
- On rebuild, all config data (including person-WP assignments and external commitments) comes from the existing workbook's Config sheet, never from config.py. Builders receive config as function parameters.
- Version stamp (shipped 1ad2021): `workbook._resolve_version_stamp()` writes `git describe --always --dirty` (fallback `"unknown"`) into Config!$B$7 (named range `version_stamp`) on every init/rebuild, so older xlsx files can be mapped back to the builder commit. Keep the single-cell placement on Config — `_META` sheet was removed 2026-04-14. Row shifted from B9 → B7 in the 2026-04-19 Config through-month cleanup.
- Rule 3 carve-out: builder-written audit metadata (e.g. `version_stamp` on Config) uses BLACK_FORMULA even though the cell value is a literal string, not a formula — signals "do-not-edit by user." Required mitigation when deviating from Rule 3 this way: a builder-code comment AND an in-cell comment that explicitly name the deviation (see `config_sheet.py` Version Stamp rows for the reference pattern).
- When applying a code-review fix that references an existing import or helper, grep the target file first — reviewer context can be stale and a confident-sounding "already imported" claim may not be true. Verify, then apply.
