# Dev Workflow — Review, Planning, Worktrees, MCP

Situational workflows loaded on demand. For always-on rules (build/test, linting, formatting, testing, invariants) see `.claude/rules/dev-workflow.md`.

## Code Review
Before executing a plan, move completed/superseded plans to docs/plans/old_already_implemented/.
When a plan claims a module is "unchanged", grep its imports to verify no coupling to changed modules.
When inserting/removing a column on any sheet, grep ALL formula references (`Config!{letter}`, `Plan!{letter}`, etc.) across every builder AND test file. Column shifts cascade silently — a test asserting the old column letter will mask the bug.
When inserting/removing a Config param row, grep ALL test files for hardcoded row numbers (e.g. `cell(4, 2)`, `wp_first_row = 8`, merge span tuples). Row shifts cascade into tests the plan doesn't list.
When changing error message format in reader/builders, grep test files for the old message pattern (`pytest.raises(match=` strings). Semantic changes in validation produce silent test failures.
When changing a Config cell from input (blue) to formula (black), grep all consumers of the corresponding reader field (e.g. `params["rate"]`) — validators, builders, demo seeder, and tests may assume a numeric type.
Backward-compat fallbacks must trigger on *marker presence* (section header, version field, sheet name), not on an empty reader result. Empty-result triggers conflate "user cleared" with "marker missing" — on rebuild, the old source silently restores state the user intentionally deleted.
When splitting a doc or rule file, grep the source for in-file anchors (`§ X`, `below`, `above`, `see § `) before and after — split moves break them silently. Also report the split to the user (what moved where, which trigger reloads it), not just "task complete".
When resuming a plan from memory, verify actual file state before acting — memory snapshots go stale between sessions and tasks marked "pending" may already be done.

## Planning & Implementation (/impag)
Write plans to `docs/plans/` as dated .md files immediately — never keep plans only in conversation context (context gets compressed).
Run brainstorm → write-plan → worktree isolation for each P-item. Don't skip planning even if the spec looks detailed.
Run `superpowers:code-reviewer` after every /impag stage — not just at the end. Each stage boundary is a review checkpoint.
Fix pre-existing test failures when encountered during implementation — don't defer them.
Commit or resolve dirty files at session start before executing any tasks.
**Before starting a `/impag` default-selected plan:** the default is the most recent plan by mtime, which is often a *just-completed* plan still in `docs/plans/`. Cross-check `project_state.md` and recent `git log` to confirm it's not already landed. If it is, ask the user which queued plan to run instead.

**Before dispatching task 1, pre-flight the plan:**
1. Grep the central symbols the plan names (`build_workbook`, `DEFAULT_SEED`, etc.) to confirm they exist in the current tree. Plan text drifts as the code evolves; catching mis-references up front saves N subagent prompts downstream.
2. If the plan opens with a hypothesis-first RED test, manually reproduce the hypothesis on current code *before* committing the test file. A RED test that comes back GREEN means the bug model is wrong — stop and revise rather than pushing on.
3. Plan `build_from_fingerprint`-style tasks with "seed injection" assume an API shape — verify the real entrypoint signature in the builder module before dispatching, not after the subagent gets stuck.
4. For each task, grep the tree for the anchors it's supposed to *introduce* (new constants, error messages, helper names). Plans sometimes stay in `docs/plans/` after individual tasks land via unrelated commits. A 30-second anchor grep per task beats paying the subagent dispatch cost to learn the task is already done.

## Worktrees
Before starting work in a worktree:
1. `git fetch origin main`
2. Check divergence: `git log --oneline HEAD..origin/main | wc -l`
3. If >0 commits behind, `git merge origin/main` before any edits
4. If >5 commits behind, grep main for the planned change before implementing — another instance may have already done it

## MCP Tools (openpyxl-read + mcp-excel)
Registered via `claude mcp add` (stored in `~/.claude.json`). Verify with `claude mcp list`. Worktree-scoped — re-register in main repo if needed.

- **openpyxl-read** (primary): raw formulas + formatting. `get_list_of_sheets`, `get_cell_value`, `get_values_of_cell_range`, `get_cell_details` (font/fill/protection), `get_content_of_cell_list`, `search_in_cell_range`. Use for: formula text, font color checks, protection flags.
- **mcp-excel** (secondary): pandas-based, returns computed values. `read_excel`, `data_summary`, `analyze_excel`, `filter_excel`, `pivot_table`. Use for: numeric sanity checks on tabular sheets. Unreliable on merged-header sheets (Summary, Trends) — pandas reads them as "Unnamed" columns.

**When to use**: Only to answer a specific question about the actual workbook output (formula text, font color, computed value). One or two targeted calls. Don't MCP-scan routinely — pytest covers structural validation. If the answer is in the code or tests, read those instead.
