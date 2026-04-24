# Runbook

Operational reference for the project: commands, common data tasks, ops / cloud env, and dev-workflow (code review, /impag, worktrees, MCP tools).

## Commands

### Dev environment setup

```bash
uv venv && uv pip install -e ".[dev]"
```

### Generate a workbook

```bash
# Empty workbook from config.py seeds
python -m budget_planner init

# Demo workbook with Capacity + Actuals data
python -m budget_planner init --demo

# Rebuild from existing workbook (preserves assignments + Actuals EUR)
python -m budget_planner rebuild
```

After any generation, run recalc + tests — see CLAUDE.md Critical Rule #2.

### Scenario comparison

```bash
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx
```

Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.

### Recalc + test

```bash
python scripts/recalc.py output/budget_plan.xlsx
pytest tests/ -v
```

## Common Tasks

### "Replace example WPs with real data"
Edit the `tbl_WorkPackages` table on the Config sheet in Excel. Run `python -m budget_planner rebuild` + pytest.

### "Replace example team with real data"
Edit the TEAM ROSTER section on the Config sheet (Person + Max FTE blue cells). Run `python -m budget_planner rebuild` + pytest.

### "Add a new work package"
Add a row to `tbl_WorkPackages` on the Config sheet in Excel. One row per WP — populate id/name/FTE/rate plus Start/End/Type (cols H/I/J, blue inputs; Type from `{flat, front-loaded, bell, back-loaded}`). Run `python -m budget_planner rebuild` + pytest.

### "Add a team member"
Add a row to the TEAM ROSTER section on the Config sheet. Run `python -m budget_planner rebuild` + pytest.

### "Add a new column/metric to Summary"
Modify `budget_planner/summary.py`. If the change introduces a new cross-sheet invariant or trust boundary, also update `docs/design-invariants.md`; if it shifts row/column layout, reflect that in `budget_planner/layout.py` + `docs/internals.md`. Run + pytest.

### "Import actuals from timesheet"
Not applicable — the Actuals sheet tracks EUR costs from the accountant, not FTE from timesheets.

### "Add actual costs"
Enter EUR values in the Actuals sheet blue input cells (one row per WP, monthly columns).

### "Verify workbook after generation"
Run pytest — it covers structural validation (sheets, named ranges, formulas, fonts). Use `openpyxl-read` MCP only for targeted follow-up questions (e.g., `get_cell_details` to check a specific formula or font color, `get_list_of_sheets` to confirm dimensions). Don't MCP-scan the whole workbook as a routine step.

### "Change the uniform vacation month"
Edit the `Vacation Month` blue input on Config — three-letter abbreviation (`"Jan"`…`"Dec"`). Run `python -m budget_planner rebuild`. The Pool formula on Timelog Plan PER-PERSON SUMMARY re-derives and the grid re-locks the new month's columns. Default is `"Jul"`.

### "Add a vacation exception for a specific person"
**Not supported in the current model.** Vacation is uniform across the team via the `vacation_month` scalar. Per-(person, year) overrides are on the roadmap as Option A (see `docs/design-invariants.md` § Timelog Plan — auditor-side invariants) — a future Config `Vacation Schedule` table keyed by person × year will layer on additively without moving Timelog Plan cells.

## Ops Quickref (AI-facing)

Short pointers for agents. Full human runbooks live under `docs/human/` and are `.claudeignore`'d — load only when explicitly asked.

### LibreOffice is a hard dependency

`soffice` is the formula recalc engine. Without it, ~40% of the test suite skips silently. Install with `sudo apt-get install -y libreoffice-calc`.

### Common failure modes

| Symptom | Fix |
|---------|-----|
| Tests report "skipped"; `demo_recalc` fixture unavailable | Install LibreOffice. |
| `recalc.py` succeeds but `data_only=True` cells are `None` | Delete `~/.config/libreoffice/4/user/basic/Standard/`, retry. |
| `recalc.py` exits 124 (timeout) | Kill stale soffice + remove lock: `pkill -f soffice && rm -f ~/.config/libreoffice/.~lock.*`, retry. |
| LO crashes with `--headless` in a container | `unset DISPLAY`, retry. |

### Debugging against the user's real workbook

Never open the user's confidential `output/*.xlsx` file. Generate a shape-preserving synthetic via `python -m scripts.make_reproducer <path>` — outputs `tmp/synthetic.xlsx` + `tmp/fp.json` with deterministic fake values (`Person_01`, `WP_A`). The `tests/test_privacy_canary.py` canary guards the pipeline against cell-value leaks. Ask the user for the synthetic path; do not ask for the real file.

Full runbooks (pre-flight script, macro-dir init, per-failure XML probes, reproducer workflow detail) live in `docs/human/` — ignored by default.

## Dev Workflow — Review, Planning, Worktrees, MCP

Situational workflows loaded on demand. For always-on rules (build/test, linting, formatting, testing, invariants) see `.claude/rules/dev-workflow.md`.

### Code Review
Before executing a plan, move completed/superseded plans to docs/plans/old_already_implemented/.
When a plan claims a module is "unchanged", grep its imports to verify no coupling to changed modules.
When inserting/removing a column on any sheet, grep ALL formula references (`Config!{letter}`, `Plan!{letter}`, etc.) across every builder AND test file. Column shifts cascade silently — a test asserting the old column letter will mask the bug.
When inserting/removing a Config param row, grep ALL test files for hardcoded row numbers (e.g. `cell(4, 2)`, `wp_first_row = 8`, merge span tuples). Row shifts cascade into tests the plan doesn't list.
When changing error message format in reader/builders, grep test files for the old message pattern (`pytest.raises(match=` strings). Semantic changes in validation produce silent test failures.
When changing a Config cell from input (blue) to formula (black), grep all consumers of the corresponding reader field (e.g. `params["rate"]`) — validators, builders, demo seeder, and tests may assume a numeric type.
Backward-compat fallbacks must trigger on *marker presence* (section header, version field, sheet name), not on an empty reader result. Empty-result triggers conflate "user cleared" with "marker missing" — on rebuild, the old source silently restores state the user intentionally deleted.
When splitting a doc or rule file, grep the source for in-file anchors (`§ X`, `below`, `above`, `see § `) before and after — split moves break them silently. Also report the split to the user (what moved where, which trigger reloads it), not just "task complete".
When resuming a plan from memory, verify actual file state before acting — memory snapshots go stale between sessions and tasks marked "pending" may already be done.

### Planning & Implementation (/impag)
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

### Worktrees
Before starting work in a worktree:
1. `git fetch origin main`
2. Check divergence: `git log --oneline HEAD..origin/main | wc -l`
3. If >0 commits behind, `git merge origin/main` before any edits
4. If >5 commits behind, grep main for the planned change before implementing — another instance may have already done it

### MCP Tools for openpyxl-read / mcp-excel
Registered via `claude mcp add` (stored in `~/.claude.json`). Verify with `claude mcp list`. Worktree-scoped — re-register in main repo if needed.

- **openpyxl-read** (primary): raw formulas + formatting. `get_list_of_sheets`, `get_cell_value`, `get_values_of_cell_range`, `get_cell_details` (font/fill/protection), `get_content_of_cell_list`, `search_in_cell_range`. Use for: formula text, font color checks, protection flags.
- **mcp-excel** (secondary): pandas-based, returns computed values. `read_excel`, `data_summary`, `analyze_excel`, `filter_excel`, `pivot_table`. Use for: numeric sanity checks on tabular sheets. Unreliable on merged-header sheets (Summary, Trends) — pandas reads them as "Unnamed" columns.

**When to use**: Only to answer a specific question about the actual workbook output (formula text, font color, computed value). One or two targeted calls. Don't MCP-scan routinely — pytest covers structural validation. If the answer is in the code or tests, read those instead.
