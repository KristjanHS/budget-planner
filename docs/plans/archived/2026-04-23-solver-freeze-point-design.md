# Timelog Solver Freeze-Point Design

---

## §0 Code-review follow-ups (2026-04-23, both post-Stage-2)

**Status at archive (2026-04-23)**: §0.1 Important #1–#2 applied (docs sweep across `internals.md` / `design-invariants.md` / `architecture.md` / `CLAUDE.md`); §0.1 Nice-to-have #7 applied (dead `_REBUILD_FROM_MONTH_NAMED_RANGE` constant removed); §0.1 Nice-to-have #3, #4, #5, #6 deferred as non-blocking (accepted as-is per reviewer notes). §0.2 Important #1–#2 applied in dotfiles (`is_git_commit()` rewritten around `_GIT_COMMIT_RE` regex with `(?!-)` negative lookahead + `[^|&;]*?` guard against piped commands); §0.2 Important #3 + Nice-to-have #4–#7 deferred as advisory.

Two code reviews ran against this session's work. Findings below (retained for audit trail):

### §0.1 Stage 2 review (`a76c5fb`)

Scope: builder + reader + orchestrator + test reshape. Reviewer verdict: **no Criticals; both Important findings map to the already-deferred step-8 docs sweep; one dead-code cleanup.**

**Important**

1. **Stale docs in three living reference files.** Already known — deferred per commit message — but cataloging precisely so the next session's sweep is mechanical, not investigative:
   - `docs/design-invariants.md` lines 54, 62, 89, 100 — line 100 is the authoritative vacation-cell invariant and still reads "on rebuild with `resolve_flag=FALSE`, any preserved WP code in a vacation cell is discarded." This is a cross-sheet invariant future agents auto-load; it now actively misleads.
   - `docs/architecture.md` line 166 — still names `version_stamp (B9)` (now B10). Line 175 still describes the `timelog_resolve_flag` toggle.
   - `CLAUDE.md` reference-index row for Timelog Plan — still reads "`timelog_resolve_flag` (TRUE = re-solve + wipe edits; FALSE = preserve) controls rebuild."
2. **`docs/internals.md` Cross-Sheet Reference Map missing the new row.** `test_config_named_range_targets_match_internals_doc` parses that table; full-pytest gate fails until `` | `timelog_rebuild_from_month` | Config!$B$9 | Named range | `` is added. Per `plan-hygiene.md` docs-only stages still trigger pytest coupling — handle before the recalc + full pytest gate.

**Nice-to-have**

3. **Vacation-month exclusion in `frozen_placements` is correct and defensive.** Reviewer confirmed: the partition drops vacation-month entries so `frozen_eur_by_wp` is never charged and the pool subtraction never counts them. Only risk is a corrupted workbook with a non-empty vacation cell in preserved_grid; guard handles that by dropping from both sums and overlay. No change needed.
4. **Mixed-mode test (Case 3) weakened post-boundary assertion is acceptable.** The "solved cell is non-empty" check is narrower than a strict "differs from swap" but tolerates legitimate solver re-placement. Inline comment explains the trade-off. Leave as-is.
5. **Overshoot test doesn't probe the Buffer col D formula.** Case 4 confirms no exception + frozen cells survive, but an executable assertion that the per-WP summary's Buffer cell reads `Config!E{wp_row}` (not the narrowed `ceiling_eur`) would convert the §6 prose invariant into a regression gate. Low-leverage — the formula wiring in `_write_wp_summary_row` is small, but hardening has merit if the summary ever grows.
6. **Named-range quoting confirmed correct.** Design doc §3 says "single-quoted sheet name"; implementation uses unquoted `f"{SHEET_CONFIG}!$B${row}"`. Per `openpyxl-infra.md`, quoting is only required for sheet names with non-alphanumeric characters. `Config` is plain alphanumeric → unquoted is correct and matches every other named-range registration in `config_sheet.py`. No fix needed; the design-doc notation is imprecise (same as Stage 1 review flagged).
7. **Dead constant `_REBUILD_FROM_MONTH_NAMED_RANGE` in `budget_planner/timelog_plan.py:~120`.** Defined but never referenced — the new named range is accessed via `config_layout["param_rows"]` and registered by `config_sheet.py`. Remove the constant + its comment line. Trivial.

**Review-noted strengths**: clean elimination of the binary toggle; solver-purity contract holds (solver never sees `preserved_grid`); belt-and-suspenders overlay comment honest about its role; reader's default-safe fallback chain mirrors `_read_vacation_month_num` exactly as designed; test Case 5 synthesises a realistic pre-v2 workbook state.



### §0.2 /impag context-budget hook review (dotfiles `8e7ffe4` + `9993e91`)

Scope: `claude/.claude/hooks/impag-budget-check.py` + `settings.json` registration + SKILL.md step-4f rewrite. Reviewer verdict: **design sound, two Important bugs in the commit detector, rest advisory.**

**Important**

1. **`git -C /path commit` is silently skipped.** `is_git_commit()` substring-matches `" git commit"`, which fails when `-C /path` sits between `git` and `commit`. Worktree-adjacent /impag flows would miss wrap-up reminders. Fix: `re.search(r'\bgit\b.*\bcommit\b(?!-)', cmd)` (keeps the `--help` guard).
2. **`git commit-tree` fires spuriously.** Plumbing used by rebase/filter-branch; current substring match accepts it because `commit-tree` starts with `commit`. Same regex fix covers it via `(?!-)` negative lookahead.
3. **stdout error path not fully sealed.** If `print()` raises `BrokenPipeError` after a partial flush, the partial output still lands as injected context. Low-probability. Cleanest mitigation: write to a file sentinel instead of stdout — but that requires a different injection mechanism and is likely over-engineering.

**Nice-to-have**

4. **Token estimator is optimistic on diff-heavy sessions.** `bytes/4` over-estimates prose (safe, fires early) but can *under*-count tool-use blocks with base64/large diff bodies (2–3 bytes/token). Cheap upgrade: count `\n` × avg-line-length, or sum `content` string lengths from sampled JSONL lines. Both stay <10ms.
5. **Matcher comment missing.** `"Bash"` matcher fires on every Bash call (short-circuits internally to /impag via `is_git_commit`). Add a settings.json inline comment so a future editor doesn't "tighten" it to `"git commit"` (which matches tool names, not command strings).
6. **Thresholds not env-configurable.** `WRAP_UP_TOKENS` / `HARD_STOP_TOKENS` are hardcoded for a 200k window + 75% autocompact. If `CLAUDE_AUTOCOMPACT_PCT_OVERRIDE` ever changes downward, thresholds miscalibrate. Defer: make them `os.environ.get("IMPAG_WRAP_UP_TOKENS", 110_000)` etc. Only pays off when the override actually changes.
7. **SKILL.md fallback heuristic imprecise.** "≳80 turns with multiple large-file reads" — Claude can't reliably count its own turns. Better framed as a qualitative signal ("responses feel slow, post-compact recovery messages appear, approaching-limit notes land"). Low-leverage wording polish.
8. **stdout-as-additional-context is correct.** Matches `session-start-health.py` and `frustration-tag.py` conventions. No structured envelope needed. (Not a finding — reviewer confirmed the injection path.)

**Review-noted strengths**: fail-open contract matches other hooks; docstring states limitations honestly; threshold constants carry rationale comments; SKILL.md covers both happy-path and hook-absent fallback.

---

**Status**: DESIGN — pending user approval before implementation plan drafts.
**Created**: 2026-04-23
**Supersedes**: `timelog_resolve_flag` (binary toggle) wholesale. No backwards-compatibility shim.
**Predecessor context**: `docs/plans/2026-04-23-aiko-port-decisions.md` § Q-E — reframe of Aiko's sparse assignee-grid edits into a first-class freeze semantic. Aiko-port implementation plan blocks on this doc landing.
**Directive (from user, 2026-04-23)**: minimize complexity; no backwards portability at all.

---

## §1 Goal

Replace the binary `timelog_resolve_flag` (TRUE = re-solve everything / FALSE = preserve everything) with a single integer Config parameter `timelog_rebuild_from_month`. The integer N defines a month boundary:

- Months with 1-indexed number **< N** → **frozen**: assignee-grid cells are preserved verbatim from the rebuilding workbook. The solver does not touch them.
- Months with 1-indexed number **≥ N** → **solved**: the solver assigns cells normally, treating frozen months as already-consumed budget + capacity.

The two old `resolve_flag` states collapse into the endpoints of this range:

| Old state | Equivalent `timelog_rebuild_from_month` |
|---|---|
| `resolve_flag = TRUE` (solve all) | `N ≤ 1` (or blank → defaults to 1) |
| `resolve_flag = FALSE` (preserve all) | `N > project_months` (any integer exceeding the span; the reader still reads every cell, the solver runs on an empty remaining-span) |

No separate boolean control exists after this change.

---

## §2 Non-goals

- Not adding a per-person or per-WP freeze — the boundary is project-wide by month.
- Not freezing the WP grid (rows 6–13) or any summary band — those remain fully derived, including under freeze.
- Not changing the solver's pass structure, greedy-pass scoring, or switch-reduction pass. Freezing is expressed as input-constraint narrowing, not a new pass.
- Not adding a "dry-run" mode. Freeze applies on every rebuild; the user controls it by editing `timelog_rebuild_from_month`.
- Not preserving vacation-cell styling logic — vacation month handling remains uniform (per-person `vacation_months` set); freeze does not override vacation.

---

## §3 Invariants (carry over unchanged)

- Rule 3 font colors (blue=input / black=formula / green=cross-sheet). The new Config cell is BLUE_INPUT (user-editable integer).
- Named-range discipline: `timelog_rebuild_from_month` registers via `DefinedName.attr_text` with single-quoted sheet name.
- Solver purity: no I/O, no openpyxl dependency. Freeze is expressed at the Person/WP dataclass layer as reduced `pool`/`ceiling_eur` + a `pre_frozen_months: frozenset[int]` field on `Person`, computed by the builder before calling `solve`.
- Overshoot-as-state: if frozen months over-consume a WP's `ceiling_eur`, the solver does not raise. The Buffer +i column (D6:D13 in Aiko-port header spec) renders the negative residual. Matches existing `G6` / `H6` overshoot formulas.

---

## §4 Config param row

New row added to `param_defs` in `budget_planner/config_sheet.py:716` between `Num Vacation Months` (row 8) and `Version Stamp` (row 9):

| Row | Label | Col B cell | Format | Style |
|----:|---|---|---|---|
| 9 | `Rebuild from month #` | integer (default: `1`) | `"0"` | BLUE_INPUT |
| 10 | `Version Stamp` | (existing) | — | BLACK_FORMULA (Rule 3 carve-out) |

`ConfigParamRows` in `budget_planner/layout.py:23` gains `"Rebuild from month #": int` at the same position. `Version Stamp` bumps row 9 → 10. The `param_rows`/`param_defs` drift-guard `assert` at `config_sheet.py:746` enforces the co-edit.

**Named range**: `timelog_rebuild_from_month` pointing at `Config!$B$9`. Registered via `_upsert_defined_name(wb, "timelog_rebuild_from_month", f"'{SHEET_CONFIG}'!$B$9")` in the same style-loop that registers `vacation_month` / `num_vacation_months` (`config_sheet.py:898–906`).

**Validation**: DataValidation integer constraint on Config!B9, `min=1`, `max=999` (a project will never realistically exceed 999 months; cheap upper guard). Invalid entries show Excel's default reject dialog.

**Label-cell comment** (via `sized_comment`, matching neighbouring params' docstring convention):

> Freeze assignee-grid cells in month-columns **before** this number; re-solve from this month onward. `1` (or blank) = solve the whole project. A value larger than Project Months = preserve all existing assignments verbatim. The solver will subtract pre-N person-months from each WP's budget and each person's pool before filling the remainder — over-consumed WPs surface as negative Buffer +i in the per-WP summary, not as a solver error.

---

## §5 Reader change

`budget_planner/reader.py:1385 read_timelog_plan` return shape changes:

- **Before**: `{"grid": {(person, month): wp_code}, "resolve_flag": bool}`
- **After**: `{"grid": {(person, month): wp_code}, "rebuild_from_month": int}`

`_read_resolve_flag(wb)` is **deleted**. Its replacement `_read_rebuild_from_month(wb)` reads `Config!$B$9` via the `timelog_rebuild_from_month` named range and returns an integer; returns `1` (solve-all equivalent, safe default) on missing name, unparseable value, or destination resolution failure — mirrors `_read_vacation_month_num` shape (`reader.py:1491–1526`).

The grid-scan loop itself does **not** filter by month. The reader returns the **entire** preserved grid it finds in the workbook; filtering happens in the builder / solver input prep (single point of truth — see §6).

### Pre-v2 workbook migration

Pre-v2 workbooks (no `timelog_rebuild_from_month` name, and possibly still carrying the old `timelog_resolve_flag` name) are handled by:

- Missing `timelog_rebuild_from_month` name → reader returns `rebuild_from_month=1` (solve-all). Preserves the old `resolve_flag=TRUE` default semantics.
- **Stale `timelog_resolve_flag` name is NOT read**. It will be silently overwritten when the builder writes the new Config row and calls `_upsert_defined_name` for the replacement. The old cell at its previous coordinates (Timelog Plan row 2, col 10) becomes a derelict user-writable TRUE/FALSE with no downstream consumer. **Explicit cleanup**: the Timelog Plan builder (§6) calls `del wb.defined_names["timelog_resolve_flag"]` if present before writing its own content, and skips writing the old toggle row-2 label/value cells (`_TOGGLE_LABEL_COL = 9`, `_TOGGLE_VALUE_COL = 10` are deleted as constants).

Matches the user directive "no backwards portability" — the old name is explicitly unregistered on first rebuild and never re-read.

---

## §6 Builder change

`budget_planner/timelog_plan.py build_timelog_plan_sheet` signature change:

- **Removed kwarg**: `resolve_flag: bool = True`
- **New kwarg**: `rebuild_from_month: int = 1`
- `preserved_grid` kwarg retained, semantics unchanged (still the reader's grid dict).

Call-site orchestrator (`budget_planner/workbook.py:307–319`) replaces its `_resolve_flag = …get("resolve_flag", True)` with `_rebuild_from_month = (preserved_timelog or {}).get("rebuild_from_month", 1)` and passes the int through.

### Input narrowing (the core freeze logic)

Inside the builder, **before** calling `solve(persons, wps, windows, project_months)`:

```python
N = max(1, int(rebuild_from_month))
frozen_months = frozenset(range(1, N))  # empty set when N <= 1

# Partition preserved_grid into (frozen, post-freeze). Only frozen wins.
frozen_placements: dict[str, dict[int, str]] = {}  # person -> {month: wp}
for (pname, mi), wp_code in (preserved_grid or {}).items():
    if mi in frozen_months:
        frozen_placements.setdefault(pname, {})[mi] = wp_code
```

Then derive **solver inputs narrowed by frozen consumption**:

- **Per-WP `ceiling_eur`**: subtract € already placed in frozen months. For each WP, sum `person.rate_eur_per_pm` for every `(pname, mi, wp_code)` where `wp_code == wp.code` and `mi ∈ frozen_months`. Result = `remaining_ceiling_eur = max(0, wp.ceiling_eur - frozen_placed_eur)`. If frozen placements over-consume, `remaining_ceiling_eur = 0` (solver places nothing more for this WP; Buffer +i in the per-WP summary renders the negative residual against the **original** `ceiling_eur` — the summary band reads Config's Budget +i, not the narrowed value).
- **Per-Person `pool`**: subtract frozen-month count. `remaining_pool = max(0, person.pool - |{mi : mi ∈ frozen_months ∧ assigned}|)`. If frozen placements exceed the pool, `remaining_pool = 0` (solver places nothing more for this person; the summary's Over-pool? flag lights up — this is existing behavior, not new).
- **Solver `project_months` span**: unchanged. The solver still iterates months 1..`project_months`. Frozen months are excluded via a new `Person` field `frozen_months: frozenset[int]`, added to the `@dataclass(frozen=True)` at `timelog_solver.py:41`. Its effect inside `greedy_pass` mirrors the existing `vacation_months` check (`timelog_solver.py:245`): `if m in frozen_months[p.name]: continue`. The solver never writes to a frozen month. Switch-reduction pass (`switch_reduction_pass`) gets the same guard — no swap may touch a frozen cell on either side.

### Writing the grid

After `solve(...)` returns `assigned_map`, the builder merges:

```python
# frozen_placements wins over solver output in frozen months.
for pname, month_wp in frozen_placements.items():
    for mi, wp_code in month_wp.items():
        assigned_map.setdefault(pname, {})[mi] = wp_code
```

The `if resolve_flag: … else: assigned_map = {...}` branch at `timelog_plan.py:217–222` is deleted. The new unconditional flow is "solve, then overlay frozen_placements."

### Solver input contract (backwards-incompatible)

`solve` signature at `timelog_solver.py:385`:

- **Before**: `solve(persons, wps, person_wp_windows, project_months)`
- **After**: same signature, but `persons: list[Person]` now carry `frozen_months: frozenset[int]` (default `frozenset()`). The `wps: list[WP]` now carry `ceiling_eur` already narrowed by the builder (no solver-side subtraction). `project_months` unchanged.

Design-doc note: the alternative — push frozen accounting into the solver — was rejected because it would require the solver to know about `preserved_grid`, which currently lives only in the builder layer. Keeping frozen-accounting in the builder maintains solver purity and matches the existing pattern where `Person.pool` is pre-computed by the builder from `max_fte × project_months`.

---

## §7 Layout impact

- Timelog Plan row 2 loses cols 9–10 (the old `_TOGGLE_LABEL_COL` / `_TOGGLE_VALUE_COL`). The traffic-light labels at cols 1–8 are unchanged. The builder writes no content at cols 9–10; if a prior workbook had values there, the new rebuild leaves them untouched unless the cell happens to coincide with a new Aiko-port layout change (out of scope here).
- Deleted constants in `timelog_plan.py`:
  - `_TOGGLE_LABEL_COL = 9`
  - `_TOGGLE_VALUE_COL = 10`
  - `_TOGGLE_NAMED_RANGE = "timelog_resolve_flag"`
- Added constants: none in `timelog_plan.py` (the new Config row is addressed via `param_rows["Rebuild from month #"]`, looked up from `config_layout`).
- `layout.py` `ConfigParamRows` TypedDict extended (§4). No new Timelog-Plan row/col constants.
- `docs/design-invariants.md` invariant lines referencing `timelog_resolve_flag` need updating to name `timelog_rebuild_from_month` + the new semantic. Audit during implementation: grep for `timelog_resolve_flag` across `docs/`, `CLAUDE.md`, `.claude/rules/`.

---

## §8 Test coverage

Delete `tests/test_timelog_resolve_flag.py` entirely. Replace with `tests/test_timelog_rebuild_from_month.py` covering:

1. **N = 1** (solve-all equivalent, pre-v2 migration default): rebuild re-solves the full grid; any user edit is discarded. Exact replacement of old `test_resolve_flag_true_triggers_resolve`.
2. **N > project_months** (preserve-all equivalent): rebuild preserves every cell verbatim; solver runs on an empty remaining-span (ceiling narrowed to 0, pool narrowed to 0). Replaces old `test_resolve_flag_false_preserves_grid`.
3. **N = project_months / 2** (new mixed-mode): edit a cell at month N-1 (frozen) AND a cell at month N+1 (solved). Rebuild preserves the frozen edit, overwrites the solved edit. Verify `ceiling_eur` math: if the frozen cell contributes €X to WP1, WP1's post-rebuild remaining fill = `original_ceiling - X`.
4. **Overshoot-under-freeze**: populate frozen months with assignments that exceed a WP's `ceiling_eur`. Rebuild completes without error. Per-WP summary's Buffer +i cell shows negative. No solver exception. This is the new-feature-only test; has no old-flag counterpart.
5. **Pre-v2 migration**: start from a workbook with `timelog_resolve_flag` defined name and no `timelog_rebuild_from_month`. Rebuild. Assert new name is registered, old name is removed, rebuild_from_month = 1 default took effect (re-solved).
6. **Named-range sheet-quoting**: `wb.defined_names["timelog_rebuild_from_month"].attr_text` matches `'Config'!$B$9` verbatim (openpyxl-infra rule).
7. **Config param-row drift guard**: `test_config_sheet.py` gains a row for "Rebuild from month #" in whichever sibling test asserts `param_rows` positional order.
8. **Docs-only gate**: `tests/test_config_sheet.py::test_config_named_range_targets_match_internals_doc` — update `docs/internals.md` Cross-Sheet Reference Map to include `` | `timelog_rebuild_from_month` | Config!$B$9 | Named range | `` BEFORE shipping, or this test fails.

Touched existing tests (non-exhaustive grep to run during implementation):

- `tests/test_timelog_plan.py:282` — `wb.defined_names["timelog_resolve_flag"]` assertion → `["timelog_rebuild_from_month"]`.
- `tests/test_timelog_plan_roundtrip.py:159–169` — reader return-shape assertions: `{"grid", "resolve_flag"}` → `{"grid", "rebuild_from_month"}`; bool → int; default `True` → default `1`.
- Any test referencing `_TOGGLE_VALUE_COL` / `_TOGGLE_LABEL_COL` / `resolve_flag_cell` layout dict key — grep during implementation.

**Layout dict key rename**: the builder's return dict (`timelog_plan.py:173`) drops `resolve_flag_cell` and adds `rebuild_from_month_cell` (the `Config!$B$9` coordinate as a string). Downstream consumers of `resolve_flag_cell` — grep `.get("resolve_flag_cell"` and `["resolve_flag_cell"]` across `budget_planner/**` and `tests/**` before the sweep.

---

## §9 Rollout

Single PR. No phased rollout (user directive: minimize complexity). Order of edits within the PR:

1. `budget_planner/layout.py` — extend `ConfigParamRows`, bump Version Stamp row.
2. `budget_planner/config_sheet.py` — add param row, named range, label comment, data validation.
3. `budget_planner/timelog_solver.py` — add `Person.frozen_months` field + `greedy_pass` / `switch_reduction_pass` guards.
4. `budget_planner/timelog_plan.py` — drop `resolve_flag` kwarg, add `rebuild_from_month`, compute `frozen_placements`, narrow `ceiling_eur` / `pool`, overlay frozen into `assigned_map`, remove toggle-row-2 writes + old name `_upsert_defined_name`, unregister old name if present.
5. `budget_planner/reader.py` — rename `_read_resolve_flag` → `_read_rebuild_from_month`; update `read_timelog_plan` return shape; keep unconditional full-grid read.
6. `budget_planner/workbook.py:307–319` — propagate new kwarg name.
7. Tests: delete `test_timelog_resolve_flag.py`, add `test_timelog_rebuild_from_month.py`, update touched siblings.
8. Docs: `docs/internals.md` Cross-Sheet Reference Map, `docs/design-invariants.md` (drop `timelog_resolve_flag` references, add the new semantic + overshoot-as-state note), `CLAUDE.md` reference-index row for Timelog Plan (swap `timelog_resolve_flag` → `timelog_rebuild_from_month`, update linked plan-doc pointer).
9. `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -v` (per CLAUDE.md Critical Rule 2).

### Commit-message skeleton

```
feat(timelog): replace resolve_flag with rebuild_from_month freeze-point

Collapses the binary toggle (TRUE=solve-all / FALSE=preserve-all) into a
single integer Config param. N < current freezes months 1..N-1 verbatim
on rebuild; solver re-runs for months N..end against ceilings and pools
narrowed by pre-N consumption. Overshoot surfaces as negative Buffer +i
in the per-WP summary, matching existing out-of-window semantics.

No backwards-compatibility shim: old `timelog_resolve_flag` named range
is unregistered on first rebuild; old toggle cell at row 2 cols 9-10 is
left as derelict workbook cruft.
```

---

## §10 Decisions log

| # | Decision | Rationale |
|---|---|---|
| D1 | Single integer param, no separate boolean | User directive: minimize complexity. Integer subsumes both endpoints of the old binary flag |
| D2 | Freeze scope = assignee grid rows 21–28 only | User-confirmed. WP grid rows 6–13 are already formula-derived; freezing derived rows is nonsensical |
| D3 | Overshoot surfaces as negative Buffer +i, not solver error | User-confirmed. Matches existing `G6` / `H6` overshoot-as-state pattern from Aiko-port spec |
| D4 | No backwards-compatibility shim for `timelog_resolve_flag` | User directive. Old name actively unregistered; old toggle cells derelict |
| D5 | Reader returns full grid unconditionally; filter in builder | Single-point-of-truth: both freeze subtraction and overlay happen in the builder. Keeps reader dumb |
| D6 | Frozen accounting in builder, not solver | Preserves solver purity (no I/O, no openpyxl dep). Matches existing pattern for `Person.pool` pre-computation |
| D7 | `Person.frozen_months: frozenset[int]` field mirrors `vacation_months` | Identical guard shape in both solver passes; minimal code shape change |
| D8 | Default N = 1 (not 0, not blank-as-sentinel) | 1 is the first project month; "rebuild from month 1" is unambiguous plain-English. Blank or 0 coerces to 1 at the reader layer via `max(1, int(...))` |
| D9 | Drop `_toggle_*` row-2 cells entirely; no replacement on Timelog Plan sheet | The param lives on Config now (all project-wide params live on Config). Puts the control where the user already edits project params |

---

## §11 Open items (non-blocking, resolve during implementation)

- **Config!B9 cell comment wording**: draft in §4 is first-pass; refine during the config-sheet.py edit.
- **Data-validation error message**: Excel's default reject dialog is fine; custom `errorTitle`/`error` text is low-leverage polish.
- **`docs/internals.md` Cross-Sheet Reference Map row text**: matches existing rows' formatting; trivial.
