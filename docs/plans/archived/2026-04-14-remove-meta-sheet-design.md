# Remove `_META` Sheet — Design

**Date:** 2026-04-14
**Status:** Design (awaiting implementation plan)

## 1. Goal and Rationale

Delete the `_META` worksheet and its 949-line builder (`budget_planner/meta.py`). Migrate the content worth keeping to project documentation under `docs/`. Preserve the authoring workflow by redirecting the "keep docs current" rule from `_META` to a new `docs/design-invariants.md`.

### Why remove it

- **No runtime consumer.** Nothing in `budget_planner/` reads the `_META` sheet. Only `tests/_shape_invariants.py` touches it, and only as an existence/shape smoke-check.
- **The intended consumer (future AI agents) reads `docs/*.md` instead.** `CLAUDE.md`'s trigger table already routes Claude to markdown under `docs/` every conversation. A doc tab inside the xlsx is a second, weaker channel.
- **High maintenance tax.** Critical Rule 6 requires updating `_META` on every structural change. Recent commits (`fix(v2-stage1-review): resolve residual <row> placeholders in _META`, `docs(v2-T3): document two-cadence reality… in _META`) show the rule firing repeatedly, each time patching hand-formatted prose against live row numbers and section layouts.
- **Duplication risk.** Most of `_META`'s content already exists in `docs/architecture.md`, `docs/internals.md`, `docs/directory-map.md`, and `layout.py`. Keeping two copies invites drift.

### What must survive

The one piece of `_META` content with **no current home** is the two-cadence design doctrine: the trust-boundary framing for hours vs EUR actuals and the "no derived view reads a single-month cell as authoritative" principle. That content lands in a new `docs/design-invariants.md`.

### What gets dropped outright

- Live row/column numbers embedded in prose (`layout.py` is the source of truth).
- Rebuild warnings appended at runtime (replaced by assertion-level failures; see §3).

## 2. Content Migration Plan

Six content buckets, each with a destination or a justified drop.

### Bucket 1 — Sheet index (one-liner purpose per tab)

- **Destination:** `docs/directory-map.md` (already documents "which file does what"; add a parallel "which sheet does what" subsection) and `docs/architecture.md` (already has the sheet list in its Mermaid diagrams).
- **Action:** Reconcile the `_META` sheet index against existing copy; promote any line that adds information (e.g. tab color, cross-sheet data-flow arrows) into `architecture.md` under a new "Sheet Responsibilities" section.

### Bucket 2 — Design-time assumptions (two-cadence, trust boundaries, derived-metric rules)

- **Destination:** **new `docs/design-invariants.md`** (verbatim migration, light edit for markdown).
- **Adds trigger-table row in `CLAUDE.md`:** "cross-sheet invariants, actuals trust boundaries, derived-metric rules → `docs/design-invariants.md`".
- This is the only bucket with no existing home and the core thing worth saving.

### Bucket 3 — Forward-notes (mid-rollout stage markers)

- **Destination:** already in `docs/plans/2026-04-14-hours-timelog-v2-design.md` and `docs/plans/2026-04-14-hours-timelog-v2-plan.md`.
- **Action:** Drop from `_META` migration entirely. When a stage lands, the plan doc archives; `design-invariants.md` gains any invariant that outlived the rollout.

### Bucket 4 — Live row/column numbers (cfg_wp_section, plan_first_wp, etc.)

- **Destination:** **none.** `budget_planner/layout.py` is the authoritative source; embedding numbers in prose creates a second copy that drifts.
- **Action:** Drop.

### Bucket 5 — Section-by-section sheet structure (freeze panes, fills, column formulas)

- **Destination:** `docs/internals.md` already has "Sheet Builders" and "Layout and Styling" sections. Promote any structural fact that isn't already there — preferring prose that describes *contract* ("Capacity's per-person rows are ordered by Config roster") over prose that restates cell addresses.
- **Action:** Audit `_META`'s section-by-section text against `internals.md`; migrate gaps; drop duplicates.

### Bucket 6 — Rebuild warnings (runtime-appended)

- **Destination:** **none.** Per §1 decision: drift during rebuild becomes a loud failure, not an advisory sheet annotation. See §3.

### Summary table

| Bucket | Destination | Migration effort |
|---|---|---|
| 1. Sheet index | `directory-map.md` + `architecture.md` (reconcile) | Low |
| 2. Design invariants | **new `docs/design-invariants.md`** | Medium (verbatim + markdown polish) |
| 3. Forward-notes | already in plan docs — drop | None |
| 4. Live row numbers | drop (`layout.py` is truth) | None |
| 5. Sheet structure | `internals.md` (reconcile) | Medium |
| 6. Rebuild warnings | drop → assertion in rebuild code | Low |

## 3. Code Removal Scope

### Files deleted

- `budget_planner/meta.py` (949 lines) — entire module.

### Files modified

**`budget_planner/workbook.py`** — the only runtime caller of `build_meta_sheet`:

- Remove `from budget_planner.meta import build_meta_sheet` (line 43).
- Remove Stage 9 `build_meta_sheet(...)` block (lines 353–367).
- Remove `wb.move_sheet(SHEET_META, offset=0)` (line 376) and update the tab-order comment above it.
- Rename the workbook's initial sheet away from `_META`. Two places create the initial sheet with title `_META` (lines 452–454 in init, 683–685 in rebuild). Retitle to a different leading sheet — candidate: rename the initial sheet to `Config` directly (init already builds Config first anyway), or keep a neutral `Sheet1` and delete it after the real sheets are built. **Decision:** replace initial-sheet retitle with a deferred-rename pattern — build everything on a throwaway `Sheet1`, then `wb.remove(wb["Sheet1"])` once the first real sheet exists. This avoids a second "ghost sheet" design decision.
- Update the module-level data-flow comment (line 9) to drop `_META`.
- Update the `_build_all_sheets` docstring (line 18) to drop `_META` from the "consumers of every prior layout" note.

**`budget_planner/layout.py`** — remove `SHEET_META = "_META"` constant (line 35). Confirm no remaining importers.

**Rebuild-warning handling** (addressing Bucket 6 from §2):

- `_validate_config` currently returns `(errors, warnings)`; `errors` already abort rebuild (line 678). `warnings` were routed to `_META` for human review.
- New behavior: warnings print to stderr at rebuild time (`workbook.py` already prints them on line 678 — keep that), and the `rebuild_warnings=` parameter on `_build_all_sheets` is deleted.
- Orphan/drift checks that *should* be fatal (e.g. "WP removed but Actuals still reference it") become `errors`, not `warnings`. Audit each `warnings.append(...)` site (lines 149–168, 636–668) and decide per-case. That audit is a separate task in the implementation plan — **not** folded into the _META removal task.

### Call-site confirmation

`meta_ws`, `SHEET_META`, and `build_meta_sheet` appear **only** in `budget_planner/workbook.py`, `budget_planner/layout.py`, `budget_planner/meta.py`, and tests (§5). No other production code touches the sheet.

## 4. Documentation & Rule Updates

The authoring workflow must survive the sheet deletion. Every cue that currently points at `_META` is redirected at `docs/design-invariants.md` (or removed when the cue is a live-row-number update that no longer applies).

### New file: `docs/design-invariants.md`

- **Seed content:** verbatim migration of `_META`'s "DESIGN-TIME ASSUMPTIONS: TWO-CADENCE ACTUALS" section, lightly edited for markdown.
- **Structure:**
  1. Two-cadence actuals (hours vs EUR) — trust boundaries, derived-boundary formula, single-month-cell design principle.
  2. (Future) Font color convention (Critical Rule 3 pointer).
  3. (Future) Named-range contract (`project_start`, `fte_rate`, `hours_actuals_through_month`, etc. — what each guarantees).
  4. (Future) Reader/builder round-trip rule ("readers use named ranges; tolerate internal blanks").
- Future-numbered sections are **not** written in this task — only the two-cadence section lands. The heading scheme is set up so later invariants accrete without re-architecting.

### `CLAUDE.md` — two edits

**1. Trigger-table row.** Line 17 currently reads:

```
| Config/Capacity structure changes | `budget_planner/meta.py` | Critical Rule 6: _META must reflect current sheet structure |
```

Replaced by two rows — one for the invariants doc, one retargeting structural-change docs:

```
| Cross-sheet invariant, trust boundary, derived-metric rule | `docs/design-invariants.md` | Two-cadence actuals, single-month-cell rule; update before merge |
| Workbook structure change (new sheet, new column, row shift) | `docs/internals.md` + `budget_planner/layout.py` | Structure lives in code + internals.md; no prose duplicate |
```

**2. Critical Rule 6.** Line 61 replaced:

```
6. **New architectural invariants must be recorded in `docs/design-invariants.md`** before merge. A new cross-sheet constraint, trust boundary, or derived-metric rule is a spec, not an implementation detail — future agents need it in loaded context.
```

Note the narrower scope: structural facts (rows/columns/freeze panes) no longer trigger this rule — `layout.py` and `internals.md` cover them. Only *invariants* do.

### `.claude/rules/dev-workflow.md` — one deletion

Line 35 (`"When updating meta.py, verify text against the actual builder…"`) is deleted. `meta.py` no longer exists.

### `.claude/rules/openpyxl-patterns.md` — one deletion

Line 37 (`"…when writing long documentation blocks (e.g. _META), split at semantic boundaries… see meta.py::_split_meta_for_cells"`) is deleted. The 32 767-char cell limit is a real openpyxl fact, but with no `_META` there's no caller writing doc-sized cells; rule without callers becomes noise.

### Memory update

The `project_version_stamp.md` auto-memory currently reads "every init+rebuild must stamp code version into _META…not yet implemented". Reframe it: the feature is still wanted, but the landing spot becomes a named cell on `Config` (e.g. `version_stamp`) and the stamped value must be a **commit identifier** (`git describe --always --dirty` or similar) — not a monotonic counter. The point is "which code produced this workbook" so an older xlsx can be mapped back to the exact commit for debugging. Implementation itself is a separate follow-up task.

## 5. Test Updates

Ten touch-points across eight test files. Three categories: **delete** (tests of removed code), **unwire** (tests that used `_META` as a sanity marker), and **adjust** (tests that hand-roll a workbook and currently set `active_sheet.title = "_META"` to mimic the init state).

### Delete outright

- **`tests/_shape_invariants.py`** — entire file. Only contents are `assert_meta_matches_live`; with `_META` gone, the helper has no subject.
- **`tests/test_workbook.py::test_meta_has_modification_guide`** (lines 123–136) — asserts `"MODIFICATION GUIDE"` appears in the sheet; tests deleted content.
- **`tests/test_workbook.py`** T5/T6 fixtures (lines 309, 332–333, 344–360) — spy on `build_meta_sheet` and test its partial-data refusal. Both rationales vanish with the function.

### Unwire (remove `_META`-specific assertions; keep the rest of the test)

- **`tests/test_rebuild_shape_drift.py`** — remove `from tests._shape_invariants import assert_meta_matches_live` and the two call sites (lines 16, 113, 243). The rest of each test still validates shape drift via reader round-trips; the `_META` check was redundant belt-and-suspenders.
- **`tests/test_rebuild_shapes.py`** — same pattern (lines 29, 111–117). Verify at least one non-`_META` assertion remains in each affected test; if a test becomes empty, delete it.
- **`tests/test_rebuild_e2e.py:653`** — sheetnames tab-order list currently includes `"_META"` as the first entry. Remove that entry; update any hardcoded sheet-count that follows.
- **`tests/test_workbook.py:28`** — same sheetnames list; same edit.

### Adjust (hand-rolled workbook fixtures affected by the init-sheet rename from §3)

These tests create a workbook via `openpyxl.Workbook()` and rename the default sheet to `_META` to mimic what `init_workbook`/`rebuild_workbook` do before handing off to `_build_all_sheets`. With the throwaway-sheet pattern from §3 the initial sheet's name no longer matters.

- **`tests/test_gapped_wp.py:61`** — drop the `meta.title = "_META"` line or rename to match whatever the new init pattern names the throwaway.
- **`tests/test_capacity.py:40–42`** — same.
- **`tests/test_trends.py:409, 422–424`** — uses `SHEET_META` constant. After §3 the constant is gone; switch to whatever `_build_all_sheets` now expects (likely: build the fixtures through the public `init_workbook` path instead of hand-rolling, which is the cleaner fix).

### Verification gate

After the edits:

- `pytest tests/ -q` — all green. Expect a ~3–8 test drop in total count (deletions outnumber adds).
- `pyright` — clean (removing `SHEET_META` import from `test_trends.py` resolves the obvious break).
- Grep `_META|meta_ws|build_meta_sheet|SHEET_META|assert_meta_matches_live|meta\.py` across the repo — zero hits outside archived `docs/plans/archived/`.

## 6. Migration Order and Verification

Six phases, each with a gate before the next starts. Order matters: docs land first (they don't depend on anything), `CLAUDE.md` updates next (they take the new docs "live" in the trigger table), then code removal, then test cleanup. Tests **must** be cleaned after code removal, not before — the tests that hand-roll fixtures need the new init-sheet pattern to exist.

### Phase 1 — Additive docs (zero risk)

- Create `docs/design-invariants.md` with the two-cadence section migrated from `_META` (§4).
- Reconcile Bucket 1 (sheet index) into `docs/directory-map.md` and `docs/architecture.md` (§2).
- Reconcile Bucket 5 (sheet structure) into `docs/internals.md` (§2).

**Gate:** `git diff docs/` reviewed; no code or test edits yet; `pytest` still green (nothing changed).

### Phase 2 — Routing update

- Edit `CLAUDE.md` trigger table and Critical Rule 6 (§4).
- Delete stale lines from `.claude/rules/dev-workflow.md` and `.claude/rules/openpyxl-patterns.md` (§4).

**Gate:** `CLAUDE.md` renders readably; trigger-table entries each point to an existing file. No runtime impact — this phase is context-layer only.

### Phase 3 — Code removal

- Delete `budget_planner/meta.py`.
- Apply the six edits to `budget_planner/workbook.py` (§3): import removal, Stage 9 removal, `move_sheet` removal, initial-sheet rename via throwaway `Sheet1`, data-flow comment and docstring cleanup.
- Delete `SHEET_META` from `budget_planner/layout.py`.
- Delete the `rebuild_warnings=` parameter from `_build_all_sheets` and its callers; keep the stderr print in `rebuild_workbook` (§3).

**Gate:** `python -m budget_planner init` succeeds; `python -m budget_planner init --demo` succeeds; open the xlsx and verify there is no `_META` tab and tab order matches the comment. `pyright` clean.

### Phase 4 — Test cleanup

Apply all edits from §5. Expected order within this phase:

1. Delete `tests/_shape_invariants.py` and the three files that import it (unwire).
2. Delete `test_meta_has_modification_guide` and T5/T6 fixtures in `test_workbook.py`.
3. Strip `_META` from the two sheetnames lists.
4. Fix the three hand-rolled-fixture tests last — they're the trickiest and benefit from the others being green first.

**Gate:** `pytest tests/ -q` green; `pyright` clean; `grep -rn '_META\|meta_ws\|build_meta_sheet\|SHEET_META\|assert_meta_matches_live' --include='*.py' .` returns zero hits.

### Phase 5 — Memory refresh

- Update `~/.claude/projects/.../memory/project_version_stamp.md` per §4 (reframe to `Config` named cell, commit-identifying value).
- Confirm no other memory files reference `_META`.

**Gate:** `grep -l '_META' ~/.claude/projects/.../memory/` returns empty.

### Phase 6 — Commit and archive

- Single commit or small commit chain (docs → routing → code → tests → memory).
- On merge, move this design file to `docs/plans/archived/` per the project convention.
- Write a companion implementation plan (`docs/plans/2026-04-14-remove-meta-sheet-plan.md`) using `superpowers:write-plan` **before** starting Phase 1 if the executor is a subagent chain; skip the plan doc if a single developer runs it interactively.

### Out of scope for this task

- Case-by-case audit of rebuild warnings → should-they-be-fatal? (Flagged in §3.)
- Implementation of the `version_stamp` named cell on Config (§4).
- Any CF/formatting changes to the xlsx.

These become follow-up tasks, tracked separately.





