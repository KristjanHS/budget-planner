# WP gap-logic removal — single-range schema, merged Config section

**Status:** Validated via brainstorming 2026-04-19 (Doc B). Ready for single-PR implementation.
**Scope:** Forbid multi-row WP_SCHEDULE entries entirely. Fold `start`/`end`/`type` into the `WORK_PACKAGES` / Config WP BUDGET table. Delete `_wp_sched_months`; collapse `capacity.py`'s gap-aware filter. Legacy workbooks with multi-row schedules or the two-section Config layout get a hard-fail on rebuild with an explicit migration recipe.
**Relationship to Doc A (`docs/plans/archived/2026-04-19-timelog-plan-enhancements-design.md`):** Doc A's D5/§5 were written to tolerate this removal — the yearly-budget formula's `_wp_sched_months(wpk.code, wp_schedule)` term collapses to `end - start + 1` with no formula rewrite. Doc B and Doc A are independent; either can ship first. If Doc A ships first, Phase 4's implementation must swap `_wp_sched_months(...)` for `end - start + 1` when Doc B lands (a one-liner diff per WP term).

---

## 1. Motivation

Multi-segment WP schedules exist today as a dormant capability:

- The seed `WP_SCHEDULE` in `config.py` has exactly one WP (WP3) with two rows — a historical artefact.
- The solver, capacity distribution, ceiling formulas, and yearly-budget proportions all carry "gap-aware" code paths (`_wp_sched_months`, `valid_months = assignment_window ∩ union(segments)`, grey-fill for gap months).
- No current user of the tool uses multi-segment WPs in a live plan. Keeping the capability means every builder reads `wp_schedule` as a `list[dict]` keyed by id-with-duplicates, every sheet has "gap month" styling paths, and every test touching schedules has to reason about whether segments matter.
- Doc A surfaced this explicitly (§1 item 3, §2 non-goal, D15) — the word "gap" was becoming overloaded; `Largest gap` on PER-PERSON SUMMARY was dropped partly to avoid collision.

Forbidding multi-row entries removes ~200 lines of conditional logic across 7 builders + the reader + `config_sheet.py`, collapses a dedicated test file (`test_gapped_wp.py`, 5 tests), and eliminates one entire section from the Config sheet. The simplification is cross-cutting; Doc A's formula layer is already designed to accommodate it (D5 / §5 note).

---

## 2. Goals / Non-goals

### Goals

- **Schema merge.** Fold `start`, `end`, `type` fields from `WP_SCHEDULE` into `WORK_PACKAGES`. One list of dicts, one entry per WP, schema `{id, name, fte, rate?, start, end, type}`.
- **Config sheet merge.** Delete the `WP SCHEDULE` section. Extend the `WP BUDGET` section with three new columns (`Start`, `End`, `Type`) appended after the existing formula columns. One row per WP.
- **Builder signature sweep.** Remove the `wp_schedule` parameter from every builder that currently takes it (11 files touched). Builders read `start`/`end`/`type` off each `work_packages` entry directly.
- **Reader dual validation.** Reader raises two distinct, specific errors: (a) legacy two-section layout detected on Config, and (b) duplicate WP id within the merged section. Both name the fix in the error message.
- **Capacity filter simplification.** `capacity.py`'s gap-aware intersection collapses to single-range clipping (`wp.start <= m <= wp.end`). Grey-fill for "out-of-WP-range but in-assignment-window" months is preserved (still a real case); "gap months" as a separate styling tier disappears.
- **Helper deletion.** `_wp_sched_months` deleted. Call sites inline `wp["end"] - wp["start"] + 1`.
- **Test sweep.** `tests/test_gapped_wp.py` deleted. One replacement test in `tests/test_reader_new.py` verifying both hard-fail paths (layout + duplicate id). Other tests' incidental references to `wp_schedule` as a separate parameter updated.
- **Doc sweep.** `docs/internals.md`, `docs/design-invariants.md`, `docs/architecture.md`, `CLAUDE.md` trigger table, `docs/common-tasks.md`, `docs/grant-domain.md` — each grep'd for "WP SCHEDULE", "multi-segment", "gap-aware", "gap months"; each updated or removed.
- **Single PR.** No phased rollout. The amputation is atomic at the schema level; intermediate states ship half-migrated complexity.

### Non-goals (YAGNI)

- **Auto-upgrade of legacy workbooks.** Not worth the code — user collapses the Config sheet manually per the error message. Covered by Q1 decision.
- **Keeping `_wp_sched_months` as a trivial wrapper.** Call sites inline the one-liner. Wrapper adds indirection without readability gain once segments are gone.
- **Reordering existing `WP BUDGET` columns.** The new `Start`/`End`/`Type` cols append on the right to minimise layout shift. A cleaner "inputs left, formulas right" reordering is possible but touches every downstream cell-ref in `config_sheet.py`; deferred — if later revisited, it becomes its own layout-shift doc.
- **Collapsing `type` into a single default.** `type` is load-bearing (`capacity.py::distribute` dispatches on flat / front-loaded / bell / back-loaded).
- **Removing "out-of-window" grey-fill.** An assignment that extends past a WP's single active range still produces out-of-WP-range months that must be grey + empty. Single-range clipping preserves this.
- **Migrating `WP3`'s seed data atomically with the code change.** `config.py`'s seed `WP_SCHEDULE` has WP3 split across two rows (M1-M3, M5-M7). Under the new schema, seed data becomes `{id: "WP3", ..., start: 1, end: 7, type: "front-loaded"}` — collapsing the gap. User-facing note: this changes the semantics of the seed WP3 (M4 was previously unassignable). Flagged in the `config.py` docstring. Downstream sample workbooks regenerated.

---

## 3. Data-model change

### 3.1 `WORK_PACKAGES` schema (before → after)

**Before:**
```python
WORK_PACKAGES = [
    {"id": "WP1", "name": "Project management", "fte": 12.0},
    {"id": "WP2", "name": "Requirements",        "fte":  8.0},
    ...
]

WP_SCHEDULE = [
    {"id": "WP1", "start":  1, "end": 14, "type": "flat"},
    {"id": "WP2", "start":  1, "end": 6, "type": "front-loaded"},
    {"id": "WP3", "start":  1, "end": 3, "type": "front-loaded"},  # multi-row
    {"id": "WP3", "start": 5, "end": 7, "type": "front-loaded"},  # multi-row
    ...
]
```

**After:**
```python
WORK_PACKAGES = [
    {"id": "WP1", "name": "Project management", "fte": 12.0,
     "start":  1, "end": 14, "type": "flat"},
    {"id": "WP2", "name": "Requirements",        "fte":  8.0,
     "start":  1, "end": 18, "type": "front-loaded"},
    {"id": "WP3", "name": "Design",              "fte": 10.0,
     "start":  1, "end": 7, "type": "front-loaded"},  # was split M1-M3 + M5-M7
    ...
]
# WP_SCHEDULE removed entirely from config.py
```

### 3.2 Invariants

- **Exactly one entry per WP id.** Reader enforces; builders assume.
- **`start` and `end` are 1-indexed project months.** Unchanged from current `WP_SCHEDULE`.
- **`1 <= start <= end <= project_months`.** Reader validates — `start > end` or out-of-range → error (was implicit before; now explicit since there's nowhere else for schedule data to live).
- **`type` ∈ {"flat", "front-loaded", "bell", "back-loaded"}.** Unchanged. Dispatch point is `capacity.py::distribute`.
- **Optional per-WP `rate`.** Existing field, unchanged. Remains optional; fallback to `project_params["rate"]` unchanged.

### 3.3 Dataclass / TypedDict impact

No formal dataclass exists for WP config today (plain dicts everywhere). This doc does NOT introduce one — staying consistent with current style. Reader returns the same `list[dict]` shape, now with the extended key set.

Any typed callsites (`WP` dataclass in `timelog_plan.py::_build_wps`) already carry `active_start` / `active_end` / `ceiling_eur` as attributes; they construct from the `WORK_PACKAGES` list. The constructor switches from "look up schedule row by id" to "read `wp["start"]` / `wp["end"]` directly." Single-segment assumption matches what the dataclass already models (no list of segments, one pair of ints).

---

## 4. Config sheet layout change

### 4.1 New `WP BUDGET` section columns

**Before:**

| col | constant | meaning |
|---|---|---|
| 1 | `COL_WP_ID` | WP id |
| 2 | `COL_WP_NAME` | name |
| 3 | `COL_WP_FTE` | FTE budget (blue input) |
| 4 | `COL_WP_RATE` | rate per-WP override (blue input, optional) |
| 5 | `COL_WP_CEILING` | `=FTE×Rate` (black formula) |
| 6 | `COL_WP_ASSIGNED_PM` | SUMIFS across Plan (black formula) |
| 7 | `COL_WP_VARIANCE` | FTE − Assigned (black formula) |

**After:** cols 1-7 unchanged; three new cols appended:

| col | constant | meaning |
|---|---|---|
| 8 | `COL_WP_START` | project-month start (blue input, 1-indexed int) |
| 9 | `COL_WP_END` | project-month end (blue input, 1-indexed int) |
| 10 | `COL_WP_TYPE` | distribution curve (blue input, data-validation list) |

Updated constants in `layout.py`:

```python
COL_WP_START = 8
COL_WP_END = 9
COL_WP_TYPE = 10
WP_BUDGET_LAST_COL = 10  # was 7
```

Deleted constants in `layout.py`:

```python
# DELETE:
COL_SCHED_START = 2
COL_SCHED_END = 3
COL_SCHED_TYPE = 4
SCHED_LAST_COL = 4
```

`SECTION_HEADER_LABELS["wp_schedule"]` entry deleted (key `"wp_schedule"` removed).

### 4.2 Styling

- `Start` / `End` cells: `BLUE_INPUT` font, integer format (`0`), data-validation range `[1, project_months]` (new — previously validated in workbook validator, now also at cell level for UX).
- `Type` cell: `BLUE_INPUT` font, data-validation list `"flat,front-loaded,bell,back-loaded"` (new — previously validated in capacity.py fallback, now explicit).
- Header row: existing `BUDGET_TABLE_HEADERS` tuple gains `"Start"`, `"End"`, `"Type"` in order.

### 4.3 Named ranges

No change. `project_start`, `fte_rate`, `version_stamp` unaffected. Schedule data was never exposed via named ranges; per-WP formulas in other sheets reference `Config!<col><wp_row>` directly.

### 4.4 Section deletion

The entire `WP SCHEDULE` section (title row, header row, data rows, trailing blank) is removed from `build_config_sheet`. Rows below shift up accordingly. Section order after this change:

1. `PROJECT PARAMETERS`
2. `WP BUDGET` (now with 10 columns)
3. ~~`WP SCHEDULE`~~ (deleted)
4. `TEAM ROSTER`
5. `RATE BY GRADE`
6. `REPORTING PERIODS`
7. `PERSON-WP ASSIGNMENTS`
8. `EXTERNAL COMMITMENTS`
9. `PLANNED FTE (PERSON-MONTHS)`

Readers locate sections by header text (`_find_section_start`) — row shift is absorbed transparently for every section parser. The only parser that needs removal is `_read_wp_schedule`.

---

## 5. Reader validation

Two distinct error paths, both raised from `reader.py::read_config` (or a helper it calls). Both errors are `ValueError` subclasses; exact messages below.

### 5.1 Legacy two-section layout detection

**Trigger:** A cell on Config col A matches the literal string `"WP SCHEDULE"` (case-sensitive, post-strip).

**Check placement:** Runs early in `read_config`, before any section parsing. Implementation:

```python
for row in range(1, ws.max_row + 1):
    val = ws.cell(row=row, column=1).value
    if isinstance(val, str) and val.strip() == "WP SCHEDULE":
        raise ValueError(
            f"Legacy Config layout detected: 'WP SCHEDULE' section found at "
            f"row {row}. Multi-segment WP schedules are no longer supported.\n"
            f"Migration: delete the WP SCHEDULE section from the Config sheet; "
            f"add 'Start', 'End', 'Type' columns to the WP BUDGET section "
            f"(cols H, I, J) and populate one row per WP. "
            f"Any WP that previously had multiple active ranges must be "
            f"collapsed to a single (start, end) pair."
        )
```

### 5.2 Duplicate WP id in merged section

**Trigger:** Reader encounters two `WP BUDGET` rows with the same `id` during section parse.

**Check placement:** Inside `_read_wp_budget` (renamed from current name if needed), immediately after parsing each row.

```python
seen_ids: set[str] = set()
for row in wp_budget_rows:
    wp_id = row["id"]
    if wp_id in seen_ids:
        raise ValueError(
            f"Duplicate WP id '{wp_id}' in WP BUDGET section. "
            f"Multi-segment schedules are no longer supported — each WP must "
            f"appear exactly once. Collapse to a single (Start, End) range."
        )
    seen_ids.add(wp_id)
```

### 5.3 Existing validator update in `workbook.py`

The current validator at `workbook.py:157-166` says:

> "Duplicate WP ID in WP BUDGET table: {id}. Use one row per ID there; for multiple active periods use WP SCHEDULE (same ID, different Start/End), or split into distinct IDs (e.g. WP4a, WP4b)."

**Rewritten:**

> "Duplicate WP ID in WP BUDGET table: {id}. Each WP must appear exactly once. If you need separable phases, split into distinct IDs (e.g. WP4a, WP4b)."

The "use WP SCHEDULE" guidance is removed; the "split into distinct IDs" fallback remains (it's the legitimate escape hatch for logical phasing).

Schedule-specific validation in `workbook.py:182-184` (checking schedule ids match work_packages ids) deletes entirely — no schedule to cross-reference.

### 5.4 Range / type validation

New validation in `reader.py` or `workbook.py`:

- `1 <= start <= end <= project_months` per WP → raise with WP id and offending values.
- `type ∈ {"flat", "front-loaded", "bell", "back-loaded"}` per WP → raise with WP id and offending type.

Previously these were loose — `_wp_sched_months` silently produced 0 for invalid ranges, and `distribute` fell back to "flat" for unknown types. With one source of truth, make both explicit.

---

## 6. Builder signature sweep

### 6.1 Files touched (`wp_schedule` parameter removed)

Grep count: 77 occurrences across 11 files. Of these, the parameter appears in signatures of 8 public functions and 1 private helper.

| File | Functions changed | New behaviour |
|---|---|---|
| `budget_planner/workbook.py` | `_build_all_sheets`, `rebuild_workbook`, `_validate_inputs` | Drop `wp_schedule` kwarg; drop schedule-specific validation block |
| `budget_planner/reader.py` | `read_config` returns dict **without** `wp_schedule` key | Callers updated; type annotation shrinks |
| `budget_planner/config_sheet.py` | `build_config_sheet` — drop `wp_schedule` param; extend `_write_wp_budget_rows` to write start/end/type cols | Deletes `_write_wp_schedule` entirely |
| `budget_planner/plan.py` | `build_plan_sheet` | Reads `wp["start"]`, `wp["end"]` directly |
| `budget_planner/capacity.py` | `build_capacity_sheet` | See §7 |
| `budget_planner/timelog_plan.py` | `build_timelog_plan_sheet`, `_build_wps`, yearly-budget term builder | `_wp_sched_months` deleted; call sites inline |
| `budget_planner/actuals.py` | `build_actuals_sheet` | Reads start/end from work_packages |
| `budget_planner/actuals_hours.py` | `build_actuals_hours_sheet` | Reads start/end from work_packages |
| `budget_planner/trends.py` | `build_trends_sheet` | Reads start/end from work_packages |
| `budget_planner/plan_forecast.py` | `build_plan_forecast_sheet` | Reads start/end from work_packages |
| `budget_planner/layout.py` | Constants listed in §4.1 | — |

Scripts also touched:
- `scripts/shape_fingerprint.py` — drop `wp_schedule` from fingerprint shape.
- `scripts/build_from_shape.py` — drop `wp_schedule` from shape restore path.

### 6.2 Call-site pattern (before → after)

**Before** (typical):
```python
sched_entry = next(s for s in wp_schedule if s["id"] == wp["id"])
active_start = sched_entry["start"]
active_end = sched_entry["end"]
```

**After:**
```python
active_start = wp["start"]
active_end = wp["end"]
```

**Before** (`timelog_plan.py:657`):
```python
def _wp_sched_months(wp_code: str, wp_schedule: list[dict]) -> int:
    return sum(
        int(s["end"]) - int(s["start"]) + 1
        for s in wp_schedule if s["id"] == wp_code
    )
```

**After:** deleted. Call sites (one in `timelog_plan.py:926` for the `wp_sched_months` literal, one in Doc A's yearly-budget term builder) become:

```python
int(wp["end"]) - int(wp["start"]) + 1
```

---

## 7. `capacity.py` filter simplification

### 7.1 Current logic (`build_capacity_sheet`, ~line 180-200)

```python
# Build per-WP active-months set from wp_schedule (may span multiple segments)
wp_active_months_map = {}
for wp in work_packages:
    active = set()
    for s in wp_schedule:
        if s["id"] == wp["id"]:
            active.update(range(s["start"], s["end"] + 1))
    wp_active_months_map[wp["id"]] = active

# Per-assignment, intersect assignment window with WP active months
for asgn in assignments:
    assignment_months = set(range(asgn["start"], asgn["end"] + 1))
    valid_months = sorted(assignment_months & wp_active_months_map[asgn["wp"]])
    monthly_values = distribute(asgn["total_fte"], len(valid_months), wp_curve)
```

### 7.2 New logic

```python
# Per-WP active range is a single interval — look up by id
wp_range_map = {wp["id"]: (wp["start"], wp["end"]) for wp in work_packages}

# Per-assignment, clip assignment window to WP's single active range
for asgn in assignments:
    wp_start, wp_end = wp_range_map[asgn["wp"]]
    start = max(asgn["start"], wp_start)
    end = min(asgn["end"], wp_end)
    valid_months = list(range(start, end + 1)) if start <= end else []
    monthly_values = distribute(asgn["total_fte"], len(valid_months), wp_curve)
```

Similarly `wp_type_map` (line 755) simplifies:

```python
# Before
wp_type_map = {}
for entry in wp_schedule:
    wp_type_map[entry["id"]] = entry.get("type", "flat")

# After
wp_type_map = {wp["id"]: wp.get("type", "flat") for wp in work_packages}
```

### 7.3 Grey-fill preservation

The "grey + empty" rendering for out-of-window months is preserved. An assignment M10-M20 against a WP ending M17 still produces M18, M19, M20 as out-of-WP-range → grey + empty, exactly as today. The only semantic loss is the "gap months inside the WP's union range" tier — which, per the new schema, can no longer exist.

`tests/test_gapped_wp.py::test_out_of_window_months_are_grey_and_empty` becomes the last surviving scenario; it moves into `test_capacity.py` (or nearest equivalent) as `test_assignment_past_wp_end_grey_empty`, rewritten to a single-range WP fixture. Whether to keep this as a dedicated test or let existing capacity coverage absorb it is an implementation judgement call — see §9.

---

## 8. Doc A reconciliation

Doc A (`docs/plans/archived/2026-04-19-timelog-plan-enhancements-design.md`) shipped 2026-04-19 and is now archived. Two hooks in Doc A reference Doc B explicitly:

1. **D5** (decisions log): "Yearly budget proportion = `active_months_in_year / total_sched_months`, both segment-aware" — consistent with `_wp_sched_months` convention; survives later WP gap removal without formula rewrite.
2. **§5 note** (line 204): "`total_sched_months_for_wpk` = `_wp_sched_months(wpk.code, wp_schedule)` — respects multi-segment WPs today; if WP gap logic is removed later (Doc B), this collapses to `end - start + 1` with no formula change."

### 8.1 Integration sequencing

Two orderings possible:

**Order 1: Doc B first, then Doc A.** Under this order, Doc A's Phase 4 (yearly budget rows) is implemented directly against the collapsed schema — its yearly-budget term builder uses `wp["end"] - wp["start"] + 1` from day one. No `_wp_sched_months` helper to call. Doc A's §5 note becomes historical.

**Order 2: Doc A first, then Doc B.** Under this order, Doc A ships with `_wp_sched_months(...)` calls; Doc B's helper deletion turns into a one-liner rewrite per term. Doc A's Phase 4 `test_yearly_budget_multi_segment_wp` test deletes when Doc B's schema lands (the scenario is impossible).

**No coupling:** neither order blocks the other. Doc A's D5 was written precisely to decouple. Recommendation: whichever is ready first ships first. Doc B's diff is wider; Doc A's is deeper on Timelog Plan specifically. A code-review capacity argument favours Order 1 (smaller Doc A diff against a cleaner schema), but it's a preference, not a requirement.

### 8.2 Concrete Doc A diffs if Order 2

- §5 line 204: delete the hedge, state `total_sched_months = wp["end"] - wp["start"] + 1`.
- §6 Phase 4: test name `test_yearly_budget_multi_segment_wp` → delete (scenario gone). The single-WP-single-year test remains.
- §5 formula spec: the term formula stays identical; only the build-time helper call changes from `_wp_sched_months(wpk.code, wp_schedule)` to `int(wpk["end"]) - int(wpk["start"]) + 1`.

---

## 9. Test sweep

### 9.1 Deletions

- `tests/test_gapped_wp.py` — entire file. Scenario is impossible under new schema.

### 9.2 Additions

- `tests/test_reader_new.py::test_rebuild_rejects_legacy_wp_schedule_section` — build a fixture workbook with a "WP SCHEDULE" section on Config; call `read_config` or `rebuild_workbook`; assert `ValueError` with "Legacy Config layout detected" in the message.
- `tests/test_reader_new.py::test_rebuild_rejects_duplicate_wp_id` — build a fixture with the merged section but a duplicated WP id row; assert `ValueError` with "Duplicate WP id" in the message.
- `tests/test_capacity.py::test_assignment_past_wp_end_grey_empty` (if not already implicit) — single-range WP, assignment extends past WP end, verify out-of-WP months are grey + empty. Replaces the surviving scenario from `test_gapped_wp.py`.

### 9.3 Incidental updates

Every test that currently constructs `wp_schedule = [...]` as a fixture list needs updating to merge start/end/type into its `WORK_PACKAGES` fixture. Grep target:

```bash
grep -rn 'wp_schedule\s*=' tests/
```

Expected count: ~10-15 test files. All are mechanical edits — move three keys per WP from one dict to another.

Per `~/.claude/rules/python-refactors.md` ("Removing exported constants"): before committing, `grep -rn 'wp_schedule' tests/` to confirm zero residual references beyond the deleted-or-updated set.

### 9.4 Sibling test-fixture guard

Per `python-refactors.md` ("Tightening optional → required params"): `build_config_sheet`, `build_capacity_sheet`, `build_plan_sheet`, etc. have direct test-fixture callers (not just orchestrator). Each signature change breaks fixtures unless fixtures update in lockstep. The single-PR discipline keeps this atomic; no phase where fixtures lag.

---

## 10. Docs to update

Per CLAUDE.md critical rule #6 and the "Superseded docs: consume before archive" global rule.

| File | Change |
|---|---|
| `docs/design-invariants.md` | Add: "Single-range WP schedule — every WP has exactly one (start, end) pair; reader hard-fails on multi-row or legacy two-section Config layout." Remove any existing invariant about gap-month semantics. |
| `docs/internals.md` | § Sheet Builders → Config: update WP BUDGET section column count (7 → 10); delete WP SCHEDULE section description. § Extraction Functions: drop `_read_wp_schedule` row. § Reader and Rebuild: note the two new hard-fail error paths. |
| `docs/architecture.md` | Mermaid diagrams referencing `wp_schedule` as a separate data node collapse to a single WP node. Cross-sheet formula graph: `_wp_sched_months` node removed. |
| `docs/grant-domain.md` | Remove any prose about multi-segment WPs or gap months. Add one line: "Every WP has a single contiguous active range." |
| `docs/common-tasks.md` | "Add WP" SOP updated: one row on WP BUDGET with start/end/type, no separate schedule step. Remove any "Add WP with gap" SOP if present. |
| `CLAUDE.md` trigger table | Remove `wp_schedule` / gap-month entries. Add "WP schedule, single range, migration from legacy workbook" row pointing at this doc (once archived). |
| `.claude/rules/openpyxl-patterns.md` | No change — no new openpyxl pattern introduced. |
| `.claude/rules/openpyxl-readers.md` | No change — reader pattern guidance still applies; specific section name `_read_wp_schedule` is no longer referenced. |
| `.claude/rules/dev-workflow.md` | No change. |

Grep targets for stale prose (per dev-workflow.md "Bucket-migration tasks" and "vN → vN+1" guidance):

```bash
grep -rni 'wp.schedule\|WP SCHEDULE\|multi.segment\|gap.month\|gap.aware' docs/ CLAUDE.md .claude/rules/
```

Every hit either updates, deletes, or moves into this doc's Decisions log as an audit trail reference before archive.

---

## 11. Decisions log (2026-04-19 brainstorm)

| # | Decision | Rationale |
|---|---|---|
| D1 | Hard-fail on rebuild when Config has multi-row per WP or legacy two-section layout | Auto-collapse would silently inflate WP active windows (Ceiling €/PM drift); init-time-only validation defers the failure to a cryptic downstream error. Hard-fail gives one clear sentence and an obvious fix. |
| D2 | Merge `WORK_PACKAGES` + `WP_SCHEDULE` into a single `list[dict]` with extended schema | Split is vestigial once multi-row is forbidden. Removes `wp_schedule` parameter from 11 files; removes one whole dict from `config.py`; removes one whole section from Config sheet. |
| D3 | Merge the Config sheet's WP SCHEDULE section into WP BUDGET as three new cols (append right) | "Remove all related complexity" implies visible complexity on Config too. Append-right minimises layout shift vs reordering. |
| D4 | Reader raises TWO distinct errors: layout detection + duplicate-id detection | Layout case is the realistic one for existing users; specific error message with migration recipe beats a cryptic section-parse failure. Duplicate-id covers hand-edited merged sections. |
| D5 | Single-PR rollout — no phased split | Amputation is atomic at the schema level; any two-phase split ships intermediate scaffolding (reader stitches old layout) that gets deleted in phase 2. Doc A's precedent for 5-phase plans doesn't apply — no re-work axis. |
| D6 | `_wp_sched_months` deleted, not preserved as a trivial wrapper | Inline `end - start + 1` at the two call sites is clearer than indirection; wrapper wouldn't survive the first code review. |
| D7 | `tests/test_gapped_wp.py` deleted entirely; two new hard-fail tests added to `test_reader_new.py` | Per-scenario test files for "reject X" are overkill; reader-rejection tests naturally live next to other reader tests. |
| D8 | `capacity.py` grey-fill preserves the "out-of-WP-range but in-assignment-window" tier | Still a real scenario under single-range schema; only the "gap months inside WP union" tier disappears. Visual semantics for user unchanged in the common case. |
| D9 | `type` field preserved on merged schema — it drives capacity-curve distribution | Briefly considered dropping; load-bearing in `capacity.py::distribute`. |
| D10 | `config.py` seed data edits WP3 from two-row (M1-M3, M5-M7) to single-row (M1-M7) | Seed was a historical artefact, not a live use case; collapsing aligns seed with the new schema. Flagged in `config.py` docstring. |
| D11 | Append new cols (Start, End, Type) on right rather than reordering inputs-left / formulas-right | Smallest ripple; the reorder is a defensible but separate cleanup. |
| D12 | Existing `workbook.py:157-166` duplicate-ID error message rewritten — "use WP SCHEDULE" guidance removed; "split into distinct IDs" fallback retained | Current message directs users to a section that no longer exists. Split-ID advice is still valid. |
| D13 | Strict range validation (`1 <= start <= end <= project_months`) and strict type enum validation added explicitly | Previously loose — `_wp_sched_months` returned 0 for invalid ranges; `distribute` fell back to "flat" for unknown types. With one source of truth, make both explicit errors. |
| D14 | Doc A and Doc B are independent; either can ship first | D5/§5 of Doc A were written to make the schema collapse formula-neutral. No sequencing constraint. |
| D15 | Scripts `shape_fingerprint.py` and `build_from_shape.py` updated in the same PR | They serialize/deserialize the config shape; shape change propagates. |
| D16 | Per-builder edits consolidated: one file = one pass (signature drop + internal rewrite together). No builder appears in two sequential steps. | Signature change and internal rewrite are causally linked for `capacity.py` and `timelog_plan.py` — the helper/intersection code iterates the very parameter being removed. Splitting the work leaves the tree non-compiling between steps. Re-work audit per skill Step 7 item 5. |

---

## 12. Implementation checklist (single-PR order of edits)

Ordered to keep the tree buildable and tests passing at each commit boundary (though the PR squashes to one commit for history hygiene). Developer working the PR can break into logical commits mid-flight; the final merge is one squash.

1. **`config.py`:** collapse seed `WP3` to single-row `(1, 7, "front-loaded")`. Move `start`/`end`/`type` fields from each `WP_SCHEDULE` row into the corresponding `WORK_PACKAGES` dict. Delete `WP_SCHEDULE` constant. Update file docstring to note single-range invariant.
2. **`layout.py`:** add `COL_WP_START=8`, `COL_WP_END=9`, `COL_WP_TYPE=10`; bump `WP_BUDGET_LAST_COL = 10`. Delete `COL_SCHED_*` + `SCHED_LAST_COL`. Delete `"wp_schedule"` from `SECTION_HEADER_LABELS`.
3. **`config_sheet.py`:** extend `_write_wp_budget_rows` to write Start/End/Type per WP with `BLUE_INPUT`, data validation. Delete `_write_wp_schedule` function + its call in `build_config_sheet`. Drop `wp_schedule` parameter from `build_config_sheet` signature.
4. **`reader.py`:** add legacy-layout detection at top of `read_config` (§5.1). In `_read_wp_budget`, read Start/End/Type cols, add duplicate-id check (§5.2). Delete `_read_wp_schedule`. Drop `wp_schedule` key from returned dict.
5. **`workbook.py`:** rewrite duplicate-id error message (§5.3). Delete schedule-specific validation at lines 182-184. Drop `wp_schedule` kwarg plumbing through `_build_all_sheets` and `rebuild_workbook`.
6. **Per-builder sweep** — each file gets ALL its edits in one pass (signature + internal logic together; splitting by concern leaves the tree non-compiling mid-sequence):
   - **`plan.py`:** drop `wp_schedule` param; callsites read `wp["start"]` / `wp["end"]` / `wp["type"]` directly.
   - **`capacity.py`:** drop `wp_schedule` param; rewrite `wp_active_months_map` construction per §7.2 (single-range clip); rewrite `wp_type_map` per §7.2 (iterate `work_packages` instead of `wp_schedule`).
   - **`timelog_plan.py`:** drop `wp_schedule` param; rewrite `_build_wps` to read start/end off each WP dict; delete `_wp_sched_months`; inline `int(wp["end"]) - int(wp["start"]) + 1` at the Config literal write site (line 926). If Doc A has already landed, also rewrite its yearly-budget term builder per §8.2.
   - **`actuals.py`:** drop `wp_schedule` param; callsites read `wp["start"]` / `wp["end"]` directly.
   - **`actuals_hours.py`:** drop `wp_schedule` param; callsites read `wp["start"]` / `wp["end"]` directly.
   - **`trends.py`:** drop `wp_schedule` param; callsites read `wp["start"]` / `wp["end"]` directly.
   - **`plan_forecast.py`:** drop `wp_schedule` param; callsites read `wp["start"]` / `wp["end"]` directly.
7. **Scripts:** `shape_fingerprint.py`, `build_from_shape.py` — drop `wp_schedule` from shape dict.
8. **Tests:** delete `test_gapped_wp.py`. Add the two reader-rejection tests per §9.2. Sweep every remaining `wp_schedule =` in `tests/` per §9.3. Full suite green (`pytest tests/ -q`, `timeout: 300000`).
9. **Recalc:** `python scripts/recalc.py output/budget_plan.xlsx`; re-run suite per CLAUDE.md Rule 2.
10. **Docs:** apply §10 updates. Grep `wp.schedule|WP SCHEDULE|multi.segment|gap.month|gap.aware` across `docs/`, `CLAUDE.md`, `.claude/rules/` — zero residual hits before commit.
11. **Pyright:** `uv run pyright` clean before final commit.
12. **Archive:** `git mv docs/plans/2026-04-19-wp-gap-removal-design.md docs/plans/archived/` once the PR merges and living docs absorb the load-bearing content per the "Superseded docs: consume before archive" global rule.

**Re-work audit outcome:** steps 1-5 touch distinct files (`config.py` / `layout.py` / `config_sheet.py` / `reader.py` / `workbook.py`) — no overlap. Step 6 consolidates all per-builder edits so each of the 7 builders is edited exactly once. Original drafts had `timelog_plan.py` split across two steps (param drop + helper deletion) and `capacity.py` split across two steps (param drop + intersection rewrite); both splits were re-work traps because the signature change and the internal rewrite are causally linked — the helper/intersection code iterates the very parameter being removed, so either both happen together or the intermediate state won't compile. Merging into one per-file pass lands every artifact in final form on first edit. Recorded as D16.

---

## Spec self-review

**Placeholder scan:** no TBDs; every section names concrete files, line numbers, or exact strings.

**Internal consistency:** §3 schema matches §4 Config layout matches §6 builder sweep matches §7 capacity logic. D12 in §11 matches the rewritten error string in §5.3.

**Scope check:** single focused change (forbid multi-row). Doc A explicitly out of scope per the top-of-doc status block. Layout reorder and `type` collapse explicitly declared non-goals in §2.

**Ambiguity check:** the one judgement call flagged is §7.3 (whether to keep a dedicated `test_assignment_past_wp_end_grey_empty` or let existing capacity coverage absorb it) — marked as implementation-judgement rather than pre-decided. Everything else is explicit.

**Re-work audit (per skill Step 7 item 5):** single-PR rollout eliminates the cross-phase re-write axis, but the within-PR step sequence was audited for same-file overlap. Initial draft split `timelog_plan.py` and `capacity.py` across two steps each (signature drop separately from internal rewrite); both splits were re-work traps because the removed parameter is iterated by the code being rewritten. §12 was restructured per-file — each builder is edited exactly once. See D16. The one cross-doc ordering concern (Doc A vs Doc B) is addressed in §8 — both orderings are formula-neutral.
