# Config Through-Month Cleanup — Design

Date: 2026-04-18
Status: Approved — ready for implementation

## Problem

Config currently carries 5 rows for through-month bookkeeping:

| Row | Named range                   | Type  | Content                                 |
|-----|-------------------------------|-------|-----------------------------------------|
| 7   | `eur_actuals_through_month`   | blue  | User-entered integer, EUR trust cutoff  |
| 8   | `hours_actuals_through_month` | blue  | User-entered integer, hours trust cutoff|
| 10  | `months_elapsed_hours`        | black | `=MAX(1, $B$8)`                         |
| 11  | `months_elapsed_eur`          | black | `=MAX(1, $B$7)`                         |
| 12  | `reconciled_through_month`    | black | `=MIN($B$7, $B$8)`                      |

User no longer wants to hand-enter the blue inputs and finds the Config clutter unnecessary — the through-month values should be **derived** from the Actuals sheets themselves (last column with any numeric value), with an **override field on the target sheet** for the rare case when user wants to force a tighter cutoff.

## Goals

1. Remove all 5 rows from Config.
2. Auto-derive EUR / Hours through-month from their respective Actuals sheets; provide a per-sheet optional override cell.
3. Minimize named-range surface: 5 → 2. Drop `months_elapsed_hours`, `months_elapsed_eur`, `reconciled_through_month`; inline their one-line formulas at each callsite.
4. Keep downstream consumers unchanged — they still resolve `eur_actuals_through_month` / `hours_actuals_through_month` via named ranges; only the `attr_text` targets move to their natural home-sheet.

## Non-goals

- No change to downstream formula semantics. Integer month-indices with `0` = "no actuals yet" behaviour preserved.
- No migration of pre-existing Config!B7/B8 values into the new override cells — override starts blank on rebuild. User re-enters if wanted.
- No renaming of the 2 surviving named ranges.

## In scope (files touched)

- `budget_planner/config_sheet.py` — delete 5 param rows, TypedDict entries, named-range definitions
- `budget_planner/actuals.py` — add override + auto-derived + effective cells; inline `MAX(1, …)` at 3 `months_elapsed_eur` callsites
- `budget_planner/actuals_hours.py` — add override + auto-derived + effective cells
- `budget_planner/trends.py` — inline `MAX(1, …)` at 2 `months_elapsed_hours` callsites; inline `MIN(eur, hours)` at 3 `reconciled_through_month` callsites
- `budget_planner/reader.py` — drop through-month parsing and `_optional_keys`
- `scripts/shape_fingerprint.py` — drop 3 names from `SAFE_NAMED_RANGES`
- Tests: `test_config_sheet.py`, `test_rebuild_shape_drift.py`, `test_shape_fingerprint.py`, plus any Config-named-range existence asserts
- Docs: `docs/internals.md`, `docs/design-invariants.md`, `docs/architecture.md`, `docs/directory-map.md`

## Out of scope

- Renumbering Config rows 1-6 and row 9 (Version Stamp) — they stay; rows 7-8 and 10-12 simply disappear, following sections shift up by 5.
- Any change to Actuals-Hours skip-and-cumulate semantics.
- Any change to Trends NA()-clipping chart behaviour.

## Section 2 — Per-sheet layout + derivation formula

### Layout pattern (identical on Actuals and Actuals--hours)

A single new row at the **top** of each sheet (shift existing rows down by +1), 4 cells in a horizontal strip:

| Col | Content | Style | Named? |
|-----|---------|-------|--------|
| A   | Label, e.g. `EUR Actuals Through Month:` | `HEADER_FONT` | no |
| B   | **Override** (blank by default; integer ∈ [1, project_months] when set) | `BLUE_INPUT`, `fmt="0"` | no |
| C   | **Auto-derived** (scan formula) | `BLACK_FORMULA`, `fmt="0"` | no |
| D   | **Effective** = `IF(ISNUMBER(B), B, C)` | `BLACK_FORMULA`, `fmt="0"` | **yes** → `eur_actuals_through_month` / `hours_actuals_through_month` |

All 4 cells get `sized_comment` cell comments: label explains trust-boundary purpose; override explains *when* to use it + blank-by-default behaviour; auto explains scan semantics + "do not edit"; effective explains consumer-facing precedence rule (override wins when numeric).

This is not a 3-row vertical stack for option-(b) — the auto-derived cell sits horizontally alongside override and effective, so option-(b) visibility is preserved in a single new row.

### Derivation formula (scan)

```
= SUMPRODUCT(MAX((COLUMN(<data_range>) - COLUMN(<first_month_col>) + 1) * ISNUMBER(<data_range>)))
```

- Returns `0` when no numeric values; otherwise the largest 1-based month-index with any numeric value. Matches Q1=(a) semantics.
- Non-CSE (regular entry). Works in LibreOffice recalc and in Excel.
- `<data_range>` scope:
  - **Actuals**: `$<month_col_start_letter>$<first_wp_row>:$<month_col_end_letter>$<last_wp_row>` — WP data block only, **excludes the TOTAL row** (TOTAL becomes non-zero once any WP cell has a value, which would mask per-WP gaps).
  - **Actuals--hours**: programme-WP data rows across all persons, **excludes External-commitment rows and per-person subtotal/EU-total rows** (we want the "programme actuals through" semantic, not external activity).

### Effective-cell formula

```
= IF(ISNUMBER(<override_cell>), <override_cell>, <auto_cell>)
```

Blank/text in override → falls through to auto-derived. Number in override → wins.

### Row position shift consequences

- Actuals: row 1 ("As of:" dates) → row 2; section header → row 3; year/month headers → rows 4-5; WP data from row 6. Builder layout dict field names stay the same; absolute row numbers in the dict shift +1.
- Actuals--hours: analogous +1 shift.
- Downstream sheets reference Actuals cells via named ranges, not hardcoded row numbers. Before implementation, grep for literal `Actuals!<col><row>` / `Actuals--hours!<col><row>` refs to confirm no such ties exist.
- `freeze_panes` on both sheets re-targets to the new header row.

### Risk flag

`.claude/rules/openpyxl-patterns.md` calls out SUMPRODUCT + 2D-array patterns as needing Excel verification. Implementation step includes: generate workbook → open in Excel → confirm Effective cell returns expected integer under both empty and partial-data states.

## Section 3 — Consumer rewrites (inlining the 3 deleted named ranges)

### Callsite inventory (verified from grep)

| Dropped named range | Callsites | Inline replacement |
|---|---|---|
| `months_elapsed_hours` | trends.py L442, L448; **actuals.py L247, L330** (4 total) | `MAX(1, hours_actuals_through_month)` |
| `months_elapsed_eur` | actuals.py L207, L294 (2 total) | `MAX(1, eur_actuals_through_month)` |
| `reconciled_through_month` | trends.py L286 (display), L327-329 (cum forecast EUR), L338-340 (cum actual EUR) (3 total) | **local helper cell on Trends** (see below) |

Correction vs. Section 1 consumer table: `months_elapsed_hours` has 4 callsites, not 2 — the two Forecast Breach In formulas in `actuals.py` belong to the hours-driven Forecast block.

### `months_elapsed_*` inlining

Each callsite is a denominator inside an existing breach-in / avg-FTE formula. Example (actuals.py L203-207):

```python
# before
f"({ceil_l}{row}-{cum_l}{row})/({cum_l}{row}/months_elapsed_eur)"

# after
f"({ceil_l}{row}-{cum_l}{row})/({cum_l}{row}/MAX(1,eur_actuals_through_month))"
```

Zero semantic change. No ROUND/format implications (these formulas return month-counts, not EUR).

### `reconciled_through_month` — keep 1 local helper cell on Trends

All 3 callsites live on Trends. Repurpose the existing display cell at trends.py L286 (currently `=reconciled_through_month`) to hold `=MIN(eur_actuals_through_month, hours_actuals_through_month)`. The other 2 callsites change `reconciled_through_month` → absolute same-sheet ref (`$<col>$<recon_row>`) to that helper.

Why: DRY at sheet level + preserves the auditor-visible "Reconciled Through Month" labelled row on Trends (the row is useful there; it's only superfluous on *Config*).

The helper cell position flows through the Trends layout dict — no new cross-function plumbing.

## Section 4 — Config shrink + reader + shape fingerprint

### `budget_planner/config_sheet.py`

1. **`param_defs` list** (L662-688): remove 2 entries (`EUR Actuals Through Month`, `Hours Actuals Through Month`). Version Stamp stays; its index shifts 9 → 7.
2. **`param_rows` TypedDict literal** (L694): drop those 2 keys; `"Version Stamp": 9` → `"Version Stamp": 7`.
3. **`ConfigParamRows` TypedDict in `layout.py`**: drop those 2 keys.
4. **Derived-params block** (L730-739): delete entirely — the `derived_params` list, the for-loop, and `derived_first_row`.
5. **`_param_label_comments`** (L745-808): remove 5 entries (2 through-months + 3 derived). Keep FTE Rate, EU Annual Ceiling, Hours per FTE Month, Version Stamp comments.
6. **Named-range definitions** (L871-900): drop 5 `DefinedName` adds. The 2 survivors are re-added by the new Actuals / Actuals--hours builders.
7. **`budget_section_row`** (L813): `len(param_defs) + len(derived_params) + 2` → `len(param_defs) + 2`. Net: Config content shifts **up by 5 rows**.
8. **Sanity assertion** (L706-709): keys off index vs. dict value — both shrink consistently. No change needed.

### `budget_planner/reader.py`

1. **`_param_map`** (L298-312): drop the 2 through-month entries; remove L293-297 collision-comment block.
2. **`_optional_keys` block** (L325-335): remove — no remaining optional keys.
3. Reader no longer surfaces `eur_actuals_through_month` / `hours_actuals_through_month` in `project_params`. Config builder's `project_params.get(...)` on those keys is removed (L675, L681 — already covered by param_defs shrink).

### New reader helpers (per-sheet override)

The new override cells (col B, new row 1 on Actuals / Actuals--hours) are blue input; rebuild must preserve user-typed values.

```python
def read_actuals_override(ws) -> int | None:
    """Return override integer at row 1, col B, or None if blank/non-numeric."""
    v = ws.cell(1, 2).value
    return int(v) if isinstance(v, (int, float)) else None
```

Orchestrator calls this on the old sheet before rebuild and passes the result into the Actuals / Actuals--hours builder via new kwargs `eur_through_override` / `hours_through_override`. Builder writes back to cell when non-None; else leaves blank.

### `scripts/shape_fingerprint.py`

**`SAFE_NAMED_RANGES`** (L39-41): drop 3 names.

```diff
 "eur_actuals_through_month", "hours_actuals_through_month",
-"months_elapsed_hours", "months_elapsed_eur",
-"reconciled_through_month",
```

Survivors keep their entries — fingerprint check asserts existence, not target sheet, so the relocation is transparent.

## Section 5 — Tests + docs

### Test updates (by file)

| File | What changes |
|------|--------------|
| `tests/test_config_sheet.py` | Drop assertions on the 5 removed rows (Config!B7/B8/B10/B11/B12, their named ranges, their comments). Update `ConfigParamRows` expectations (9 → 7 for Version Stamp). Update any `budget_section_row` position asserts by −5. |
| `tests/test_shape_fingerprint.py` | Drop expectations for the 3 removed entries in `SAFE_NAMED_RANGES`. |
| `tests/test_rebuild_shape_drift.py` | Verify fixtures that set `eur_actuals_through_month` / `hours_actuals_through_month` no longer target Config — they either set the new override cells (if testing override path) or rely on derivation (if testing default). Any absolute-row-number asserts on Config after the deleted rows shift by −5. |
| `tests/test_trends.py`, `tests/test_trends_reconciliation.py` | Any `wb.defined_names["months_elapsed_hours"]` / `...["reconciled_through_month"]` assertions become cell-coord assertions against the Trends helper + inline formulas. Behaviour assertions (computed values) stay — values unchanged. |
| `tests/test_actuals.py`, `tests/test_actuals_forecast_block.py`, `tests/test_actuals_breach_todate.py` | Tests that pre-set a through-month value on Config!B7/B8 → switch to setting Actuals!B1 override. Tests that read the through-month via named range still work (named range moved, consumers didn't). Breach-in formula assertions — if they inspect the formula *text*, update expected string from `months_elapsed_eur` → `MAX(1,eur_actuals_through_month)`. |
| `tests/test_actuals_hours.py`, `tests/test_actuals_hours_forecast_shading.py` | Analogous for hours side (Actuals--hours!B1 override, inlined `months_elapsed_hours`). |
| `tests/test_capacity_todate.py` | Windowed-plan-sum idiom (`IF(hours_actuals_through_month<=0,0,SUM(...))`) unchanged — consumer reads via named range. Setting fixtures: switch Config!B8 → Actuals--hours!B1 override. |
| `tests/test_plan_forecast.py`, `tests/test_summary_partial_year.py` | Same pattern — switch through-month fixture setter from Config to Actuals override. |

Full-round-trip test required per `.claude/rules/dev-workflow.md` § Testing: at least one test must cover `init → user sets override via openpyxl → rebuild → reopen → assert effective cell honours override`. Add to `test_rebuild_shape_drift.py` if not naturally covered by an existing sheet-preservation test.

### Doc updates

| Doc | What changes |
|-----|--------------|
| `docs/internals.md` | Per-Sheet Notes — Config section loses 5 row descriptions; Actuals + Actuals--hours sections gain "Row 1: through-month control strip (override / auto / effective)". Extraction Functions table — `read_config_sheet` no longer returns through-months in `project_params`; add entry for new `read_actuals_override`. |
| `docs/design-invariants.md` | Trust-boundary invariant wording updates to reflect: (a) through-month is now *derived* (auto), with user override as the explicit opt-in mechanism; (b) cells live on their home-sheet, not Config. |
| `docs/architecture.md` | Any Mermaid C4/flow diagram that shows "user edits Config through-month" → update node to "auto-derived from Actuals, override cell on Actuals". |
| `docs/directory-map.md` | Minor — if it lists `config_sheet.py`'s responsibilities for through-month, drop those; add note under `actuals.py` / `actuals_hours.py`. |
| `CLAUDE.md` | Reference index row for "Cross-sheet invariant, trust boundary" still points at `design-invariants.md`; no path change. |
| `docs/plans/2026-04-14-t5-simplification-design.md` | Live (non-archived) plan referencing `hours_actuals_through_month`. Check: is it describing a shipped thing (reference only) or pending work (consumer needing update)? If shipped/reference, no change needed beyond a relocation note. |
| Memory pointers | `project_state.md` + `MEMORY.md` entries that reference "through-month" cells on Config become stale; update after implementation lands. |

## Section 6 — Implementation sequencing + verification gate

### Why one atomic commit (not staged)

Intermediate states break. Concrete example: if we delete Config!B7 *before* relocating `eur_actuals_through_month`, the derived row 11 formula `=MAX(1,$B$7)` returns `#REF!` and every downstream consumer breaks. The inverse order fails too — you can't have two `DefinedName("eur_actuals_through_month", ...)` entries. So the builder-side change lands in a single coherent commit:

**Commit 1 (all of):**
1. `config_sheet.py` — delete 5 param rows, 3 derived formulas, 5 named-range defs.
2. `actuals.py` — new row 1 strip (label + override + auto + effective); shift existing rows +1; `freeze_panes` retarget; register `eur_actuals_through_month` DefinedName; inline `MAX(1,…)` at 2 callsites (L207, L294); inline `MAX(1,…)` at 2 more callsites (L247, L330) that use `months_elapsed_hours`.
3. `actuals_hours.py` — analogous strip + DefinedName for `hours_actuals_through_month`; shift rows +1.
4. `trends.py` — inline `MAX(1,…)` at L442, L448; repurpose L286 display cell to `=MIN(eur_actuals_through_month, hours_actuals_through_month)`; L327-340 swap to `$<col>$<recon_row>` same-sheet refs.
5. `reader.py` — `_param_map` shrink, `_optional_keys` removal, new `read_actuals_override` helper; orchestrator plumbing.
6. `shape_fingerprint.py` — `SAFE_NAMED_RANGES` shrink.
7. `layout.py` — `ConfigParamRows` TypedDict shrink.
8. Tests — all updates from Section 5.

**Commit 2 (docs):** `docs/internals.md`, `docs/design-invariants.md`, `docs/architecture.md`, `docs/directory-map.md`, memory pointer updates.

### Verification gate (per CLAUDE.md Critical Rule 2)

After Commit 1 staged, before landing:

1. `python scripts/init_budget.py` — generate fresh workbook from config.py seed.
2. `python scripts/recalc.py output/budget_plan.xlsx` — LibreOffice recalc so `data_only=True` reads return numbers, not None.
3. `pytest tests/ -q` — full suite green. Expected: 510 → 510 (± test-count delta from Section 5 additions/removals).
4. `uv run pyright` — 0 errors / 0 warnings.
5. **Open in Excel** (not just LO — per `.claude/rules/dev-workflow.md`): confirm scan formula on each Actuals sheet returns expected integer under three states:
   - Empty Actuals → returns 0
   - Partial Actuals (data through month 3) → returns 3
   - With override set to 2 → effective returns 2 (override wins)
6. Manual round-trip: open generated workbook in Excel → type `5` into Actuals override cell → save → run `init` to rebuild → reopen → confirm override cell still shows `5` and effective shows `5`.

### Shape-drift safety

Per `.claude/rules/openpyxl-patterns.md`, `ws.insert_rows()` doesn't update DefinedNames. We're not using `insert_rows()` — we're changing the builder to write the new row 1 explicitly and shift subsequent row indices in code. The named range `attr_text` is set fresh on every build. No shape-drift risk from this change.

Post-implementation: memory note — the `project_cross_session_gotchas.md` line that says *"`eur_actuals_through_month` / `hours_actuals_through_month` default to 1"* becomes stale (default is now 0 from the scan formula, and no explicit default in builder). Update after landing.

