# 2026-04-22 — Indirect-view restructure (with-indirect vs without-indirect EUR)

**Status:** APPROVED 2026-04-22 (mybrain session). Pending implementation — scheduled for a future session. Do not auto-start implementation; the user explicitly chose to land the plan without starting the code changes.

**Pre-merge TODO — RESOLVED 2026-04-22:** user confirmed the accounting feed is without-indirect at source. D7 stands as written; no escalation to option C needed.

## Motivation

Per the finance lead: outside this project, every EUR number carries the EU 25% flat indirect rate (organisation P&L, other projects, EU disbursements, management queries). The workbook today stores all EUR without indirect and applies ×1.25 only in one place (Summary "Total Claimable"). Every time a management-facing view is pulled from the workbook, the reader has to mentally convert — and often doesn't, leading to silent misreads of "we have €X left" when the actual disbursable ceiling is €X × 1.25.

Additionally, the accounting monthly feed (what accounting sends as "actuals") is **without indirect** by convention, AND is structurally a **prognosis** (logged hours × assumed hourly rate baking in vacation + social tax + possibly internal overhead), not a reconciled actual. True actuals only crystallise at year-end reconciliation and can swing by ~€20k either direction. The workbook must preserve the without-indirect form for accounting-feed reconciliation while defaulting every external-facing view to with-indirect.

**Goal:** make the default-visible surface of every sheet with-indirect (external-facing, ready to present to management), while keeping the without-indirect accounting-reconciliation zones intact but collapsible behind Excel outline groups.

**Out of scope for this plan:**

1. **Formally modelling "accounting prognosis vs year-end reconciled actuals"** as two separate data streams. Larger independent change. This plan preserves today's single `Actuals` sheet as "best-effort accrual from accounting" and does not split it.
2. **`Trends` and `Summary` sheet overhaul.** Per user decision 2026-04-22, these sheets need a bigger semantic overhaul in a follow-on plan: they should switch from reading `Plan` as their budget-plan source to reading `Timelog Plan` instead. That overhaul will also redesign how Summary aggregates + how Trends displays headlines. This plan does NOT touch `Trends` or `Summary` — they keep their current structure and terminology verbatim until the follow-on plan lands. (Ceiling→Budget rename DOES NOT apply inside these two sheets for now; the follow-on plan will handle it.)
3. **`Plan` sheet semantic shift.** The same follow-on plan will strip EUR/budget semantics from `Plan` and make it purely "actual resource usage of people per WP". This plan leaves `Plan` EUR content (if any) untouched and unsuffixed — touching it would waste work the overhaul will undo.
4. **Multi-cost-category decomposition.** This workbook is single-category (Personnel only). There is no Travel / Equipment / Other direct-cost aggregate distinct from Personnel, so no "Total Eligible +i" aggregate-name layer is introduced — the two labels `Budget +i` and `Budget -i` cover every EUR budget view in scope.

## Terminology (global to this plan and everything it touches)

- **"Ceiling" is retired.** The workbook, docs, code, named ranges, and column headers all use **"Budget"** from now on. This plan executes the full sweep as part of its diff — every current "Ceiling" identifier across `budget_planner/**`, `tests/**`, `docs/**`, and workbook cells is renamed in a single pass.
- **Suffix convention for EUR headers** (decided in D6): `+i` for with-indirect, `-i` for without-indirect. Light-green fill `#E2EFDA` for with-indirect headers, pale-orange fill `#FBE5D6` for without-indirect headers. Applies uniformly across **in-scope sheets** (D8). Deferred sheets (Plan, Trends, Summary) keep their current unsuffixed headers until the follow-on overhaul.
- **Management** replaces any named individual from the management tier in all prose and examples.

## Decisions log

### D1 — Storage convention: combination of A + B

- **Default storage: A (without-indirect).** Blue-input cells continue to carry direct EUR (matches accounting's native feed, matches EU audit line items). Paired black-formula cells derive `×1.25` for with-indirect display.
- **Exception: B (with-indirect) where the with-indirect number is the authoritative external source.** Decided in D4: grant-agreement WP budgets (`Budget +i` — €1.625M is what EU disburses) store with-indirect and derive without-indirect (`Budget -i`) for the accounting-reconciliation check.
- **Why:** Aligns storage with whichever side is the authoritative input. Flipping accounting's monthly prognosis to with-indirect would be an anti-pattern (the source-of-truth cell would be derived-on-input, fighting the feed's native form). Flipping grant ceilings to with-indirect is correct (the grant agreement is the authoritative source and quotes with-indirect numbers).

### D2 — Structural level: B (section-band within each sheet)

- On every sheet, the default-visible region is **with-indirect** (external-facing, ready to present to management).
- The **without-indirect** reconciliation band sits below/beside and is wrapped in an Excel row/column outline group so it collapses via the `–` toggle.
- No new sheets are created; no cell-pair siblings throughout. The collapsible band is the single segregation primitive.
- **Why:** Matches user's literal requirement ("so I can collapse these") — Excel outlines collapse contiguous bands cleanly; they do not collapse individual cells. A sheet-level split (`Actuals-Accounting` + `Actuals`) would double the cross-sheet formula graph for no extra clarity.

### D3 — Actuals monthly grid: single grid with view-flip rollups (option A)

- The monthly grid (~48 months × WP rows) stays blue-input **without-indirect** (accounting's native feed form) and is wrapped in a **column outline group** so the whole monthly band collapses via the `–` toggle.
- **Default-visible rollup columns** to the right of the grid (Cum EUR, Budget, Variance, % of Budget, Breach-In) flip to **with-indirect**. Black formulas bake `×1.25` into the rollup.
- A small without-indirect rollup sibling (`Cum EUR -i`, `Budget -i`, etc.) sits **inside** the collapsible band for direct accounting reconciliation — hidden by default under the collapsed group.
- **Why not B (mirror grid):** doubles row count on `Actuals`, adds 48 pure `×1.25` formula cells per WP, no extra clarity since accounting's feed is cell-for-cell without-indirect anyway.
- **Why not C (flip monthly grid to with-indirect):** would force accounting's feed through a `/1.25` pre-entry cell, fighting the native form of the authoritative input source.

### D4 — Config grant-budget inputs: flip blue-input to with-indirect (option B), layout chosen to protect deferred consumers

- **Rationale:** the Grant Agreement states Total Eligible Costs as "EUR with-indirect". The authoritative source document quotes the gross number, so that's what the user types.
- **Column layout on `WP BUDGET` section** — chosen to preserve the numeric semantic of the old `Ceiling` column position:
  - **Old `Ceiling` column position** is retained. Header renames to **`Budget -i`** (orange fill). Cell type changes from blue-input to **black-formula** = `Budget +i / 1.25`. Numeric value in this cell equals today's `Ceiling` value after migration, so any code reading by position or by legacy named-range alias gets the same number.
  - **New `Budget +i` column** is **inserted immediately adjacent** (to the right). Blue-input, green fill, user types the with-indirect gross (e.g. `1,625,000`).
- **Why this layout:** deferred consumers (`summary.py`, `trends.py`, `plan.py`) read today's `Ceiling` column to get a without-indirect number. After migration, the cell at that position still holds a without-indirect number (derived via `÷ 1.25` instead of blue-input, but numerically identical). Deferred code keeps working unchanged. The follow-on overhaul can retarget at leisure.
- **Migration logic (one-shot, on rebuild):**
  1. Read pre-plan workbook's `Ceiling` column values at the old position.
  2. Seed the new `Budget +i` column with `value × 1.25`.
  3. Rewrite the old position's cell as black-formula `= Budget +i / 1.25`. Numeric output equals the original `Ceiling` value.
  4. Rename header text from `Ceiling` → `Budget -i`.
- **Migration safety (double-apply guard):** assumes pre-plan workbooks have without-indirect `Ceiling` values (matches every prior rebuild). If any row's `Ceiling × 1.25` would exceed the sanity threshold `€3,000,000` per WP, the reader emits a `RuntimeWarning` and requires the user to set an explicit env flag `GRANT_ACCEPT_CEILING_MIGRATION=1` before the rebuild proceeds. Protects against the double-apply failure mode where someone manually pre-with-indirect'd a cell.
- **Layout detection (as shipped in T1):** the reader detects pre-plan vs post-plan workbooks by inspecting the **col-5 header string** on the WP BUDGET header row — `"Ceiling (€)"` → legacy (run migration), `"Budget -i"` → post-plan (no-op). No separate version-stamp gate is required: `version_stamp` in this codebase is already a git-describe literal (`workbook._resolve_version_stamp()`), so it bumps naturally on every commit and would carry no schema-version semantics anyway. The header-string check is the deterministic gate, and it auto-disables migration on subsequent rebuilds since the builder now writes `Budget -i` at col 5.
- **Named ranges:** no existing named range currently targets the `Ceiling` column (grep confirmed — not in the named-range contract table of `docs/design-invariants.md`). This plan adds `wp_budget_with_indirect` (targets the new `Budget +i` column) as a reader-friendly alias. No legacy named range to retire.
- **In-scope downstream consumers** that reference the budget value in new-semantic form (i.e. want with-indirect) retarget to `Budget +i`: `actuals.py` Budget/Variance/% rollups, `plan_forecast.py` rollups, `timelog_plan.py` Yearly Budget row, `capacity.py` if applicable. The `Budget -i` sibling is consumed only by the collapsible-band accounting-reconciliation rows.

### D5 — Full rename sweep: "Ceiling" → "Budget" (global)

- **Scope of sweep:** every occurrence of `Ceiling` / `ceiling` across `budget_planner/**`, `tests/**`, `docs/**`, named-range strings, column headers, variable names, docstrings, and plan files is renamed to `Budget` / `budget` as part of this plan's diff.
- **Exception:** when quoting programme regulation or grant-agreement source text where the original wording is "budget ceiling" or "cost ceiling", preserve the quoted phrase and mark it as a quotation. Do not rewrite regulatory source text.
- **Named ranges** currently containing "ceiling" (if any) are renamed in the same pass; consumers follow. Readers of pre-plan workbooks resolve old names via one-shot migration.
- **Why in this plan, not a separate rename:** the indirect-view restructure already touches every consumer of the old ceiling column; doing the rename as a drive-by in this plan is cheaper than a separate PR that re-edits the same lines.

### D6 — Header suffix `+i` / `-i` with light-green / pale-orange fills (option A)

- **Suffix:** `+i` for with-indirect, `-i` for without-indirect. 2-character ASCII, speaks as "plus-i / minus-i". Rendered uniformly across fonts and LibreOffice.
- **With-indirect header fill:** `#E2EFDA` (light green). Signals "external-facing, management-ready". Applied to the default-visible zone on every sheet.
- **Without-indirect header fill:** `#FBE5D6` (pale orange). Signals "internal reconciliation only, do not quote externally". Applied to headers inside the collapsible without-indirect band.
- **Non-EUR headers** (month names, WP IDs, person names, labels) keep today's `LIGHT_BLUE_FILL` (`#D6E4F0`). The blue/green/orange three-way palette makes the EUR-vs-label distinction visible even on collapsed views.
- **No collision:** `#FFF2CC` stays reserved for `VACATION_FILL` on Timelog Plan; section-title dark blue `#2F5496` unchanged.
- **Suffix uniformity (on in-scope sheets only):** every EUR-bearing column header on sheets listed in D8's in-scope set carries `+i` or `-i`. No unmarked EUR headers permitted on in-scope sheets. Variance, %, Cum, Breach-In rollups all inherit the suffix from the underlying EUR view they represent. Deferred sheets (Plan, Trends, Summary) retain their current unsuffixed EUR headers until the follow-on overhaul.
- **Font color convention (Rule 3)** is additive, not replaced: blue=input / black=formula / green=xref continues to apply independently. A `Budget +i` black-formula cell still uses black font on the green fill.

### D7 — Hourly-rate semantics clarified: grade rate is without-indirect (sidebar A)

- **Resolution:** the transcript note that "the finance lead's rate bakes 25% in" is treated as a misreading. The workbook's `RATE BY GRADE` and `TEAM ROSTER` rates stay without-indirect — they already encode vacation, social tax, and internal house overhead, but NOT the EU 25% flat indirect. This matches what the code does today. **Confirmed by user 2026-04-22:** the accounting feed is without-indirect at source.
- **Consequence for hours → EUR conversion:** `Actuals-Hours` forecast-EUR = `hours / 143 × rate` continues to produce a **without-indirect** EUR number. The conversion formula is not touched.
- **Consequence for indirect presentation:** with-indirect forecast-EUR is a **display rollup** — `×1.25` applied at the consumer sheet (Summary, Plan-Forecast, Actuals forecast-monitor columns), not baked into the source-of-truth formula.
- **Grade rate header suffix:** per D6's uniformity rule, rate columns carry `-i` (orange fill) since they are without-indirect EUR values. The suffix applies to `Rate EUR/PM` cells in both `RATE BY GRADE` and the `TEAM ROSTER` rate column.
- **TODO — RESOLVED 2026-04-22:** user confirmed the accounting feed (and the rate it encodes) is without-indirect at source. No escalation to option C.

### D8 — Per-sheet change map and ordering

**Sheets in scope:** Config, Actuals, Plan-Forecast, Capacity, Actuals-Hours, Timelog Plan. Plus docs and tests.

**Sheets deferred to follow-on overhaul:** Trends, Summary, Plan (see Out-of-scope items 2–3 above).

#### Per-sheet diff

**Config** (detailed column-layout rules in D4)
- `WP BUDGET` section: old `Ceiling` column position becomes **`Budget -i`** (black-formula, orange fill `#FBE5D6`, value unchanged numerically). New **`Budget +i`** column inserted adjacent (blue-input, green fill `#E2EFDA`, user types with-indirect gross EUR, e.g. `1,625,000`). Net effect: deferred code reading the old position still gets a without-indirect number.
- `RATE BY GRADE` rate header: add `-i` suffix (orange fill). Values unchanged.
- `TEAM ROSTER` rate column header: add `-i` suffix (orange fill). Values unchanged.
- Any other `Ceiling` identifiers on Config (docstrings, comments, section titles): rename to `Budget`.
- New named range `wp_budget_with_indirect` targets the `Budget +i` column. No existing named range to rename.
- Bump `version_stamp`.

**Reader migration (`reader.py`)**
- One-shot migration on rebuild: if a pre-plan workbook has the old `Ceiling` column header (without-indirect), read its value, seed the new `Budget +i` column as `value × 1.25`, keep the original row mapping.
- Covers single migration path only — no separate script. Subsequent rebuilds see `Budget +i` already present and no-op the migration.

**Actuals**
- Monthly EUR grid (cols C..C+N−1) stays blue-input without-indirect. Wrap in Excel **column outline group** (collapsible `–` toggle).
- Month-section header row: add `-i` suffix on the section label; individual month names (Jan 2026, etc.) unchanged.
- Default-visible rollup columns to the right of the grid flip to with-indirect:
  - `Sum +i`, `Budget +i`, `Variance +i`, `Cum EUR +i`, `% of Budget +i`, `Breach-In +i`
  - Green header fill; black-formula fonts; formulas bake `×1.25` where the underlying aggregation is row-sum of without-indirect monthly cells.
- Inside the collapsible band (hidden by default), sibling rollups: `Sum -i`, `Budget -i`, `Variance -i`, `Cum EUR -i`, `% of Budget -i`, `Breach-In -i`. Orange header fill. These feed accounting-reconciliation visually when the band is expanded.
- Forecast-monitoring block (parallel cross-refs to Plan-Forecast) follows the same `+i` default / `-i` in-band pattern.
- Named ranges `eur_actuals_through_month` unaffected semantically — the through-month scalar doesn't change.

**Plan-Forecast**
- Monthly forecast EUR grid stays without-indirect (pulled from Actuals-Hours forecast-EUR leaves, which stay without-indirect per D7). Wrap in column outline group.
- Rollup columns flip to with-indirect: `Sum +i`, `Budget +i` (xref Config), `Variance +i`, `Cum EUR +i`, `% of Budget +i`.
- Inside band: `-i` siblings for the same rollups.
- Month-section header gets `-i` suffix.

**Capacity**
- FTE-dominant sheet. Scan for any EUR columns; if present, suffix per D6 and apply collapsible pattern. Rename any `Ceiling` reference to `Budget`.

**Actuals-Hours**
- Per-person hours grid: unchanged (hours are not EUR, no suffix).
- Forecast-EUR band: stays without-indirect at source per D7. Header gets `-i` suffix + orange fill. Wrap this band in a row outline group.
- New default-visible `Forecast EUR +i` single rollup column per person = `row-sum of forecast-EUR band × 1.25`. Green fill, black formula.

**Timelog Plan** — non-trivial layout change
- **Layout shift (per user request):** `Yearly total €` row and `Yearly budget €` row move to **column H**. The per-year breakdown columns (currently showing yearly numbers) start at column I onward.
- **Pre-impl verification step:** before editing, read current Timelog Plan layout (`budget_planner/layout.py` + `budget_planner/timelog_plan.py` helpers) to confirm what occupies column H today, and decide where displaced content shifts. Expected: a summary/label column currently sits in H and should move right, but this needs concrete verification at implementation time. Document the before/after column map in the impl PR description.
- **EUR display flips to `+i`:** all EUR cells on Timelog Plan (per-year EUR cols, `Monthly total €`, `Yearly total €`, `Yearly budget €`) display with-indirect — green header fill, formulas multiply the solver's without-indirect output by `1.25` at the display cell.
- `Yearly budget €` → renamed `Yearly Budget +i`. Pulls directly from Config `Budget +i` (no `×1.25` applied, already with-indirect).
- No `-i` band on Timelog Plan — this sheet is forward-looking / external-facing by nature, not an accounting reconciliation surface.

**Docs**
- `docs/grant-domain.md`: add "Indirect-view convention" section explaining `+i` / `-i` suffixes, fill colors, Budget/Ceiling terminology.
- `docs/design-invariants.md`: add invariant — "Every EUR-bearing column header carries `+i` or `-i`" — with named-range contract updates.
- `docs/internals.md`: update Cross-Sheet Reference Map for renamed named ranges; update per-sheet layout notes for Timelog Plan H-column shift.
- `CLAUDE.md` Reference Index: update pointer entries that mention "Ceiling".
- Note in all three docs: "Trends and Summary pending overhaul plan — their terminology is intentionally not updated in this pass."

**Tests**
- Update header-assertion fixtures: `tests/test_config_sheet.py`, `tests/test_actuals.py`, `tests/test_plan_forecast.py`, `tests/test_timelog_plan.py`, `tests/test_actuals_hours.py`, `tests/test_capacity.py` for new column names.
- Add migration round-trip test: pre-plan workbook with `Ceiling` column → rebuild → `Budget +i` populated with `old_value × 1.25`.
- Add indirect-view invariant check: for every WP, assert `Budget +i / 1.25 ≈ Budget -i` within floating-point tolerance.
- Add header-suffix invariant: scan every EUR-bearing column header on in-scope sheets and assert it ends with `+i` or `-i`.

#### Implementation ordering

Dependency-driven, single-PR rollout. Re-work audit applied: each file is touched in exactly one step (except `layout.py`, which lands once with all column-position changes).

1. **Config + layout.py + reader migration** (`budget_planner/config_sheet.py`, `budget_planner/layout.py`, `budget_planner/reader.py`, `budget_planner/styles.py`) — add `WITH_INDIRECT_HEADER_FILL` and `WITHOUT_INDIRECT_HEADER_FILL` styles; introduce `Budget +i` / `Budget -i` columns; suffix rate headers; migrate old `Ceiling` column.
2. **Per-sheet pass (parallelizable across files)** — `actuals.py`, `plan_forecast.py`, `timelog_plan.py` (includes H-column layout shift), `capacity.py`, `actuals_hours.py`. Each sheet builder gets: column/row renames, suffix/color header updates, outline-group wiring for the without-indirect band, formula retargets.
3. **Docs + CLAUDE.md** — `grant-domain.md`, `design-invariants.md`, `internals.md`, `CLAUDE.md` reference index. Add deferral note for Trends/Summary.
4. **Tests** — fixture updates, migration round-trip, invariant checks.
5. **Gate** — `python scripts/recalc.py output/budget_plan.xlsx && pytest tests/ -v`; manual LibreOffice open to visually confirm the collapsible outline groups work and the color fills render correctly.
6. **Commit** — one commit per step; PR contains 6 commits.

**Re-work audit:** step 2 files are each touched once. `layout.py` is touched once in step 1 (column positions and the H-column shift for Timelog Plan are baked in at the same time). `styles.py` gets two new fills added in step 1 and is not revisited. No file appears in two steps.

---

## Self-review outcome

Spec self-review applied 2026-04-22 (mybrain Step 7):

1. **Placeholder scan:** no `TBD` / `(pending)` / incomplete sections remain. One explicit non-blocking TODO at the top of the doc (finance-lead confirmation on grade-rate semantics) — intentional and scoped.
2. **Internal consistency:** D1's exception cross-reference updated from placeholder "D3" to correct "D4". Terminology section rewritten to reflect D6 locked. D4 ripple list pruned of deferred sheets. D6 suffix-uniformity rule qualified to in-scope sheets only.
3. **Scope check:** six sheet builders in scope, three deferred, single-PR rollout, ~6 commits. Not too large.
4. **Ambiguity check:** Config column-layout rewrite ensures the old `Ceiling` position's numeric semantic is preserved, so deferred consumers (Summary/Trends/Plan) continue to work without retargeting. Timelog Plan H-column shift gets an explicit pre-impl verification step. Migration has a double-apply sanity guard (`€3M` threshold + env-flag override). Version-stamp bump disambiguates pre/post-plan workbooks.
5. **Re-work audit:** each file touched in exactly one step. `layout.py` lands all column-position changes at once. `styles.py` gets two new fills added once. No cross-step re-writes.

Design ready for user approval. Next: on approval, switch to the `superpowers:writing-plans` flow and `superpowers:using-git-worktrees` for implementation.

---

## Implementation Plan (appended 2026-04-22, **restructured 2026-04-22** — reorder + T7 split)

**Status snapshot (post-restructure):**
- ✅ **T1** (foundation layer) shipped in commits `42d6c28` + `6186a3c`. Styles, layout constants, Config columns, reader legacy-migration all live in `main`. `wp_budget_with_indirect` named range exists. `WITH_INDIRECT_HEADER_FILL` / `WITHOUT_INDIRECT_HEADER_FILL` exist.
- ✅ **T2** (stage-1 gate) passed. Tests green; deferred consumers (`summary.py`, `trends.py`, `plan.py`) continue to resolve the old Ceiling-position cell to a numerically identical without-indirect value via the `Budget -i` black-formula bridge.
- ⏭ **Remaining work reordered below.** Timelog Plan (original T7) runs first, split into four sub-tasks (T7.1–T7.4) so each fits comfortably in a 200k-token session context. Per-sheet passes (T3–T6) follow Stage A. Docs / invariants / archive (T8–T10) close out.

**Reordering rationale:** user wants Timelog Plan shipped in the next implementation session. With T1 already in `main`, Timelog Plan has no outstanding blockers — every prerequisite (named range, fills, `Budget +i` column) is live. Splitting T7 into helper-seam sub-tasks keeps each commit under the 200k-token context ceiling even when pytest debug loops engage.

**T7 split — re-work audit:** each sub-task edits a disjoint slice of `budget_planner/timelog_plan.py`:
- **T7.1** (survey) — read-only, doc commit.
- **T7.2** (yearly rows) — `_write_yearly_rows` + `layout.py` Timelog-Plan constants + yearly-row test fixture only.
- **T7.3** (monthly + per-year cols) — `_write_monthly_total_row` + year-col section of `_write_wp_summary_row` + corresponding test fixtures.
- **T7.4** (cleanup + Stage-A gate) — grep sweep only; no edits to regions already rewritten in T7.2/T7.3.

`layout.py` touched once (T7.2). Each builder helper touched once. `tests/test_timelog_plan.py` touched in T7.2 and T7.3 but in disjoint assertion regions. Clean per CLAUDE.md re-work audit rule ("same target, two steps" → merge; disjoint regions in one file → acceptable).

**Execution guidance for `/impag`:**
- T7.1 is read-only + a doc commit — prefer direct orchestrator work (no subagent needed; the output lands verbatim in this plan doc under "T7.1 survey outcome").
- T7.2 / T7.3 are each 3-file tasks (builder helper + layout.py or none + test fixture) — direct edits preferred per CLAUDE.md "small task" rule; the cross-file coherence (column-constant → builder cell-address → test assertion) is too tight for subagent safety.
- T7.4 is ops + a grep sweep — direct edits.
- T3–T6 unchanged in content from the pre-restructure plan, only reordered. Same "no parallelism across sheet builders" rule applies.
- T8 (docs) is 4-file but pure prose — subagent-eligible.
- T9 (invariant tests) is 1 new test file + cross-sheet asserts — direct edits.
- T10 is ops only.

### Stage A — Timelog Plan (T7.1 .. T7.4) — **runs first**

#### T7.1. Pre-impl column-map survey (read-only, doc-only commit)

**Purpose:** D8's Timelog Plan section specifies "`Yearly total €` row and `Yearly budget €` row move to **column H**. The per-year breakdown columns (currently showing yearly numbers) start at column I onward." The semantics of "move to column H" are ambiguous in the current code — col H today holds Rate in PER-PERSON SUMMARY, and the 7th metric in PER-WP SUMMARY (per `timelog_plan.py` L96-97). T7.1 resolves the ambiguity before any code moves, and records the resolution inline in this plan doc.

**Reads (offset+limit — do NOT full-read `timelog_plan.py`):**
- `budget_planner/layout.py` — Timelog Plan column constants (may be absent today; T7.2 may need to add them).
- `budget_planner/timelog_plan.py` L77..L110 — section-title anchors + PER-WP / PER-PERSON SUMMARY metric-col layout documentation.
- `budget_planner/timelog_plan.py` L927..L1017 — `_write_yearly_rows` body.
- `budget_planner/timelog_plan.py` L1018..L1064 — `_write_monthly_total_row` body.
- `budget_planner/timelog_plan.py` L1065..L1174 — `_write_wp_summary_row` (year-col write section).
- `output/budget_plan.xlsx` via `mcp__openpyxl-read` — visually confirm what occupies col H in PER-WP SUMMARY on the last-generated workbook.

**Deliverable (appended to this plan doc, under "T7.1 survey outcome" stub at end):**
- Concrete before/after column map: cols A..H + year cols today vs. after shift — what moves, what stays, what's new.
- Cell-by-cell: which helper writes which cell in the new layout. Include new `Yearly Budget +i` xref formula literal referencing `wp_budget_with_indirect`.
- If D8's "move to column H" spec is ambiguous on arrival, pick a concrete resolution and record it as an inline **D9** addendum in the Decisions log (above).

**Commit:** `docs(plan): T7.1 Timelog Plan column-map survey — before/after + H-shift resolution`

#### T7.2. `_write_yearly_rows` rewrite + layout constants

**Files:** `budget_planner/timelog_plan.py` (`_write_yearly_rows` only), `budget_planner/layout.py` (Timelog-Plan column constants per T7.1 outcome), `tests/test_timelog_plan.py` (yearly-row assertion block only).

**Edits:**
1. Rewrite `_write_yearly_rows` per T7.1 column map:
   - `Yearly total €` label → rename to `Yearly total +i`; apply `WITH_INDIRECT_HEADER_FILL` to the header cell.
   - `Yearly budget €` label → rename to `Yearly Budget +i`; cross-ref Config `wp_budget_with_indirect` named range; **no ×1.25** applied (value is already with-indirect at source).
   - `Yearly total +i` formula: `SUM` of per-year `+i` cols on this row (which are ×1.25-display per T7.3).
   - Cell-address targets updated per T7.1 column map (H-shift).
2. `layout.py`: add/update Timelog-Plan column constants per T7.1 resolution. This file is touched once across all T7 sub-tasks.
3. `tests/test_timelog_plan.py`: update **only** the yearly-row assertion block — header text (`Yearly total +i`, `Yearly Budget +i`), fill (`WITH_INDIRECT_HEADER_FILL`), formula shape (xref where expected; SUM where expected), cell positions (H-shift). Do NOT touch monthly-total or year-col asserts — T7.3 owns those.

**Targeted tests:** `pytest tests/test_timelog_plan.py -k "yearly" -q && pyright budget_planner/timelog_plan.py budget_planner/layout.py`

**Commit:** `feat(timelog-plan): Yearly total/Budget +i rename, H-column shift, Config xref`

#### T7.3. `_write_monthly_total_row` + per-year cols in `_write_wp_summary_row`

**Files:** `budget_planner/timelog_plan.py` (`_write_monthly_total_row` + year-col section of `_write_wp_summary_row`), `tests/test_timelog_plan.py` (monthly-total + year-col assertion blocks).

**Edits:**
1. `_write_monthly_total_row`:
   - Header label → `Monthly total +i`; apply `WITH_INDIRECT_HEADER_FILL`.
   - Formula multiplies the solver's without-indirect per-month total by ×1.25 at the display cell.
2. `_write_wp_summary_row` year-col writes:
   - Year-col headers `Y{yyyy} €` → `Y{yyyy} +i`; apply `WITH_INDIRECT_HEADER_FILL`.
   - Per-WP per-year value formula wraps solver output × 1.25 at the display cell.
3. `tests/test_timelog_plan.py`: update monthly-total + year-col assertion blocks — header text, fill, `×1.25` formula shape.

**Re-work audit:** T7.2 touched `_write_yearly_rows` only. T7.3 touches `_write_monthly_total_row` + `_write_wp_summary_row` (year-col region). No overlap. `tests/test_timelog_plan.py` touched in disjoint assertion regions.

**Targeted tests:** `pytest tests/test_timelog_plan.py -k "monthly or year_col or summary" -q && pyright budget_planner/timelog_plan.py`

**Commit:** `feat(timelog-plan): Monthly total +i + year-col +i display (×1.25 at display cell)`

#### T7.4. Cleanup + Stage-A gate

**Files:** `budget_planner/timelog_plan.py` (grep sweep only), plus ops.

**Edits:**
1. Grep `budget_planner/timelog_plan.py` for `Ceiling` / `ceiling` residue — rename any survivors to `Budget` / `budget` (variables, docstrings, comments, literals). If the grep comes up clean, skip to step 2.
2. `python scripts/recalc.py output/budget_plan.xlsx`.
3. `pytest tests/ -q` — full suite. Confirms no regression on other sheets from the `layout.py` Timelog-Plan-constants edit.
4. Manual LO check — open Timelog Plan and confirm:
   - `Yearly Budget +i`, `Yearly total +i`, `Monthly total +i` render with green fill.
   - Numeric values match the pre-restructure workbook × 1.25 (eyeball a couple of WP rows).
   - Per-year `Y{yyyy} +i` cols also green; values ×1.25 of old display.
5. Update `memory/project_state.md`: "Timelog Plan (+i) shipped; next: per-sheet passes T3–T6 (Actuals / Plan-Forecast / Capacity / Actuals-Hours)."

**Commit:** `chore(timelog-plan): Ceiling→Budget residue sweep` — skip commit if step-1 grep was empty.

**Stage A boundary — next `/impag` session may stop here.**

### Stage B — Per-sheet passes (T3–T6)

Each task: update the sheet builder, update the test fixture in the same commit, run targeted tests + pyright, commit. No cross-task parallelism (same reasoning as the pre-restructure plan).

#### T3. Actuals sheet

**Files:** `budget_planner/actuals.py`, `tests/test_actuals.py`

**Edits:**
- Monthly EUR grid (cols C..C+N−1): wrap in column outline group (openpyxl `ws.column_dimensions[col].outline_level = 1; ...hidden = True`).
- Month-section header row: append `-i` suffix to the section label, orange fill.
- Default-visible rollup columns to the right of the grid — flip to with-indirect:
  - `Sum +i`, `Budget +i` (xref Config `wp_budget_with_indirect`), `Variance +i`, `Cum EUR +i`, `% of Budget +i`, `Breach-In +i`.
  - Green header fill. Black-formula font. Formulas: where the aggregation is a row-sum of without-indirect monthly cells, wrap in `* 1.25`. Where it's a cross-ref to a value already with-indirect (Config Budget +i), no `×1.25`.
- Inside the collapsible band (hidden by default): sibling rollups `Sum -i`, `Budget -i`, `Variance -i`, `Cum EUR -i`, `% of Budget -i`, `Breach-In -i`. Orange header fill. Place these columns *inside* the outlined band so they collapse with the monthly grid.
- Forecast-monitoring block: same `+i` default / `-i` inside-band pattern.
- Rename any `Ceiling` identifier → `Budget` in this file.

**Targeted tests:** `pytest tests/test_actuals.py -q && pyright budget_planner/actuals.py`

**Test fixture updates:** new column header names; outline-group assertions (`assert ws.column_dimensions["C"].outline_level == 1`); `×1.25` formula assertion for `Sum +i`.

**Commit:** `feat(actuals): default +i rollups, -i band inside collapsible monthly grid`

#### T4. Plan-Forecast sheet

**Files:** `budget_planner/plan_forecast.py`, `tests/test_plan_forecast.py`

**Edits:**
- Monthly forecast EUR grid: wrap in column outline group.
- Rollup columns flip to `+i`: `Sum +i`, `Budget +i` (xref Config), `Variance +i`, `Cum EUR +i`, `% of Budget +i`.
- Inside band: `-i` siblings.
- Month-section header: `-i` suffix.
- `Ceiling` → `Budget` rename.

**Targeted tests:** `pytest tests/test_plan_forecast.py -q && pyright budget_planner/plan_forecast.py`

**Commit:** `feat(plan-forecast): default +i rollups, -i band, outline-grouped forecast grid`

#### T5. Capacity sheet

**Files:** `budget_planner/capacity.py`, `tests/test_capacity.py`

**Pre-impl read:** scan `capacity.py` for EUR-bearing columns. Per D8 this sheet is FTE-dominant, so the diff may be rename-only (`Ceiling` → `Budget`) with no `+i`/`-i` work.

**Edits:**
- If EUR columns exist: apply the `+i` default / `-i` in-band pattern per D6.
- If not: rename `Ceiling` → `Budget` wherever it appears (variable names, docstrings, header strings).

**Targeted tests:** `pytest tests/test_capacity.py -q && pyright budget_planner/capacity.py`

**Commit:** `feat(capacity): Ceiling→Budget rename[; +i/-i header suffixing if EUR cols present]`

#### T6. Actuals-Hours sheet

**Files:** `budget_planner/actuals_hours.py`, `tests/test_actuals_hours.py`

**Edits:**
- Per-person hours grid: unchanged (hours, not EUR).
- Forecast-EUR band: header suffix `-i`, orange fill. Wrap the band in a **row** outline group (not column — this band is vertical per person).
- New default-visible `Forecast EUR +i` single rollup column per person: green fill, black formula = `row-sum of forecast-EUR band × 1.25`. Placed outside the outlined band so it stays visible on collapse.
- `Ceiling` → `Budget` rename if any.

**Targeted tests:** `pytest tests/test_actuals_hours.py -q && pyright budget_planner/actuals_hours.py`

**Commit:** `feat(actuals-hours): Forecast EUR +i rollup, -i band row-outline-grouped`

*(Old T7 "Timelog Plan" subsection here has been relocated to Stage A as T7.1–T7.4 — see above.)*

### Stage C — Docs, invariants, archive (T8–T10)

#### T8. Docs sweep

**Files:** `docs/grant-domain.md`, `docs/design-invariants.md`, `docs/internals.md`, `CLAUDE.md`

**Edits:**
- `grant-domain.md`: new section "Indirect-view convention" — explains `+i` / `-i` suffixes, fill color code (green #E2EFDA / orange #FBE5D6), Budget (not Ceiling) terminology, storage convention (A + B exception per D1/D4).
- `design-invariants.md`: add invariant — *"Every EUR-bearing column header on in-scope sheets (Config, Actuals, Plan-Forecast, Capacity, Actuals-Hours, Timelog Plan) carries `+i` or `-i` suffix."* Update Cross-Sheet Reference Map named-range contract table with `wp_budget_with_indirect`.
- `internals.md`: update per-sheet layout notes for Timelog Plan H-column shift (before/after column map). Update Cross-Sheet Reference Map to reflect new named range.
- `CLAUDE.md` Reference Index: update any entries that mention "Ceiling" (reference index row for Config / WP Budget). Rename pointer descriptions.
- **All three docs:** insert a deferral note — *"Trends, Summary, and Plan intentionally not updated in this pass; their terminology will be aligned in the follow-on overhaul plan."*
- Rename all `Ceiling` → `Budget` across `docs/**` (grep sweep).

**Targeted tests:** `pytest tests/test_config_sheet.py::test_config_named_range_targets_match_internals_doc -q` — this test parses `docs/internals.md`, so restructuring the Cross-Sheet Reference Map table breaks pytest. Run it before commit. Full `pytest tests/ -q` as a docs→test-coupling safety check.

**Commit:** `docs: indirect-view convention, +i/-i invariant, Ceiling→Budget sweep`

#### T9. Cross-sheet invariant tests

**Files:** new `tests/test_indirect_view_invariants.py` (or extend an existing invariants test file if present — grep for `test_invariants` first)

**Tests to add:**
1. `test_every_eur_header_carries_suffix` — iterate all in-scope sheets (6), scan header cells with a font or named-range-based EUR classifier, assert each ends with `+i` or `-i`.
2. `test_budget_plus_i_minus_i_round_trip` — for every WP, assert `Budget +i / 1.25 ≈ Budget -i` within `1e-6`.
3. `test_header_fill_matches_suffix` — `+i` headers have `WITH_INDIRECT_HEADER_FILL`; `-i` headers have `WITHOUT_INDIRECT_HEADER_FILL`. Non-EUR headers still have `LIGHT_BLUE_FILL`.

**Targeted tests:** the new file itself, plus full `pytest tests/ -q`.

**Commit:** `test: indirect-view cross-sheet invariants (header suffix, fill, budget round-trip)`

#### T10. Final gate + archive

**Ops only.**
1. `python scripts/recalc.py output/budget_plan.xlsx`
2. `pytest tests/ -v` — full suite.
3. **Manual LO check** — open workbook:
   - Config `WP BUDGET`: Budget +i green/input, Budget -i orange/formula
   - Actuals: collapse the `–` toggle, monthly grid + `-i` rollups hide; expand, both visible
   - Plan-Forecast: same toggle check
   - Actuals-Hours: row outline collapses forecast-EUR band, `Forecast EUR +i` stays visible
   - Timelog Plan: columns H, I... laid out correctly; `Yearly Budget +i`, `Yearly total +i`, `Monthly total +i` green-filled
4. `git mv docs/plans/2026-04-22-indirect-view-restructure-design.md docs/plans/archived/`
5. Grep `2026-04-22-indirect-view-restructure-design` across `budget_planner/**`, `tests/**`, `docs/**`, `CLAUDE.md`, `.claude/rules/**` — update any references to point at `docs/plans/archived/`.
6. Update `memory/project_state.md` — mark indirect-view restructure shipped; clear the "In-flight plans" entry; if user has not yet written the Trends/Summary follow-on overhaul plan, leave the nice-to-have pointer.
7. Verify `.claudeignore` names `archived/` (not `archive/`).

**Commit:** `chore: archive indirect-view restructure plan post-ship`

### Task dependency summary (post-restructure)

```
✅ T1 ─ ✅ T2 ─┬─ T7.1 ─ T7.2 ─ T7.3 ─ T7.4 ─┬─ T3 ─ T4 ─ T5 ─ T6 ─┬─ T8 ─ T9 ─ T10
               │                              │                     │
               └─ shipped (commits            └─ Stage-A boundary   └─ Stage-B boundary
                   42d6c28 + 6186a3c)
```

- T1 + T2: ✅ shipped. Foundation styles / layout / Config / reader-migration live in `main`; Stage-1 gate passed.
- T7.1 → T7.2: survey resolves H-column ambiguity and records the column map inline in this doc before any code moves.
- T7.2 → T7.3: disjoint helpers (`_write_yearly_rows` vs `_write_monthly_total_row` + `_write_wp_summary_row` year cols); sequential only for one-file-at-a-time pytest discipline.
- T7.3 → T7.4: cleanup and Stage-A gate.
- T7.4 → T3: Timelog Plan stable before per-sheet sweeps add new outline-group and header-suffix patterns to sibling builders.
- T3 → T4 → T5 → T6: serial for simplicity (same reasoning as the pre-restructure plan — T3/T4 share the collapsible-band pattern).
- T6 → T8: docs describe the final state; wait for all sheet builders to land before writing the before/after column maps in `internals.md`.
- T8 → T9: invariant tests assert things described in `design-invariants.md`; keep the doc landing ahead of the test it codifies.
- T9 → T10: final gate runs the full suite including the new invariants.

### Pre-ship checklist

- [x] D7 confirmation recorded in Decisions log — 2026-04-22
- [x] Person-name sweep in this doc done — 2026-04-22
- [x] `feedback_stakeholder_naming.md` rule tightened to cover all persons — 2026-04-22
- [x] `version_stamp` bumped in `config_sheet.py` — T1 (commit `42d6c28`)
- [x] `wp_budget_with_indirect` named range landed in Config — T1 (commit `42d6c28`)
- [x] `WITH_INDIRECT_HEADER_FILL` + `WITHOUT_INDIRECT_HEADER_FILL` added to `styles.py` — T1 (commit `42d6c28`)
- [x] **T7.1** column-map survey outcome appended to this doc (before/after + H-shift resolution + optional D9 addendum) — 2026-04-22
- [x] Timelog Plan H-column shift landed (T7.2) with `layout.py` Timelog-Plan constants — 2026-04-22
- [ ] Timelog Plan `+i` rename landed (T7.2 yearly + T7.3 monthly + year-cols) and full pytest green (T7.4)
- [ ] Reader migration env-flag override (`GRANT_ACCEPT_CEILING_MIGRATION`) documented in `docs/ops-quickref.md` failure-mode table
- [ ] Archive `git mv` done with `.claudeignore` verified (`archived/`, not `archive/`)

---

## T7.1 survey outcome

*Landed 2026-04-22 under `/impag` Stage A.*

### Current column layout (verified against `budget_planner/timelog_plan.py` @ `main`)

**PER-WP SUMMARY band** (`_write_wp_summary_row`, L1065–L1173; headers at L99–L103 — `_WP_SUMMARY_BASE_HEADERS`):

| Col | # | Current header | Content |
|-----|---|----------------|---------|
| A | 1 | `WP code` | literal WP code (black-formula font, centered) |
| B | 2 | `Ceiling €` | green xref `=Config!E{cfg_wp_row}` — resolves to Config col 5 = `Budget -i` (post-T1, numerically unchanged) |
| C | 3 | `Assigned €` | `=ROUND(SUMPRODUCT((grid="WP")*rate_vec),0)` — without-indirect (rate is -i) |
| D | 4 | `Buffer €` | `=C{row}-B{row}` |
| E | 5 | `Assigned months` | `=SUMPRODUCT(1*(grid="WP"))` |
| F | 6 | `WP sched months` | literal `end−start+1` |
| G | 7 | `Burn rate needed` | `=IF(AND(D<0,wallclock>0),-D/wallclock,"")` |
| H | 8 | `Any out-of-window?` | `=SUMPRODUCT((grid="WP")*((midx<start)+(midx>end)))` |
| I..  | 9.. | `Y{yyyy} €` | per-year `=ROUND(SUMPRODUCT((year_subgrid="WP")*rate_vec),0)` — without-indirect |

**Yearly total € / Yearly budget € rows** (under PER-WP SUMMARY; `_write_yearly_rows`, L927–L1015):

| Col | Yearly total € | Yearly budget € |
|-----|----------------|-----------------|
| A | `Yearly total €` (label, HEADER_FONT) | `Yearly budget €` (label, HEADER_FONT) |
| B | blank | blank |
| C | `=SUM(C{first}:C{last})` — across-WP Assigned € cross-check | `=SUM(B{first}:B{last})` — across-WP Ceiling € cross-check |
| D–H | blank | blank |
| I..  | `=SUM(I{first}:I{last})` per year-col | `=Σ prop_ky × Config!E{wp_row}` (literal proportions × Config col 5 = Budget -i) |

Rows are placed directly below the last PER-WP SUMMARY WP row: `yearly_total_row = wp_summary_last_row + 1`, `yearly_budget_row = wp_summary_last_row + 2` (L238–239). Same outline group as PER-WP SUMMARY band.

**Monthly total € row** (under TIMELOG GRID; `_write_monthly_total_row`, L1018–L1062):

| Col | Content |
|-----|---------|
| A (1) | `Monthly total €` (label, HEADER_FONT) |
| B..(month_col_end) | per-month `=ROUND(SUMPRODUCT((col<>"")*rate_vector),0)` — without-indirect; vacation months carry `VACATION_FILL` (formula still evaluates to 0) |

Row placement: `grid_total_row = grid_person_last_row + 1` (L257), inside the TIMELOG GRID outline group.

**PER-PERSON SUMMARY** (not touched by T7 — included only for completeness):

| Col | Content |
|-----|---------|
| A (1) | Person name |
| B–G (2–7) | 6 metric cols (`Assigned`, `Pool`, `Over-pool?`, `Switches`, `First log`, `Last log`) |
| H (8) | Rate — single source, `_PERSON_SUMMARY_RATE_COL = 8` (L92–94), consumed by PER-WP Assigned-€ SUMPRODUCT |

### D9 addendum — "move to column H" resolution

**D8 quote:** *"`Yearly total €` row and `Yearly budget €` row move to **column H**. The per-year breakdown columns (currently showing yearly numbers) start at column I onward."*

**Decoded:**
- The **per-year breakdown columns already start at col I today** (`_WP_SUMMARY_BASE_LAST_COL + 1 + yi = 8 + 1 + yi`, so first year-col = col 9 = I). D8's second sentence is a *confirmation of the current state*, not a change.
- The actual change is: **the row label moves from col A to col H**, placing the `Yearly total +i` / `Yearly Budget +i` label immediately adjacent (left) to the per-year breakdown cols — improves readability since the reader's eye no longer has to traverse blank cols B..G between label and values.

**Concrete resolution adopted:**

1. **Label cell moves A → H** on both `yearly_total_row` and `yearly_budget_row`. Col A becomes blank on these rows.
2. **Col C cross-checks stay at col C** — they are "across-WP SUMs of col B / col C" and sit under their own column header semantically. Moving them would orphan them from the header they aggregate. They remain a formula-drift detector (col C cross-check vs. SUM-of-year-cols should agree up to rounding).
   - *Semantic note:* T7.2 retargets `yearly_budget_row` col C cross-check from `=SUM(B{first}:B{last})` (sums PER-WP SUMMARY col B = Ceiling €, currently -i via Config!E) to a formula that yields the +i total. Two options for T7.2 to choose between:
     - (a) keep `=SUM(B{first}:B{last})` and accept the result is -i (divergent from the row's +i per-year sums).
     - (b) change to `=SUM(B{first}:B{last})*1.25` for +i parity (preferred; lowest-impact formula tweak, preserves the cross-check's drift-detection purpose).
   - *T7.2 picks option (b)* — see T7.2 edit list below. Analogous for `yearly_total_row` col C: `=SUM(C{first}:C{last})*1.25` for +i parity (see next point for why PER-WP SUMMARY col C stays -i).
3. **PER-WP SUMMARY col B/C/D stay at their current semantics for T7 scope.** Col B `Ceiling €` and col C `Assigned €` and col D `Buffer €` (per-WP scalar rollups) are **not listed** in D8's Timelog Plan EUR-flip bullet (*"all EUR cells on Timelog Plan (per-year EUR cols, `Monthly total €`, `Yearly total €`, `Yearly budget €`) display with-indirect"*). T7 leaves these three columns at their current header text and -i semantics. A later pass (out-of-scope for this plan, candidate for the follow-on Trends/Summary overhaul or a dedicated Timelog-Plan polish) can flip them if desired.
   - *Consequence:* the +i display surface on Timelog Plan after T7 is: per-year cols (I..), `Monthly total +i`, `Yearly total +i`, `Yearly Budget +i`. PER-WP SUMMARY base cols A–H remain -i-or-neutral. Acceptable per D8's explicit EUR-cell enumeration.
4. **Monthly total € row label stays at col A.** D8 only names "Yearly total € row and Yearly budget € row" as moving to col H; Monthly total is under TIMELOG GRID (separate section), its col A label sits cleanly to the left of the grid's month cols (B onward), and shifting it has no readability benefit. Kept at col A.

### T7.2 edit list (derived from column map)

`_write_yearly_rows` rewrite in `budget_planner/timelog_plan.py`:

- **Remove** the col-1 label writes (L954–L961).
- **Add** label writes at col 8 (`_WP_SUMMARY_BASE_LAST_COL`): `Yearly total +i` on `yearly_total_row`, `Yearly Budget +i` on `yearly_budget_row`. Font: `HEADER_FONT`. Fill: `WITH_INDIRECT_HEADER_FILL` (import from `styles`). Alignment: `CENTER`.
- **Retarget col C cross-check** (L964–L969):
  - `yearly_total_row` col 3 → **unchanged** `=SUM(C{first}:C{last})` in T7.2. Per-year cells on this row sum PER-WP SUMMARY year-cells, which are -i until T7.3's ×1.25 wrap at the leaf SUMPRODUCT. Both sides stay -i in T7.2; T7.3 adds the ×1.25 wrap on col C at the same time as the leaf wrap.
  - `yearly_budget_row` col 3 → `=SUM(B{first}:B{last})*1.25` (+i parity). Per-year cells on this row flip to +i in T7.2 via the Config E→F retarget below, so col C must flip simultaneously to keep the drift detector valid.
  - Keep `BLACK_FORMULA` + `fmt="€#,##0"`.
  - *Asymmetry discovered during T7.2 impl (2026-04-22):* yearly_total and yearly_budget rows have different data paths — yearly_total aggregates PER-WP SUMMARY year cells (owned by T7.3); yearly_budget aggregates Config Budget cells (owned by T7.2). They cannot flip together.
- **Per-year col formulas** (L984–L990, L992–L1015):
  - `yearly_total_row` per-year: `=SUM(<yearcol>{first}:<yearcol>{last})` unchanged in shape — the underlying year cells become +i via T7.3, so the SUM is automatically +i.
  - `yearly_budget_row` per-year: change `cell_ref(SHEET_CONFIG, cfg_wp_first_row + wi, 5)` → `cell_ref(SHEET_CONFIG, cfg_wp_first_row + wi, COL_WP_BUDGET_PLUS_I)`. Import `COL_WP_BUDGET_PLUS_I` from `budget_planner.layout`. Formula becomes `=Σ prop_ky × Config!F{wp_row}` — pulls Budget +i directly, no ×1.25 multiplication needed (already +i).
- **Docstring update:** L938–L953 docstring says "Col C: cross-check sum — total Assigned € (yearly_total_row) or total Ceiling € (yearly_budget_row)". Update to reflect ×1.25 wrap and new Config xref target (col F / Budget +i).

`layout.py` additions (single touch per T7 split rule):

- `TIMELOG_WP_SUMMARY_BASE_LAST_COL = 8` — mirror of `_WP_SUMMARY_BASE_LAST_COL` currently local to `timelog_plan.py`. Adding to `layout.py` satisfies T7.2's layout-constants deliverable; tests can assert against this constant.
- `TIMELOG_YEARLY_LABEL_COL = 8` — the post-shift label column. Expresses the H-shift explicitly as a layout decision (not buried in the helper).
- Leave `_WP_SUMMARY_BASE_LAST_COL` in `timelog_plan.py` as a module-local alias (`= TIMELOG_WP_SUMMARY_BASE_LAST_COL`) for internal callers, or inline replacement — T7.2 to decide.

`tests/test_timelog_plan.py` yearly-row assertion block updates:

- Header text: `Yearly total €` → `Yearly total +i`; `Yearly budget €` → `Yearly Budget +i`.
- Label cell: assert at col H (col 8), not col A.
- Fill: assert `WITH_INDIRECT_HEADER_FILL` on the label cell.
- Cross-check formula: assert `*1.25` wrap.
- Per-year `yearly_budget_row` formula: assert Config xref targets col F (Budget +i), not col E.
- Do **not** touch monthly-total or year-col assertions — T7.3 owns.

### T7.3 edit list (derived from column map)

`_write_monthly_total_row`:
- Header text: `Monthly total €` → `Monthly total +i`. Apply `WITH_INDIRECT_HEADER_FILL` to the label cell.
- Per-month formula (L1060–L1062): wrap in ×1.25 at the display cell — `=ROUND(SUMPRODUCT((col<>"")*rate_vec)*1.25,0)`. Keep `ROUND(..., 0)` per openpyxl.md EUR rule.

`_write_wp_summary_row` year-col region (L1151–L1172):
- Header for year cols: `Y{yyyy} €` → `Y{yyyy} +i`. Apply `WITH_INDIRECT_HEADER_FILL` to year-col header cells (year-col header write lives in the header-row setup at L339, not in `_write_wp_summary_row` — T7.3 also touches that header-row site).
- Year-col per-WP formula (L1167–L1170): wrap in ×1.25 — `=ROUND(SUMPRODUCT((year_subgrid="WP")*rate_vec)*1.25,0)`.

Tests: monthly-total + year-col assertion blocks get `+i` / `×1.25` updates.

### Risks / uncertainties surfaced

1. **Col C cross-check semantic divergence** — after T7.2 wraps cross-check in ×1.25 but PER-WP SUMMARY col B/C stay -i, the col C cells on yearly rows display a number that equals `across_WP_sum × 1.25`, not equal to the PER-WP SUMMARY col it sits under. Mitigation: the cell is a drift-detector, not a reader-facing aggregate — and the formula is self-documenting. Acceptable; note in docstring.
2. **`Ceiling €` label on PER-WP SUMMARY col B (L99, L1090–L1092)** — still uses the word "Ceiling" both in the header string literal and in a docstring comment. T7.4 grep sweep catches these.
3. **Docstring at L23–L25** references "col H" of PER-PERSON SUMMARY as the Rate col. After T7 shifts labels to col H of PER-WP SUMMARY (a *different* band), the docstring stays correct for PER-PERSON SUMMARY — but T7.2 should add a clarifying line distinguishing the two col-H uses across bands.

### Commit message for T7.1

`docs(plan): T7.1 Timelog Plan column-map survey — D9 H-shift resolution + T7.2/T7.3 edit lists`


---

## Correction — T3 re-touched `layout.py` (2026-04-23)

**Context:** the original plan specified *"`layout.py` lands once in T1"* (Stage 1). T1 shipped the foundation: `WITH_INDIRECT_HEADER_FILL`, `WITHOUT_INDIRECT_HEADER_FILL`, `COL_WP_BUDGET_MINUS_I`, `COL_WP_BUDGET_PLUS_I`, and the `wp_budget_with_indirect` named range.

**Discovered during T3:** the `val_columns()` / `forecast_val_columns()` helpers that key the Actuals rollup-block geometry are Actuals-specific and couldn't be expanded at T1 — the +i offsets (7..12) plus the forecast +i keys (GUTTER+5..+8) only make sense once Actuals is the sheet consuming them. Deferred consumers (Summary, Trends, Plan) still read -i values at offsets 1..6, which is preserved.

**Resolution:** T3 added the following to `budget_planner/layout.py`:
- `VAL_OFFSET_SUM_PLUS_I` .. `VAL_OFFSET_BREACH_PLUS_I` (7..12) — offsets for the +i mirror rollups.
- Six new keys on `val_columns()`: `sum_plus_i`, `budget_plus_i`, `diff_plus_i`, `cum_plus_i`, `pct_ceil_plus_i`, `breach_plus_i`.
- Four new offsets + keys on `forecast_val_columns()`: `cum_forecast_plus_i`, `var_forecast_plus_i`, `pct_ceil_forecast_plus_i`, `breach_forecast_plus_i`.
- `forecast_val_columns()` re-based on `val_columns(..)["breach_plus_i"]` (was `breach`). Legacy `cum_forecast` / `var_forecast` / `pct_ceil_forecast` / `breach_forecast` keys now point at -i siblings in the new base band. Adjacency (`var_forecast == cum_forecast + 1` etc.) preserved within the -i sub-block, so `test_actuals_forecast_block.py` continues to pass unchanged.

**No downstream breakage:** `config.py` derives `VAL_COL_SUM/BUDGET/DIFF/CUM/PCT_CEIL/BREACH` via the -i keys, unchanged. Other builders (`plan_forecast.py`, `plan.py`, `capacity.py`, `actuals_hours.py`) read `_vcols["sum"]` only and keep working. No migration path hit.
