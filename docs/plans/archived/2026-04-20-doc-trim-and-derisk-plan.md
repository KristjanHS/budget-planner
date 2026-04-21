# Doc Trim & De-risk Plan — 2026-04-20

**Status:** v2 — user picks locked in (2026-04-20). Two-PR rollout.
**Informed by skills:** `brutal-honesty-review` (bloat calls), `condense` (dedup), `claude-md-progressive-disclosurer` (CLAUDE.md tiering), `kaizen:analyse-problem` (A3 shape), `mybrain` (option selection).

## Decisions (v2)

User picks 2026-04-20:
- **1C** — no new `layout-reference.md`; delete `internals.md` §Layout outright, CLAUDE.md points at `budget_planner/layout.py`.
- **2C** — no new `timelog-plan.md`; trim duplicates, CLAUDE.md points at archived plans + code.
- **3C** — no new `actuals.md`; `design-invariants.md` is canonical, trim `internals.md` duplicate.
- **4C** — delete `directory-map.md` entirely.
- **5 (hybrid)** — trim and split: AI-facing summary in loadable tier; full human runbooks behind `.claudeignore`.
- **6B** — paragraph in `dev-workflow.md`; soft enforcement. **+ new B0: strip prose/incident bloat from dev-workflow.md.**
- **7B** — replace Large Files table with `wc -l` command + threshold rule.
- **8A** — single `.claude/rules/openpyxl.md` with subsection headings; trim prose.
- **9B** — two PRs (Part A files, Part B process).

## Problem (unchanged from v1)

1. **Drift-prone helper docs** — live `docs/*.md` prose enumerates column offsets, row indices, cell coordinates that duplicate `layout.py` and builders. Cross-edit burden on every layout change; silent drift when missed.
2. **Per-feature doc trail** — "design → plan → ship → archive" produces 2–4 markdown files per feature (26 archived in 14 days). Live rule files accrete dated-incident footnotes.

## Target

- Delete duplicative live prose; code + archived plans + one canonical invariants doc are enough.
- Rules state rules, not incidents. Incident details live in git log.
- CLAUDE.md's "Large Files" table is grep-derived.
- Human runbooks (bash scripts, reproducer workflows) sit behind `.claudeignore` with a thin AI-loadable summary.

---

## Part A: Files (one PR)

### A1. Delete `docs/internals.md` §"Layout and Styling" (lines ~179–246)

- Replace the section with one line: `> Layout constants live in budget_planner/layout.py. Row/column positions and offsets are not duplicated here to avoid drift.`
- Update CLAUDE.md Reference Index row "Column offsets, month positioning..." to point at `budget_planner/layout.py` + a one-line note about freeze panes (or drop the row if the file header is self-explanatory).
- **Grep sweep:** `docs/**/*.md` for `MONTH_COL_START`, `Row \d`, `Col \d` in prose; remove or rewrite each hit.

### A2. Trim Timelog Plan duplication

- `docs/internals.md` §Timelog (~lines 42–43): replace with a 1-line pointer: `> Timelog Plan builder — see timelog_plan.py module docstring + docs/plans/archived/2026-04-18-timelog-plan-design.md (v1) + 2026-04-19-timelog-plan-v2-design.md (v2) + 2026-04-19-timelog-plan-enhancements-design.md (phases 3–4).`
- `docs/design-invariants.md` §"Timelog Plan" (~lines 49–105): keep **only** the cross-sheet invariants (toggle cell location, vacation-cell trust rule). Move solver + grid implementation details into `timelog_solver.py` + `timelog_plan.py` module docstrings. Target length ~25 lines, down from 56.
- CLAUDE.md trigger row updated to reference: `timelog_plan.py + timelog_solver.py + archived plans`.

### A3. Trim Actuals duplication

- `design-invariants.md` §"Two-cadence actuals" (~lines 5–47): stays as canonical home, unchanged.
- `docs/internals.md` Actuals + Actuals-Hours notes: replace with 1-line pointer to design-invariants.md.

### A4. Absorb `docs/wp-schedule-type-suggestions.md` (23 lines)

- Add as a table in `docs/grant-domain.md` (natural home — WP type semantics are domain rules).
- `git rm docs/wp-schedule-type-suggestions.md`.
- No CLAUDE.md row for this file currently; nothing to sweep there.

### A5. Human-runbook split (hybrid — per user pick 5)

Goal: AI load tier stays lean; humans still have full docs; nothing duplicated.

**New directory:** `docs/human/` (ops runbooks, human workflows).
**`.claudeignore` addition:** one line: `docs/human/`.

- `git mv docs/cloud-env-setup.md docs/human/cloud-env-setup.md` (96 lines of bash pre-flight + failure-mode table — human runbook).
- `git mv docs/debugging-your-real-file.md docs/human/debugging-your-real-file.md` (54 lines of `make_reproducer.py` workflow — human workflow).
- Keep `docs/dev-workflow-review.md` in place (43 lines, already AI-targeted on-demand via CLAUDE.md trigger). Dedup with `dev-workflow.md` §Column/Row shifts (see B0) — keep rule in one file only.

**New `docs/ops-quickref.md`** (≤30 lines, AI-facing, loadable):
- LibreOffice is required (1 line + failure-mode #1).
- Failure modes 2, 3, 4 as a 4-line table (symptom → one-word fix).
- Reproducer: `python -m scripts.make_reproducer <file>` — 2 lines. Never load the real workbook.
- Pointer: "Full runbooks for humans in `docs/human/` (not auto-loaded)."

**CLAUDE.md sweep:** replace 2 old reference-table rows ("Cloud/container env..." and "Debugging against user's confidential workbook") with 1 row pointing at `docs/ops-quickref.md`.

### A6. Delete `docs/directory-map.md` (83 lines)

- `git rm docs/directory-map.md`.
- Remove its CLAUDE.md Reference Index row ("Directory map (function → file)").
- **No replacement needed.** `ls budget_planner/` + module docstrings cover orientation; `architecture.md` covers build order.

### A7. Part A pointer sweep

After A1–A6 land, sweep CLAUDE.md Reference Index (lines 7–25) one time:
- Every row pointing at a merged/deleted file → rewritten or removed.
- No dangling links to files that no longer exist.

---

## Part B: Process (second PR)

### B0. Trim dev-workflow.md bloat (NEW — per user critique request)

Current state: 52 lines / 9.7KB dense prose, loads on every `.py` edit via `paths: "**/*.py"`.

**Trim operations:**
1. Strip dated-incident footnotes from lines 15, 31, 32, 36, 42, 44, 45, 50. Keep rules, drop "2026-04-19: ..." citations. Git log is the audit trail.
2. Line 45 grep target `docs/internals.md` becomes stale after A1 — rewrite to `budget_planner/layout.py` + any doc citing row numbers.
3. Line 31 pytest-timeout rule: compress to one line ("Pass `timeout: 300000` for full pytest runs; the suite exceeds Bash's default.").
4. Line 50 version-stamp row-shift history: delete entirely (pure git history, not a rule).
5. Dedup with `dev-workflow-review.md`: column-shift grep cascade appears in both. Keep in `dev-workflow.md` only; `dev-workflow-review.md` cross-refs it.

**Path-gating split:** the Plan Hygiene + Docs sections (lines 34–45) fire on every `.py` edit but only concern doc/plan work. Extract into a new `.claude/rules/plan-hygiene.md` with `paths: ["docs/**/*.md", "**/CLAUDE.md", "docs/plans/**/*.md"]`. Narrower gate, lower per-edit cost.

**Target:** `dev-workflow.md` ~30 lines (down from 52); `plan-hygiene.md` ~15 lines (new, narrowly gated).
**Estimated saving:** ~40% load cost on every `.py` edit.

### B1. design→plan single-file convention (soft)

Add one paragraph to `.claude/rules/dev-workflow.md` (after trim):
> Per-feature design + plan + later phases go in **one** dated file under `docs/plans/`, appended in sections: *Design*, *Plan*, *Corrections*, *Phase N*. No `-v2-` / `-enhancements-` siblings — amend in place with a changelog entry. `mybrain` + `impag` skills both write into this one file; the archive rule (`git mv` after ship) is unchanged.

No hook, no mybrain/impag skill patch — soft rule enforced by review. Revisit if churn continues.

### B2. Consume-before-archive checklist (from v1, carried forward)

Add to `.claude/rules/dev-workflow.md` (after trim):
> Before `git mv`ing a design doc to `archived/`, grep §Design content against `docs/internals.md`, `docs/design-invariants.md`, `docs/grant-domain.md`. Any load-bearing fact not present in a living doc is merge-debt — migrate first.

### B3. Replace CLAUDE.md "Large Files" table

Current: 10 hand-maintained rows with line counts.
Replace with one line:
> Files ≥800 lines trigger large-file discipline (prefer Grep or `Read(offset, limit)`). Find them on demand: `wc -l budget_planner/*.py tests/*.py | sort -rn | head`.

### B4. Two `.claude/rules/openpyxl*.md` files — revised from pick 8A → 8A'

**Original pick 8A** — single merged file. **Revision (2026-04-20 iteration 2):** rule-gating audit (see B7) found infra vs builder patterns want different path gates. Splitting into two files keeps narrow-gating discipline; still a net reduction (3 → 2 files).

**File 1: `openpyxl.md`** (merges patterns + charts + readers §consumer-side content)
- `paths:` = union of sheet builders (`plan.py`, `summary.py`, `capacity.py`, `actuals.py`, `actuals_hours.py`, `plan_forecast.py`, `trends.py`, `config_sheet.py`, `workbook.py`). **Remove `styles.py`** (over-gate — styles emits constants only).
- Subsections: `## Core patterns (all builders)`, `## Charts (trends.py only)`.
- Content: CF, formula-eval patterns, layout, `cell_ref` absolute, money ROUND, cross-sheet refs, MATCH/LOOKUP fix with SUMPRODUCT pattern.

**File 2: `openpyxl-infra.md`** (NEW)
- `paths: ["budget_planner/workbook.py", "budget_planner/reader.py", "budget_planner/config_sheet.py", "scripts/**/*.py"]`.
- Subsection: `## Workbook / Reader / Named-range infrastructure`.
- Content: `wb.active` None guard, `ws.cell` keyword form, `insert_rows` + DefinedName auto-extend, `DefinedName.attr_text` quoted-sheet-name rule, protection-state retention, reader type-guards (lines 9–11 of current `openpyxl-readers.md`).
- **Gate fix:** includes `reader.py` which was missing from current patterns.md — DefinedName rule was under-gated.

**Prose trim while merging:**
- CF rule-count churn paragraph → 1 line.
- MATCH/LOOKUP fix: keep rule + reference formulas, drop `2026-04-19` incident citation.
- `cell_ref` absolute default: keep rule, drop incident.
- Cross-sheet row-map plumbing "smell" (line 41): move to `dev-workflow.md` §Design or drop — architectural advice, out of place in openpyxl patterns.

**Targets:** `openpyxl.md` ~40 lines; `openpyxl-infra.md` ~15 lines. Net 3 files → 2 files; ~30% prose reduction.

### B7. Additional rule-gating audit (iteration 2 — 2026-04-20)

Per user request "find more candidates for splitting into narrower-gated ones." Additional extractions from `dev-workflow.md` beyond B0:

**Extract from `dev-workflow.md` (currently `paths: "**/*.py"`):**

| New file | Path gate | Lines lifted from dev-workflow.md | Rationale |
|----------|-----------|-----------------------------------|-----------|
| `.claude/rules/testing-project.md` (NEW) | `["tests/**", "**/test_*.py"]` | Lines 28–32 (6 testing rules: no hardcoded seed values, round-trip reader/builder test, pytest timeout, docs-only exemption) | Rules reference test-only constructs; currently fires on every builder edit. |
| `.claude/rules/rule-authoring.md` (NEW) | `[".claude/rules/**/*.md"]` | Line 44 (grep shipped-fix form, not design-doc form, when adding rule examples) | Fires only when editing a rule file. Currently never fires on that path since gate is `*.py`. |
| `.claude/rules/plan-hygiene.md` (from B0) | `["docs/**/*.md", "**/CLAUDE.md", "docs/plans/**/*.md", ".claude/rules/**/*.md"]` | Lines 34–37 (plan hygiene) + lines 40–42 (docs/bucket migration) | Already proposed in B0; confirm gate includes `.claude/rules/**` per rule-authoring overlap. |
| Move lines 43 + 51 (font-color / Rule 3 enforcement) | Into merged `openpyxl.md` §Core | — | Rules apply when `style_cell()` / `font=...` is being written, i.e. builder paths. Tests/scripts don't emit styled cells. |

**After B7, `dev-workflow.md` contents:**
- Build + Test (core: LibreOffice dep, recalc, no-overwrite, formula-probe, `data_only` diagnostic, stale editable-install, output/*.xlsx commit rule, editor-config commit rule).
- Linting (pyright sweep, subagent dead code, import-resolution masking).
- Formatting (preserve cell formatting).
- Invariants (code-review grep-first, rebuild config flow). Version stamp history deleted (git log). Font color rule moved to `openpyxl.md`.

**Target:** `dev-workflow.md` ~18 lines (down from 52 — ~65% reduction on the file that fires per-`.py`).

**Net rule-file count after all of Part B:**
- `dev-workflow.md` (trimmed, narrow-gated by content, still `**/*.py`)
- `testing-project.md` (NEW, `tests/**`)
- `plan-hygiene.md` (NEW, docs + plans + rules)
- `rule-authoring.md` (NEW, `.claude/rules/**`)
- `openpyxl.md` (merged from 3, builder paths)
- `openpyxl-infra.md` (NEW, workbook + reader + scripts)

**6 rule files, 4 of them narrowly gated.** Current state: 4 rule files, 1 (dev-workflow.md) over-broadly gated at `**/*.py`. File count grows by 2, but per-edit token load drops substantially — most `.py` edits now load 1 file (dev-workflow trimmed) instead of 1 file (dev-workflow bloated).

### B5. Memory files — no action

Current 8 memory files (176 lines total) are lean and well-scoped. No changes.

### B6. Archived plans — no action

26 archived plans in 14 days; `.claudeignore` already excludes `archived/`. No harm.

---

## Part C: Active plan hygiene (carried from v1)

`docs/plans/2026-04-19-wp-gap-removal-design.md` (34KB, ready-to-implement): **not in scope**. But flag in its PR: absorb §Data-model changes into `docs/grant-domain.md` + §Config layout into `docs/internals.md` **before** `git mv` to archive (per B2).

---

## PR shape (per pick 9B)

**PR 1 (Part A — files):**
1. A4 (absorb wp-schedule-type-suggestions).
2. A2 + A3 (trim internals.md duplicates; tighten design-invariants.md Timelog section).
3. A1 (delete internals.md Layout section; CLAUDE.md sweep row).
4. A5 (`docs/human/` dir; `.claudeignore`; new `ops-quickref.md`; CLAUDE.md row swap).
5. A6 (delete directory-map.md; CLAUDE.md row removal).
6. A7 (final CLAUDE.md Reference Index sweep).

**PR 2 (Part B — process):**
1. B0 (trim dev-workflow.md — strip incident citations, version history, compress pytest timeout).
2. B7 (extract testing-project.md, rule-authoring.md, plan-hygiene.md from dev-workflow.md; move font-color rule to openpyxl.md).
3. B1 + B2 (append paragraphs to trimmed dev-workflow.md).
4. B3 (Large Files table swap).
5. B4 (merge 3 openpyxl rule files into 2: `openpyxl.md` + `openpyxl-infra.md`, with prose trim; fix under-gating of DefinedName rule).

## Acceptance criteria

- [ ] `docs/*.md` file count drops by ≥3 (commands was not in scope — re-decided, kept; A4, A6, A5-moved = 3+).
- [ ] No live doc enumerates column offsets or row indices in prose (grep `docs/*.md` for `MONTH_COL_START`, `Row \d`, `Col \d` returns empty).
- [ ] CLAUDE.md Reference Index has no dead links.
- [ ] `.claude/rules/` contains 6 files: `dev-workflow.md` (trimmed), `testing-project.md`, `plan-hygiene.md`, `rule-authoring.md`, `openpyxl.md`, `openpyxl-infra.md`.
- [ ] Every rule file's `paths:` frontmatter gates narrower than `**/*.py` EXCEPT `dev-workflow.md` (core cross-cutting rules only).
- [ ] `.claudeignore` excludes `docs/human/`.
- [ ] `dev-workflow.md` ≤18 lines of rules (~65% reduction from 52).
- [ ] `openpyxl.md` ≤40 lines; `openpyxl-infra.md` ≤15 lines.
- [ ] No rule file contains `2026-\d{2}-\d{2}:` incident citations (git grep).
- [ ] `openpyxl-infra.md` `paths:` includes `reader.py` (fixes current under-gating of DefinedName rule).
- [ ] `openpyxl.md` `paths:` does NOT include `styles.py` (removes over-gate).
- [ ] Full test suite passes unchanged on both PRs (docs/rules only — no code changes expected).

## Non-goals (unchanged)

- Not touching `docs/analysis/`, `docs/session_logs/`, `docs/grant-architecture.html`.
- Not modifying memory files.
- Not deleting any archived plan.
- Not changing hooks.
- Not touching the active WP gap-removal plan (handled by its own PR).

## Resolved open questions

- **A7 conditional delete gate** — overridden by pick 5: split instead of delete. Resolved.
- **B1 enforcement strictness** — soft (paragraph in dev-workflow.md). Resolved.
- **`retro` scope gate** — dropped (not in picks). Can revisit if retro output becomes noisy.

## Remaining open question

- **Part B0's `plan-hygiene.md` gate** — proposed `paths: ["docs/**/*.md", "**/CLAUDE.md", "docs/plans/**/*.md"]`. Too narrow? Broaden to include `.claude/rules/**/*.md`? Flag for user check at PR 2 time.

---

## Corrections (2026-04-20, Part B execution)

1. **B1+B2 routing** — plan text said "Add paragraph to `dev-workflow.md`". Executed in `plan-hygiene.md` instead, because B0/B7 extracted all plan/doc-hygiene content out of `dev-workflow.md` into `plan-hygiene.md` for path-gating reasons. Adding B1/B2 to `dev-workflow.md` (`**/*.py` gate) would re-load plan rules on every Python edit — the opposite of B0's goal. Code-reviewer confirmed this routing is better than the plan.
2. **`plan-hygiene.md` gate** (remaining open question) — resolved to `paths: ["docs/**/*.md", "**/CLAUDE.md", "docs/plans/**/*.md", ".claude/rules/**/*.md"]`. B7's own table listed the fourth path to cover rule-authoring overlap; used that form.
3. **Font-color rules relocation** — B7 removed Rule-3 font-color rules from `dev-workflow.md`; B4 added them to `openpyxl.md` §Core (`### Styling`). Intermediate absence spanned one commit (30d08b3..f0d66ae); acceptable since `.claude/rules/` path-gating means no consumer was loading the rules during that window that needed them.
4. **`dev-workflow.md` line count** — shipped at 31 total lines (17 rule lines). Plan acceptance criterion said ≤18 lines. Interpreted strictly on total-line count, this is unmet; interpreted on rule lines, it's met. Compressing further requires merging 10 distinct Build+Test rules into paragraphs, which hurts skimmability. Accepted as-is; noted as open design question for future audit.
