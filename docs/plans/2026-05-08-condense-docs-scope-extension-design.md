# condense skill: scope extension to `docs/*.md`

**Status:** design
**Date:** 2026-05-08
**Originating note:** `~/.claude/projects/-home-<you>-projects-budget-planner/memory/project_state.md` (line 107) — recorded as deferred follow-up during the 2026-04-26 docs-bloat-gate-redesign session.

## Background

Two pieces of existing infrastructure shape this work:

1. **`condense` skill** (`~/.claude/skills/condense/SKILL.md`). Periodic-dedup arm of the delta-update discipline. Today it scopes to `**/CLAUDE.md`, `**/CLAUDE.local.md`, `.claude/rules/*.md`, and `~/.claude/CLAUDE.md`. Phase 4 produces per-item Edits with user approval — never a "consolidated v2" Write. Run trigger: `/retro` or judgement-based ("section accumulates ≥3 overlapping bullets").

2. **`docs-bloat-gate`** (`.claude/hooks/docs-bloat-gate.py`, shipped 2026-04-25, redesigned 2026-04-27). Write-time PreToolUse hook on `docs/spec.md` and `.claude/rules/*.md`. Three signals: S1 char-delta cap (rule=150, doc=800, spec=2000) — bypassable; S2 slop stoplist — unbypassable; S3 lexical density <0.45 — unbypassable. S1 bypasses: brand-new L2 heading exempt, plus override sentinel (1/session, ≥30-char reason).

The deferred follow-up: extend `condense`'s globs to also cover `<project>/docs/*.md`. The "lower urgency post-hook" annotation reflects that the bloat-gate now stops *new* bloat at write-time — `condense` would address *accumulated* dup/drift, complementary to the gate.

## Goals

- Apply `condense`'s analyses to the docs/ tree (primarily `docs/spec.md`, plus active plans and human-facing runbooks).
- Add docs-specific analyses: trigger-table drift, plan supersession, section-size flagging.
- Eliminate rule-drift between the bloat-gate's slop/density rules and condense's pre-check by sharing one source.
- Keep the existing posture: per-item Edits, user-approved, no whole-file rewrites.

## Non-goals

- Editing `docs/plans/archived/**` — frozen post-mortems, never modified.
- Touching `docs/analysis/**` — `.claudeignore`'d one-offs.
- Auto-firing on hook events. New triggers surface as `/retro` reminders, never invoke condense unprompted.
- Changes to the bloat-gate's enforcement logic (sentinel cap, override log, tier caps). Only rule *definitions* move.
- A `~/.claude/lib/` global library. Deferred until a second project adopts the bloat-gate.

## Chosen approach: extend in place + factor out shared library

Considered three approaches:

| Approach | Description |
|---|---|
| A — Extend in place | Edit `condense/SKILL.md` directly; copy slop+density definitions into the skill |
| B — New `condense-docs` skill | Fork: `condense` keeps CLAUDE.md/rules; `condense-docs` handles docs/ |
| **C — Extend + shared library** | A's skill changes, plus extract bloat-gate's S2/S3 rule definitions to `.claude/lib/docs_quality.py`. Hook + skill import from one source. |

**C wins** because the user-chosen "pre-check before Edit" gate-interaction posture implicitly requires the skill to know the slop+density rules. A satisfies this by *copying* (rule-drift risk); C satisfies it by *sharing* (rule-drift impossible). Same effort modulo a one-time hook refactor.

Evaluation matrix (1-5):

| Criterion | A | B | C |
|---|---|---|---|
| Complexity | 2 | 3 | 3 |
| Time | 1 | 2 | 2 |
| Risk (rule drift / divergence) | 2 | 3 | 2 |
| Extensibility | 3 | 3 | 4 |
| Alignment | 4 | 2 | 4 |

## Design

### Library: `.claude/lib/docs_quality.py`

Public surface:

```python
def find_slop(text: str) -> list[str]:
    """Return slop terms present in `text`. Empty list = clean."""

def slop_hit_counts(text: str) -> Counter[str]:
    """Counter of slop occurrences. Hook needs counts (not just presence)
    for net-added detection; find_slop is a thin wrapper over this."""

def lexical_density(text: str) -> float:
    """Content-word ratio. Returns 0.0–1.0."""

def density_threshold() -> float:
    """The shared S3 cutoff (0.45 today)."""
```

Constants `SLOP_WORDS`, `SLOP_PHRASES`, `FUNCTION_WORDS`, `DENSITY_THRESHOLD` move from the hook into this module. Hook concerns that *do not* move: tier caps, sentinel parsing, override log writing, Bash write-pattern regex, L2-heading exemption, payload handling.

### Hook refactor (`.claude/hooks/docs-bloat-gate.py`)

Add a `sys.path` insert for `.claude/lib/` (the hook is invoked as `python3 .claude/hooks/docs-bloat-gate.py`; sibling import via `sys.path.insert(0, str(Path(__file__).parent.parent / "lib"))` is the pragmatic shape). Replace inlined S2/S3 logic with calls to `find_slop` / `lexical_density`. Net delta: hook shrinks; library is new (~80 lines).

Existing 75 hook tests in `tests/hooks/test_docs_bloat_gate_v2.py` cover S2/S3 behavior end-to-end and need no migration. Optional: thin `tests/lib/test_docs_quality.py` (~5 tests) pinning slop list contents and density edge cases — recommended for rule-change visibility, not strictly required since hook tests exercise the library indirectly.

### Skill changes (`~/.claude/skills/condense/SKILL.md`)

**Phase 1 — Discovery, expanded globs:**

```
1. **/CLAUDE.md, **/CLAUDE.local.md  (existing)
2. .claude/rules/*.md                 (existing)
3. ~/.claude/CLAUDE.md                (existing)
4. docs/spec.md                       (NEW)
5. docs/human/*.md                    (NEW)
6. docs/plans/*.md                    (NEW — top level only, no archived/)
```

Explicit non-targets: `docs/plans/archived/**`, `docs/analysis/**`.

**Phase 2 — Analysis, four new types layered on existing three:**

| Analysis | What it detects | How |
|---|---|---|
| Intra/cross-file dup | Same point in spec.md + CLAUDE.md, or spec.md + a rule | Existing semantic-similarity scan, wider corpus |
| Trigger-table drift | CLAUDE.md reference-index rows pointing to `docs/spec.md` § references that no longer exist or were renamed | Parse the trigger table's "Read" column; grep spec.md for matching `## ` headings |
| Plan supersession | Active plans where `SHIPPED` / `DONE` / `landed` / `archived to` markers sit next to verbose pre-ship phase prose | Find the markers in `docs/plans/*.md` (not archived/), flag the section above as collapse-candidate |
| Section-size flagging | spec.md L2 sections >1500 chars | Mirrors the bloat-gate's `NEW_SECTION_CHAR_EXEMPT` threshold |

Each finding is presented separately in Phase 3, user approves per-item.

**Phase 3 — Interaction:** unchanged. Findings render with file paths + character preview + proposed action. Trigger-table drift findings show old-row → new-target side-by-side. Plan supersession findings show "collapse N lines into 1 done-line" with the proposed replacement.

**Phase 4 — Implementation, gate-aware:**

For each approved Edit on `docs/spec.md` or any future gated path:

1. Compute the proposed `new_text`.
2. Call `find_slop(new_text)` — if non-empty, surface to user before applying ("This Edit would trip S2 with terms: [...]. Reword?").
3. If `len(new_text) > 100`, call `lexical_density(new_text)` — if `< density_threshold()`, surface to user.
4. Char-delta tier cap (S1) is **not** pre-checked. Condense Edits are usually subtractive; on rare additive cases the sentinel bypass is the right escape hatch.

Pre-check failures are *advisory*, not blocking. The user can override (reword or proceed). The hook stays the unbypassable line of defence.

### Trigger doctrine update (top of SKILL.md)

Existing line: *"Run on `/retro` or when a section accumulates ≥3 overlapping bullets."*

Adds: *"Also run when (a) `docs/analysis/bloat-overrides.log` accumulates ≥3 sentinel-bypass entries since last condense run, or (b) `docs/spec.md` crosses a 100-line growth boundary (600, 700, 800…)."*

These are reminders surfaced *to the user when they `/retro`* — the retro skill greps the override log and checks spec line count, then appends a "condense recommended because…" note. NOT auto-firing. Same posture as today's "≥3 overlapping bullets" rule.

### SKILL.md size

~85 lines today → ~135 lines after. No progressive-disclosure split needed.

## Decisions log

1. Library lives at `<project>/.claude/lib/docs_quality.py` (project-local). Global lib deferred per YAGNI.
2. Library exports rule definitions only (`find_slop`, `slop_hit_counts`, `lexical_density`, `density_threshold`); enforcement stays in the hook. `slop_hit_counts` is added because the hook needs Counter-valued net-added detection, not just boolean presence; `find_slop` wraps it for the skill's "what hit?" use case.
3. Glob extension covers `docs/spec.md`, `docs/human/*.md`, `docs/plans/*.md` (top level only).
4. Four new analysis types added: intra/cross-file dup (extended corpus), trigger-table drift, plan supersession, section-size flagging.
5. Phase 4 pre-check is advisory, not blocking. Hook is the unbypassable line.
6. New triggers (bypass-log threshold, line-count growth) surface as reminders during `/retro`, never auto-fire.
7. Land library extraction + hook refactor as one commit; SKILL.md changes as a separate commit.

## Re-work audit

Two-commit plan, each commit touches a disjoint surface:

- Commit 1: `.claude/lib/docs_quality.py` (new) + `.claude/hooks/docs-bloat-gate.py` (refactor) + optional `tests/lib/test_docs_quality.py` (new).
- Commit 2: `~/.claude/skills/condense/SKILL.md` (extend).

No file is written-then-rewritten across commits. No phase ordering re-work.

## Open questions

None remaining at design time. Implementation may surface:

- ~~Whether the `sys.path` import shape is robust across the hook's invocation contexts (CLI, regression-test harness, future automation).~~ **Resolved at implementation:** both invocation paths (settings.json `python3 "$CLAUDE_PROJECT_DIR"/.claude/hooks/docs-bloat-gate.py` and the test harness `subprocess.run([sys.executable, str(HOOK_PATH)])`) produce an absolute `__file__`, so `Path(__file__).parent.parent / "lib"` resolves to `<project>/.claude/lib` in both cases. 72/72 in-scope hook tests pass.
- Whether trigger-table drift should *propose* replacement rows (best-fit existing § match) or only *flag* drift. Default: flag-only; promote to propose-and-apply if first-pass usage shows the user always picks the obvious fix.

## Deferred follow-ups

- `tests/lib/test_docs_quality.py` (~5 tests pinning slop list contents + density edge cases). Plan rated it "recommended for rule-change visibility, not strictly required." Skipped at implementation since the 72-test hook suite exercises the library indirectly. Worth adding when the slop list next changes — a slop-list edit can pass hook tests while silently shifting which terms are checked, because no test pins the list directly.
