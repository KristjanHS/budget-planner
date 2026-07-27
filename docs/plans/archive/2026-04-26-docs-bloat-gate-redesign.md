# docs-bloat-gate v2 — Design

**Status:** approved 2026-04-26 via `mybrain` quick mode. Supersedes `2026-04-26-docs-bloat-gate-redesign-brainstorm.md`.
**Driven by:** 2026-04-26 incident — Edit on `.claude/rules/dev-workflow.md` swapped one ~120-char line for one ~360-char line. Line delta = 0 → v1 gate passed. User noticed in review.

## 1. Goals

1. Catch prose-densification bloat (same line count, more characters / lower information density).
2. Remove single-bypass-exempts-everything composition. Each unbypassable signal is a wall; bypassable signal narrows budget rather than zeroing it.
3. Make sentinel abuse impossible by design (hard cap), not just discouraged.
4. Stay <50ms per hook invocation; zero new pip deps.

## 2. Scope

**In:** redesign of `.claude/hooks/docs-bloat-gate.py` and its README entry. Update CLAUDE.md Critical Rule 7 if it cites signal list.
**Out:** Jaccard/MinHash signal (signal 4 in research §3); LLM-as-judge; perplexity; reason-cites-artifact regex; budget carryover; retro stamp file. All revisitable from override-log evidence.

## 3. Architecture

Single Python file (~350 lines), PreToolUse hook on Write/Edit/Bash for gated paths. Stateless except per-session sentinel counter at `~/.claude/state/docs-bloat-gate-sentinels-<session_id>.txt`.

Gated paths (unchanged): `docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`, `.claude/rules/*.md`.

## 4. Signal composition

```
block if any:
  S1: char_delta > step_cap(file_tier)               # bypassable (sentinel + L2 heading)
  S2: slop_phrase_count > 0                          # UNBYPASSABLE
  S3: lexical_density < 0.45 AND char_delta > 100    # UNBYPASSABLE
```

### S1 — adaptive char-delta cap

Step tier on existing line count:

| Tier | Lines | Cap (chars) |
|---|---|---|
| `rule` | <50 | **150** |
| `doc` | 50–500 | **800** |
| `spec` | ≥500 | **2000** |

Rule cap intentionally tight: a single sentence-length addition passes; a paragraph dump trips. Rationale: rule files are an explicitly-managed terse tier.

### S2 — slop stoplist (unbypassable)

Initial ~30 patterns, case-insensitive whole-word match:

- **Transition adverbs:** `furthermore`, `moreover`, `additionally`, `consequently`, `subsequently`, `accordingly`, `nevertheless`, `notwithstanding`
- **Hedges:** `it's worth noting`, `it should be noted`, `it is important to`, `that said`, `having said that`, `in essence`, `at its core`, `fundamentally`, `essentially`, `ultimately`, `while this may vary`, `in the context of`
- **AI tells:** `comprehensive`, `robust`, `seamless`, `leverage`, `utilize`, `delve into`, `navigate the`, `landscape of`, `realm of`, `tapestry of`

Expand from override-log review at retro (Q5 = hybrid). Rationale per research §3.4: these phrases appear 100–1000× more in LLM output than human text and are rare in current rule files (verifiable via grep).

### S3 — lexical density (unbypassable)

`density = (tokens not in function_word_stoplist) / total_tokens`

- Function-word stoplist: ~150 English function words (articles, prepositions, conjunctions, auxiliaries, pronouns), inlined as Python set.
- Tokenize with `re.findall(r"\w+", added_text.lower())`.
- Block if `density < 0.45 AND char_delta > 100` (conjunction guards small edits where one filler word swings the ratio).

Hand-rolled stoplist chosen over `nltk` POS tagging: <1ms vs ~5ms+model-load; tracks POS within ~5% on technical prose at this coarse threshold.

## 5. Bypass logic (S1 only)

S2 and S3 ignore all bypasses. Pointer bypass removed from v1 — even a narrowed pointer is a free 150-char add and the L2-heading + sentinel paths cover legitimate cases.

| Bypass | Effect | Logged? |
|---|---|---|
| New L2 heading | Exempts up to 30 lines OR 1500 chars under that heading on introducing edit only. Subsequent edits to the section count normally. | yes |
| Override sentinel `<!-- docs-bloat-gate-override: <reason> -->` | Bypasses S1 only. Reason ≥30 chars, no anti-stopwords (`update`, `rewrite`, `improvement`, `cleanup`, `fix`). **Hard cap N=1 per session.** | yes |

## 6. Sentinel cap (Q9)

- **N=1 per session.** First use passes (with log entry); any subsequent S1 block on a gated doc hard-blocks regardless of sentinel content.
- Counter file: `~/.claude/state/docs-bloat-gate-sentinels-<session_id>.txt` (single integer).
- Session-id source: `CLAUDE_SESSION_ID` env → hash of transcript path → mtime-based rolling token. Counter auto-resets on session-id mismatch.
- Every S1 block message carries the cap warning *before* any sentinel is burned (Q9 L2 — visible counter in error).

## 7. Error messages (LLM-facing)

S1 block:
```
docs-bloat-gate: char_delta=<N> exceeds cap=<cap> (<tier> tier, file=<path>).
Options:
  1. Revise edit to fit within budget
  2. Add L2 heading if introducing genuinely new section (≤30 lines / 1500 chars exempt)
  3. Override: <!-- docs-bloat-gate-override: <reason ≥30 chars, no stopwords> -->
     WARNING: 1 sentinel per session. After use, all further edits
     to gated docs hard-block until next session.
Sentinels used this session: <K>/1.
```

S2 block:
```
docs-bloat-gate: AI-slop phrases detected: <list>. UNBYPASSABLE.
Remove these phrases — they signal LLM-generated bloat regardless of edit size.
```

S3 block:
```
docs-bloat-gate: lexical density <D> below 0.45 on <N>-char addition.
Function-word ratio too high — prose is filler-dense. UNBYPASSABLE.
Tighten by removing hedges, articles, and transition phrases.
```

Sentinel-cap-hit block:
```
docs-bloat-gate: session sentinel cap exhausted (1/1 used).
All gated-doc edits hard-block until next session.
Previous override: <reason snippet> (file=<path>, time=<ts>).
```

## 8. Override log

`docs/analysis/bloat-overrides.log` (gitignored, format unchanged from v1). One JSONL line per bypass use: `{ts, file, signal, bypass_type, reason?, char_delta, session_id}`. Reviewed at retro to (a) expand slop stoplist from observed evasions, (b) calibrate caps if FP rate is high, (c) flag rule files that need restructuring if they consistently trip.

## 9. Rollout (Q6)

Blocking immediately, all signals. No shadow mode.

- LLM gets stderr feedback on every block; must revise content, add L2 heading, or use sentinel (subject to cap).
- Thresholds (S1 caps, S3 0.45) set conservatively from day 1 — better to under-block and tighten than over-block and erode trust.
- Tuning loop: override log → retro review.

## 10. Tests

`tests/hooks/test_docs_bloat_gate_v2.py`:

- **S1 caps:** each tier (rule/doc/spec); at-cap passes; over-cap blocks
- **S1 bypasses:** L2-heading + 25-line passes; L2-heading + 31-line blocks; sentinel passes once
- **S2 slop:** each stoplist phrase blocks; phrase removal passes; sentinel/L2 ignored
- **S3 density:** low-density 200-char addition blocks; low-density 50-char passes (delta guard); sentinel/L2 ignored
- **Sentinel cap:** 1st use passes with log entry; 2nd use hard-blocks; cross-session counter resets via session-id mismatch
- **Reason validation:** 29-char reason fails; "cleanup" stopword fails; valid reason passes
- **Existing v1 tests:** preserve regression coverage on the gated-paths list and exemption-log writes

## 11. Migration

1. Replace `.claude/hooks/docs-bloat-gate.py` in place (no data migration; override log format unchanged).
2. Update `.claude/hooks/README.md` table entry: signal list, sentinel cap, removed pointer bypass.
3. Grep `CLAUDE.md` for references to v1 bypass list (Critical Rule 7 cites the gate); update if signal/bypass names changed.
4. Add `tests/hooks/test_docs_bloat_gate_v2.py`; remove v1 test file once v2 lands.
5. `git rm` `docs/plans/2026-04-26-docs-bloat-gate-redesign-brainstorm.md` after this design lands (the brainstorm is mid-process scratch; this doc supersedes it).

## 12. Decisions log

| # | Decision | Rationale |
|---|---|---|
| Q4 | Step-tier char cap; rule=150, doc=800, spec=2000 | Predictable; explainable in error message; rule tier kept strict to catch dump-attacks where they're most likely |
| Q5 | Slop stoplist hybrid (generic + retro-expand), unbypassable | Generic list near-zero FP on current rule style; override log feeds expansion |
| Q6 | Blocking immediately, all signals | LLM gets stderr feedback and retries; conservative thresholds substitute for shadow mode |
| Q7 | Hand-rolled function-word stoplist for density | <1ms, zero deps; tracks POS within 5% at coarse threshold |
| Q8 | V1 ships S1+S2+S3; defer Jaccard | Avoids `datasketch` dep; lexical density covers most paraphrase-bloat |
| Q9 | Sentinel hard cap N=1 + visible counter; pointer bypass removed | Definitive prevention of serial abuse; pointer was a free add even when narrowed |

## 13. Files referenced

- `.claude/hooks/docs-bloat-gate.py` — current implementation (224 lines)
- `.claude/hooks/README.md` — hook docs
- `docs/analysis/bloat-overrides.log` — override audit trail (gitignored)
- `CLAUDE.md` Critical Rule 7 — gate list
- `docs/plans/archive/2026-04-25-doc-bloat-root-causes-brainstorm.md` — original gate design
- `docs/plans/archive/2026-04-25-prose-rules-to-hooks-audit.md` — Wave-1 hook backlog
