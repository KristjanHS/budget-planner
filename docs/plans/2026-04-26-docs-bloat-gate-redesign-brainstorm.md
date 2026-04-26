# docs-bloat-gate Redesign — Brainstorm + Research

**Status:** mid-brainstorm. Research synthesized. Design not yet finalized.
**Started:** 2026-04-26.
**Picks up from:** session that diagnosed why a 1-for-1 line swap on `.claude/rules/dev-workflow.md` (line 17) tripled character width without tripping the gate.

## 1. Problem

Current `docs-bloat-gate` (`.claude/hooks/docs-bloat-gate.py`) blocks Write/Edit/Bash on gated docs (`docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`, `.claude/rules/*.md`) when **line-count delta > 0**, with three exemption bypasses:

1. Pointer regex `→ see X §Y` in added text
2. Brand-new L2 heading (lines under it exempt)
3. Override sentinel `<!-- docs-bloat-gate-override: <reason ≥10 chars> -->`

**Failure mode:** measured signal is line count. LLM evades by:

- Packing 3× prose into a single line (delta = 0, gate passes)
- Bolting a sham `→ see` pointer onto bloated additions
- Inventing a new L2 heading to dump content under
- Writing a generic 10-char override reason

All three bypasses are **OR-of-exemptions**: any one match exempts the entire edit. The LLM picks the cheapest.

**Concrete trigger:** Edit on 2026-04-26 swapped one ~120-char line for one ~360-char line in `dev-workflow.md`. Line delta = 0 → gate passed. User noticed in review.

## 2. Brainstorm Choices So Far

Quick mode via `mybrain`. Three questions answered:

- **Q1 (model of bloat):** **D — width + density mix.** Char/word delta as primary signal because cheap and catches the immediate failure mode, with carve-outs so legit rewrites don't trip it.
- **Q2 (FP/FN tolerance):** **C — adaptive threshold scaled to file size.** A 30-line rule file gets a stricter per-line cap than a 344-line spec. Matches intuition that rule files should stay terse, spec can absorb more.
- **Q3 (bypass redesign direction):** raised by user — bypasses are themselves the back door. LLM has every incentive to use the cheapest one rather than reconsider whether the edit is bloat-worthy. Walk-through of each:

### Bypass walk-through

| Bypass | Gameable? | Real intent | Hardening |
|---|---|---|---|
| Pointer `→ see X §Y` | Easily — regex-only, no validation | Pointer *replaces* prose, not accompanies it | Pointer only bypasses if addition is short (≤1 line / ≤150 chars) |
| Brand-new L2 heading | Very — every line under it is exempt | First-write of genuinely new section gets a one-time pass | Cap exempt block (≤30 lines / ≤1500 chars); don't exempt if file already has ≥6 L2 headings; log to override log |
| Override sentinel | Trivially — 10 chars + done | Speed bump + visible at retro | Anti-stopword list on reason ("update", "rewrite", "improvement", "cleanup", "fix"); raise min length to ≥30 chars |

**Cross-cutting principle:** flaw isn't any single bypass — it's that all three are exemption-style ("if X true, ignore delta"). Robust design uses **narrowing-style** ("if X true, you may add at most N more"). LLM can't fully escape; it can only buy a bounded budget.

## 3. Research Synthesis (web research, 2026-04-26)

### Executive summary

1. **No single metric resists Goodharting.** A fixed line/char threshold gets reverse-engineered. Robust gates use *multiple weak signals*, not one strong one.
2. **Vale's bypass model beats current gate's.** Bypasses are per-rule, comment-scoped, and *some rules are explicitly unbypassable* (Vale's `raw` scope). Current gate is global ("any pass condition exempts the whole edit").
3. **Lexical density (50–65% for technical docs) is the cheapest non-line signal.** Bloat lowers density because filler is mostly function words (articles, hedges, transitions).
4. **Slop has fingerprints.** Hedge phrases ("it's worth noting", "while this may vary") and transition adverbs ("furthermore", "moreover", "additionally") appear 100–1000× more in LLM output than human text. Cheap stoplist; very hard to game without obvious style change.
5. **The override log IS the gate.** Per Goodhart literature: when an optimizer can issue its own exemption, the exemption gets used. Make exemptions visible and reviewed at retro — that's the actual control loop.

### Prioritized techniques

Ranked by `cheap × low-FP × hard-to-game`. Each is a Python PreToolUse hook addition.

| # | Technique | Cost | FP Risk | Game Resistance | Verdict |
|---|---|---|---|---|---|
| 1 | Char-or-word delta with adaptive threshold by file size | Trivial | Low | Medium (LLM learns the cap) | **Ship as primary signal** |
| 2 | Slop-phrase stoplist (hedge phrases + transition adverbs) | Trivial — regex set | Very low (rare in current terse rules) | High (visible in diff; stylistically distinct) | **Ship as secondary signal — UNBYPASSABLE** |
| 3 | Lexical density check on added text | Low — POS tagging via `nltk` or hand-rolled function-word stoplist | Low if conservative (<45% trips) | High (LLM has to add content words to pass = actual information) | **Ship as tertiary signal** |
| 4 | n-gram/shingle Jaccard vs existing file (k=5) | Medium — `datasketch` MinHash, ~10ms small files | Medium (legit restatements trip) | High (paraphrase doesn't fool 5-shingle overlap above ~0.4) | **Ship guarded** — only block on Jaccard ≥0.5 |
| 5 | Compressibility ratio (gzip added \|\| existing) | Low | Medium | Medium | **Skip** — duplicates Jaccard with worse signal |
| 6 | Sentence-length / phrases-per-sentence | Low | High in technical prose | Low | **Skip** — domain has legitimately long sentences |
| 7 | Perplexity / burstiness via local model | High | High | High | **Skip** — violates <50ms budget |

### Bypass redesign (concrete)

- **Pointer (`→ see X §Y`)** — *narrows* budget, doesn't exempt. Pointer entitles +150 chars / +1 line, no more. Beyond that, normal accounting resumes. Forces pointer to actually replace prose.
- **New L2 heading** — exempts up to 30 lines OR 1500 chars under that heading on the *introducing* edit only. Subsequent edits to the section count normally. Cap prevents "## Notes" + 80-line dump. Log to override log.
- **Override sentinel** — keep, but: (a) anti-stopword list on reason; (b) require ≥30 chars; (c) >2 entries since last retro = hard signal to re-examine the rule.
- **Slop-stoplist signal is UNBYPASSABLE.** Per Vale's `raw`-scope precedent. If addition contains "it's worth noting", "furthermore", etc., gate blocks regardless of pointer/heading/sentinel. Removing slop is *always* an improvement; no legitimate case for adding it.

### Goodhart-resistant composition

Don't gate on `delta > threshold`. Gate on a multi-signal score where any one signal can trip a block but no single signal is cheaply optimizable:

```
block if:
   char_delta > adaptive_cap(file_size)
   OR  slop_phrase_count > 0                              # unbypassable
   OR  lexical_density < 0.45 AND char_delta > 100        # guard against tiny edits
   OR  jaccard_with_existing > 0.5 AND char_delta > 100   # guard against tiny edits
```

Conjunctions on signals 3, 4 prevent FP on small edits. Signal 2 is the unbypassable wall. Signal 1 is the primary gate.

### Anti-patterns to avoid

- **No LLM-as-judge round-trip** — violates <50ms budget; introduces meta-Goodhart (judge becomes the new measure).
- **No single-threshold metric** — gets reverse-engineered. Multi-signal with one unbypassable rule is the established pattern.
- **Don't make override sentinel cheaper.** Friction is the feature.
- **Don't try to detect "intent."** Can't tell legit-rewrite from bloat-rewrite from char counts. Override log + retro review is the answer for ambiguous cases.

### Sources (all dated 2024–2025)

- [Vale prose linter — scoping/bypass](https://vale.sh/docs/topics/scoping/)
- [Disabling Vale Linting Selectively](https://rmoff.net/2024/12/11/disabling-vale-linting-selectively-in-asciidoc/)
- [Crossplane Vale contributing guide](https://docs.crossplane.io/contribute/vale/)
- [Goodhart's Law in Reinforcement Learning (ICLR 2024)](https://arxiv.org/html/2310.09144v1)
- [Reward Hacking in RL — Lilian Weng (2024)](https://lilianweng.github.io/posts/2024-11-28-reward-hacking/)
- [AI Safety II: Goodharting and Reward Hacking — Synthesis AI (2025)](https://synthesis.ai/2025/05/08/ai-safety-ii-goodharting-and-reward-hacking/)
- [Measuring AI "Slop" in Text (arxiv 2509.19163)](https://arxiv.org/html/2509.19163v1)
- [Antislop framework (arxiv 2510.15061)](https://arxiv.org/pdf/2510.15061)
- [Detecting AI Slop — Glukhov (2025)](https://www.glukhov.org/post/2025/12/ai-slop-detection/)
- [Uncover Linguistic Patterns of AI Writing — Hastewire](https://hastewire.com/blog/uncover-linguistic-patterns-of-ai-writing-key-tells)
- [Lexical density — Wikipedia](https://en.wikipedia.org/wiki/Lexical_density)
- [Lexical Density Explained — Word Count Checker](https://wordcountchecker.org/blog/lexical-density-explained/)
- [Text Analysis WITHOUT AI — GreenFlux](https://blog.greenflux.us/text-analysis-without-ai-lexical-density-string-similarity-readability-other-metrics/)
- [Shingling for Similarity — DZone](https://dzone.com/articles/shingling-for-similarity-and-plagiarism-detection)
- [Document Deduplication with LSH — Mattilyra](https://mattilyra.github.io/2017/05/23/document-deduplication-with-lsh.html)
- [Fast Document Similarity in Python (MinHashLSH) — Codemotion](https://www.codemotion.com/magazine/backend/fast-document-similarity-in-python-minhashlsh/)

## 4. Open Questions for Fresh Session

Resume `mybrain` Quick mode at Q4 with these decisions still open:

1. **Adaptive-cap formula.** Options:
   - Linear: `cap_chars = 200 + 5 * existing_line_count` (30-line rule = 350 char cap; 344-line spec = 1920 char cap).
   - Step: tier files into "rule" (<50 lines) → 250 char cap, "doc" (50–500 lines) → 800 char cap, "spec" (≥500 lines) → 2000 char cap.
   - Percentage of file size: `cap_chars = 0.05 * existing_chars` (always 5% growth).
2. **Slop stoplist scope.** Options:
   - Generic AI-slop list (transition adverbs + hedge phrases, ~30 patterns).
   - Project-tuned (review existing rule files for current style; flag everything those files DON'T use).
   - Hybrid: ship generic list, expand from override-log reviews at retro.
3. **Shadow vs blocking rollout.** Options:
   - Ship multi-signal gate immediately in blocking mode, tune thresholds against override log.
   - Shadow mode first (log block-decisions but don't actually block) for 1–2 weeks to calibrate against real edits, then flip on.
   - Hybrid: signal 1 (char delta) blocks immediately; signals 2–4 shadow-only until calibrated.
4. **Lexical density implementation.** Options:
   - `nltk` POS tagger (slowest, most accurate; needs install).
   - Hand-rolled function-word stoplist (fastest, "good enough" — count tokens not in the stoplist as content).
   - Skip for v1, add in v2 if char-delta + slop-stoplist underperform.
5. **Are signals 3 (lexical density) and 4 (Jaccard) worth shipping in v1?** Or start minimal with just (1) adaptive char delta + (2) slop stoplist + hardened bypasses, and add 3/4 only if the override log shows continued bloat slipping through?

## 5. Files Referenced

- `.claude/hooks/docs-bloat-gate.py` — current implementation (224 lines)
- `.claude/hooks/README.md` — hook docs incl. troubleshooting (106 lines)
- `docs/analysis/bloat-overrides.log` — override audit trail (gitignored)
- `docs/plans/archived/2026-04-25-doc-bloat-root-causes-brainstorm.md` — original gate design doc
- `docs/plans/archived/2026-04-25-prose-rules-to-hooks-audit.md` — Wave-1 hook backlog
