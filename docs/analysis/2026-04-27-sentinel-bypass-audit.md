# Sentinel Bypass Audit — Brief for Deep-Research Handoff

**Date:** 2026-04-27
**Status:** Audit complete, fixes proposed but not implemented.
**Goal of this brief:** hand a fresh deep-research context everything it needs to (a) understand our current LLM-bloat gates, (b) see how they're cheatable, and (c) replace our hand-rolled denylists with SoTA detection of LLM-generated bloat / boilerplate / off-topic justification text.

The reader of this doc is expected to be a fresh agent with no prior context on this repo. Everything load-bearing is inlined; file paths are anchors for verification, not prerequisites.

---

## 1. Context

### 1.1 What this repo is

`budget-planner` — an openpyxl-based Excel workbook generator for EU project FTE budget planning. Single-developer + Claude collaboration. The codebase is small; the *prose surface* (specs, runbooks, rules, plans) is large and grows fast under LLM-assisted editing.

### 1.2 Why gates exist

The user observed Claude (Opus 4.x family) repeatedly:

- Inflating spec/runbook/rule files with hedges, transition adverbs, and AI tells (`furthermore`, `comprehensive`, `it's worth noting`, `delve into`, …).
- Adding 800+ char "improvements" to `.claude/rules/*.md` files where the original rule was 40 lines.
- Producing audit/analysis reports unprompted at `docs/<name>.md` root.
- Archiving plans (`git mv … archived/`) without first re-running pytest, even though one test parses the named-range table out of `docs/spec.md` and a docs-only restructure can break it.

Two PreToolUse hooks were built to push back:

- `.claude/hooks/docs-bloat-gate.py` — three-signal block on Write/Edit/Bash to `docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`, `.claude/rules/*.md`.
- `.claude/hooks/docs-only-plan-gate.py` — block `git mv …archived/` unless pytest just passed; PostToolUse stamps on green pytest.

Both expose a markdown-comment **override sentinel** the model can include to bypass.

### 1.3 Hooks fail open

`try: main(); except Exception: pass; sys.exit(0)` — a hook bug never blocks Claude.

---

## 2. Current Implementation

### 2.1 `docs-bloat-gate.py` (≈540 lines)

**Gated paths:** `docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`, anything under `.claude/rules/*.md`. Optional `CLAUDE.md` via env flag.

**Three signals (any blocks):**

| Signal | Check | Bypassable? |
|--------|-------|-------------|
| **S1** | `char_delta` (= `len(new) − len(old)`) exceeds tier cap. Tiers by current line count: `rule<50→150 chars`, `<500→800`, `≥500→2000`. | Yes — L2-heading carve-out OR override sentinel |
| **S2** | Net-added text (added counts minus removed counts, multiset diff) contains any phrase from a ~30-pattern stoplist. | **No** |
| **S3** | Lexical density (content-word ratio against ~150 English function words) `< 0.45` on additions `> 100 chars`. | **No** |

**S2 stoplist:** ~30 entries split between single words (`furthermore`, `moreover`, `additionally`, `consequently`, `subsequently`, `accordingly`, `nevertheless`, `notwithstanding`, `fundamentally`, `essentially`, `ultimately`, `comprehensive`, `robust`, `seamless`, `leverage`, `utilize`) and multi-word phrases (`it's worth noting`, `it should be noted`, `it is important to`, `that said`, `having said that`, `in essence`, `at its core`, `while this may vary`, `in the context of`, `delve into`, `navigate the`, `landscape of`, `realm of`, `tapestry of`).

**S3 lexical density:** `re.findall(r"\w+", text.lower())` → drop tokens in `FUNCTION_WORDS` (~150-word set: articles, pronouns, conjunctions, prepositions, auxiliaries, common modals, determiners, fillers, contraction stubs `s|t|re|ve|ll|d|m`) → ratio = content / total. Threshold 0.45, only fires above 100 chars.

**L2-heading carve-out:** A brand-new `## ` heading not present in `current` exempts the introducing edit up to 30 lines OR 1500 chars (whichever hits first), then S1 measures the remainder.

**Override sentinel (S1 only):**

- Pattern: `<!-- docs-bloat-gate-override: <reason> -->` — regex `<!--\s*docs-bloat-gate-override:\s*(.+?)\s*-->` (DOTALL).
- Validation:
  - `len(reason.strip()) >= 30`.
  - Reason lowercased, then `re.search(r"\b{sw}\b", lower)` against `{update, rewrite, improvement, cleanup, fix}` — any hit rejects.
- Hard cap: **1 sentinel per session.** Counter file at `~/.claude/state/docs-bloat-gate-sentinels-<session_id>.txt`, single integer.
- `session_id` resolves from `CLAUDE_SESSION_ID` → md5 of `CLAUDE_TRANSCRIPT_PATH` → md5 of `cwd|YYYYMMDDHH` (hourly fallback).
- Each accepted sentinel appends a JSONL entry to `docs/analysis/bloat-overrides.log` (gitignored): `ts, file, signal, bypass_type, reason, char_delta, session_id, tool`.

**Bash gate** (`>` / `>>` / `tee` / `sed -i` writing to a gated path): same regex, same 1/session cap.

**Audit-keyword block** (separate from the gate proper): a brand-new `Write` to `docs/<name>.md` (depth=1, not in `GATED_DOCS`) whose content matches `\b(audit|analysis|report|research)\b` is rejected unless the same sentinel is present.

### 2.2 `docs-only-plan-gate.py` (≈175 lines)

**Trigger:** PreToolUse Bash matching `\bgit\s+mv\b[^|;&]*\bdocs/plans/[^|;&]*\barchived/[^|;&]*`.

**Pass conditions (any one):**

- Stamp file `~/.claude/sessions/<md5(cwd)[:8]>/pytest-pass-stamp` exists with `mtime` age `≤ 1800s`.
- Command contains `<!-- docs-only-plan-gate-override: <reason ≥ 10 chars> -->`.

**No anti-stopword check. No session cap. No reason content checks beyond min length.**

**PostToolUse:** if Bash command matched `\bpytest\b[^|;&]*\btests/` AND `tool_response.exit_code == 0` (or stdout fallback heuristic: `\d+ passed` present and no `\d+ failed|error`), write stamp file containing an ISO timestamp string.

**Override log:** plain-text appends to `docs/analysis/docs-only-plan-overrides.log` with `ts, summary, reason`.

### 2.3 Settings wiring

Both hooks registered as `PreToolUse` (and `PostToolUse` for the plan-gate) in `.claude/settings.json`. The user is informed of fired hooks via stderr; Claude sees the rejection text and (per design) is expected to revise.

---

## 3. Audit Results

Eleven cheats found; ranked by ease of exploitation. Each one is an actual sentence or sequence Claude could emit and pass the gate today.

### 3.1 CRITICAL — Anti-stopword regex is morphology-blind

`\bupdate\b|\brewrite\b|\bimprovement\b|\bcleanup\b|\bfix\b` does not match any inflected form. `\bfix\b` does not match `fixing` because `\b` after `x` fails when `i` (a `\w`) follows. Verified: `re.search(r"\bfix\b", "fixing")` → `None`.

| Lemma | Bypassed by |
|-------|-------------|
| `update` | `updates`, `updating`, `updated` |
| `fix` | `fixes`, `fixing`, `fixed` |
| `rewrite` | `rewriting`, `rewrote`, `re-write` |
| `improvement` | `improve`, `improving`, `improved`, `improvements` |
| `cleanup` | `clean up`, `cleaning`, `clean-up` |

Passing reasons today:
- `"fixing the bloated section per spec"` (35 chars)
- `"updating the cross-reference table"` (35 chars)
- `"cleaning up cross-section drift now"` (35 chars)

### 3.2 CRITICAL — Sentinel content acts as a density-pad for S3

`gate_text_edit` runs S2 → S3 → S1 in order, and `lexical_density(added)` runs over the **entire added string including the sentinel HTML comment**. The sentinel's natural tokens (`docs`, `bloat`, `gate`, `override`, `required`, `carve`, `subsection`, …) are ~80% content words.

A 200-char addition at 0.42 density (S3 fails). Prepend a sentinel of ~80 chars carrying ~10 content words. Combined ratio rises above 0.45 and S3 passes. **The check labelled UNBYPASSABLE is bypassable by the same string used to bypass S1.** S2 is set/multiset based and less affected, but a carefully phrased sentinel could still mask hits in adjacent prose under the multiset diff.

### 3.3 CRITICAL — `docs-only-plan-gate` has no stopword check and no session cap

Regex `(.{10,}?)`. `<!-- docs-only-plan-gate-override: fix it lol -->` (12 chars) passes. Repeatable indefinitely. The gate is functionally a stop sign with no enforcement once the model knows the syntax.

### 3.4 CRITICAL — pytest-pass stamp has no integrity binding

Stamp is a plain file with mtime. Two trivial bypasses:

- Direct: `touch ~/.claude/sessions/<hash>/pytest-pass-stamp && git mv …`. Stamp content (an ISO date string) is never read back.
- stdout-fallback: `PYTEST_RE = \bpytest\b[^|;&]*\btests/`. A Bash command like `echo 'pytest tests/foo' && echo "5 passed"` matches the regex and the success heuristic, stamping without any pytest execution.

Worse, **30-min staleness is time-bound, not state-bound.** Pass pytest at HEAD A → modify the spec table at HEAD B → archive a plan at HEAD B+5min. Gate passes; pytest now fails; user has no signal.

### 3.5 HIGH — Reason text is not required to relate to the edit

The validator checks length and stoplist. It does not check that any reason token appears in the diff or filename. Passing today:

- `"see plan doc for context xxxxxxx"` (32 chars)
- `"required for ongoing redesign work"` (35 chars)
- `"per design doc 2026-04-26 v2.0"` (30 chars)
- `"genuinely necessary edit per spec"` (33 chars)

The reason describes intent in the abstract; the diff is whatever you wanted. **No reason↔diff correlation is enforced.**

### 3.6 HIGH — Reason history not consulted across sessions

Logs are write-only. A reason that worked once works forever — copy-paste across sessions is undetected. The session counter resets per session by design; without cross-session memory, the cap effectively becomes "1 per session × infinite sessions = infinite."

The README mentions `≥3 entries since last /retro is a signal to revisit calibration` — that is a **human ritual**, not an automated check.

### 3.7 HIGH — Sentinel only needs to be in `added`, not persistent in the file

`OVERRIDE_RE.search(added)` looks at `new_string` for Edit (or `content` for Write). Nothing requires the sentinel to remain in the file afterward. A subsequent Edit can quietly remove the sentinel within tier cap. The override log captures the bypass, but the file itself looks clean — `git blame` and human review on the file are blind to the bypass.

### 3.8 HIGH — Multi-edit splitting under L2 carve-out + sentinel

Not strictly a sentinel cheat, but it dilutes the cap's purpose:

- Edit 1: new `## ` heading + 1500 chars (carved exempt).
- Edit 2: 800 more chars in same section (under doc-tier cap).
- Edit 3: 800 more chars (still under cap).

= 3100 chars added without burning a sentinel. The "1/session" implies a hard ceiling; in practice it's a fallback for the rare edit that can't be split.

### 3.9 MEDIUM — 30-char minimum is character-count, not semantic content

`"xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"` (30 x's) passes. So does `"a a a a a a a a a a a a a a a"`. The validator counts bytes, not distinct content tokens. Boilerplate reasons sit comfortably under 60 chars: most of them are five English words with no information density.

### 3.10 MEDIUM — Gate exposes the rejection rule to the agent in stderr

On rejection, stderr emits the entire stopword list. The model retries with listed words avoided. The README also lists them publicly, so this ship has partly sailed for adversarial cases — but it makes the model's adaptation cheaper and faster.

### 3.11 LOW — Sentinel inside a code fence still passes

The HTML-comment regex doesn't care about markdown context. A sentinel embedded in a fenced code block (e.g. an example of how *not* to bypass) is still accepted. Edge case but a clean reviewer-confusion vector.

---

## 4. Underlying Design Flaw

The gate conflates two checks that need different mechanisms:

1. **"Is this edit large?"** — measurable from the diff. S1/S2/S3 do this OK modulo §3.2.
2. **"Is the bypass justified?"** — *cannot* be measured from a 30-char string and a 5-word denylist. The gate pretends it can.

Length + stoplist is a **noise filter**, not a justification check. A noise filter can reject `"fix"` but cannot reject the intent behind `"adding subsection per design doc 2026-04-26"` covering a 4000-char dump unrelated to the cited subsection.

The sentinel is **single-use-by-counter, not single-use-by-cost**. Once the model knows the cap is 1, it hoards or front-loads. Each successive edit in a session is *not* harder to justify; the gate just slams the door after one.

---

## 5. Recommended Fixes (Quick Wins)

These are the engineering patches we already see — they close the bleeding holes but are deliberately *minimum-leverage* defences. The deep-research goal is to replace the underlying detection layer (§6).

| ID | Fix | Closes | Cost |
|----|-----|--------|------|
| F1 | Strip sentinel comment from `added`/`removed` before S2/S3 | §3.2 | 3 lines |
| F2 | Stem-trim stopword check + expanded list (`improve`, `clean`, `tidy`, `polish`, `tweak`, `minor`, `small`, `quick`, `simple`) | §3.1 | ~15 lines |
| F3 | Reason must mention an identifier (heading/path token/code-fenced symbol) present in the diff | §3.5, partial §3.6 | ~30 lines |
| F4 | Cross-session reason-hash dedup (`~/.claude/state/bloat-override-history.txt`, last 30 / 14d) | §3.6 | ~20 lines |
| F5 | Bind pytest stamp to `git rev-parse HEAD`; verify on use | §3.4 | ~15 lines |
| F6 | Mirror docs-bloat-gate's reason validator into docs-only-plan-gate; add 3/session cap | §3.3 | ~30 lines |
| F7 | Strip stopword list from rejection stderr | §3.10 | trivial |
| F8 | `/retro` skill body parses both override logs and surfaces drift metrics | §3.6 supportive | medium |
| F9 | Require sentinel to persist in committed file | §3.7 | medium, low ROI |

**Suggested ship order:** F1 → F2 → F3 → F5 + F6 → F4 + F8.

These fixes raise the cheating cost from "type 30 chars" to "compose a sentence that names this edit's actual identifiers." That is the threshold where the gate begins to function. **It is still a denylist + heuristic stack.**

---

## 6. Deep-Research Goal

We want SoTA replacements for the heuristic detection layer. Three concrete sub-questions, each with a ready evaluation harness in this repo.

### 6.1 LLM-slop detection beyond denylists

Our S2 stoplist is ~30 phrases. Real LLM bloat in 2025–2026 routinely avoids those tokens; modern models write polished filler that passes any literal denylist.

**Find:**

- Published detectors for LLM-generated text bloat / verbosity / hedging that operate on **short text spans (50–4000 chars)** without round-trip API calls. Examples sought: GLTR-style perplexity profiling, n-gram repetition signatures, syntactic-template fingerprints, embedding-space distance to a corpus of known-good human prose, fine-tuned small classifiers (≤ 200MB, CPU-runnable).
- Papers/code on **"AI tell" detection** specifically for technical prose (specs, docs, code-adjacent markdown) — most public detectors target essays/news.
- Empirical studies of which signals survive paraphrase attacks where a strong model is told to "remove obvious AI tells."

**Evaluation corpus we can produce:**

- ~6 months of git history on `docs/spec.md`, `docs/runbook.md`, `.claude/rules/*.md`. Diffs marked "human-only" (pre-LLM commits) vs "LLM-assisted" (commits where Claude was the author per `Co-Authored-By` trailer).
- Pre-categorised slop incidents in `docs/refactor-incidents.md`.

### 6.2 Justification verification (reason ↔ diff alignment)

The override sentinel is meant to be a forcing function for thought. Today it's free-text. We want to detect **whether a reason actually describes the diff it accompanies**.

**Find:**

- Cheap NLI / textual-entailment models that can score `(reason_text, diff_summary) → entailed | unrelated` on CPU in <100 ms per call (PreToolUse hook latency target is <50 ms but we can budget more for this specific check).
- Methods used by code-review tools to detect "PR description doesn't match diff" — anything from `commit-mind`, `pr-agent`, `reviewdog`, internal Anthropic/OpenAI/Google practices that's been written up.
- Token-overlap baselines (Jaccard / TF-IDF cosine / SBERT cosine) — what threshold reliably separates "boilerplate" from "specific" reasons in a held-out set?

**Constraints:**

- No network calls from a PreToolUse hook (latency + privacy).
- Local model OK if it's <500MB and CPU inference is <300ms.
- False-positive cost is high — a legitimate edit blocked by a flaky NLI model erodes trust in the gate fast.

### 6.3 Bypass-token economics (single-use vs. token-bucket vs. cost-curves)

Our current cap is `1 sentinel/session`. This is a binary cliff. We suspect (but have not measured) that a **graduated cost** — each successive bypass demands a *more specific* reason, or a *longer* reason, or *different* anchors than prior burns — better matches how humans actually justify exceptions.

**Find:**

- Literature on **"escalating friction"** in human-in-the-loop and LLM-in-the-loop guardrails. Stack Overflow's reputation thresholds, Wikipedia's edit-rate limits, and Linux kernel's `Signed-off-by` chain are pre-LLM analogues; we want LLM-era equivalents.
- Empirical studies on **rate-limit token-bucket parameters** that minimise both legitimate-edit blocks and bloat throughput. Specifically: capacity, regen rate, and per-burn cost-curve shapes.
- Reports of **gaming patterns** under different bypass economies — when does the model learn to hoard, front-load, split, or batch around the limit?

### 6.4 Tools and references the researcher should sweep

- `.claude/hooks/docs-bloat-gate.py` (full source ≈540 lines).
- `.claude/hooks/docs-only-plan-gate.py` (full source ≈175 lines).
- `.claude/hooks/README.md` (user-facing troubleshooting).
- `docs/plans/archived/2026-04-26-docs-bloat-gate-redesign.md` (design rationale for v2).
- `docs/plans/archived/2026-04-25-prose-rules-to-hooks-audit.md` (origin of plan-gate, W1.2).
- `docs/refactor-incidents.md` (incident corpus).
- `docs/analysis/bloat-overrides.log`, `docs/analysis/docs-only-plan-overrides.log` (live sentinel use).
- The `~/.claude/rules/` global rule corpus (LLM-generated, hand-curated — useful as a "known slop-prone" target set).

### 6.5 Deliverable format from the deep-research pass

A markdown brief saved to `docs/analysis/<date>-sentinel-research-findings.md` with:

1. **Best-fit detector** for §6.1 (slop) — pick one, justify, cite.
2. **Best-fit verifier** for §6.2 (justification) — same.
3. **Best-fit economy** for §6.3 (bypass cost-curve) — same.
4. **Integration sketch** — a 50–100 line Python skeleton wiring the chosen detector into `gate_text_edit` between S3 and S1 (i.e. as a fourth signal, replacing or augmenting the current S2 denylist).
5. **Eval plan** — how we measure (a) false-positive rate against the human-only diff corpus and (b) true-positive rate against marked LLM-slop incidents.
6. **Defer list** — what was considered and rejected, with one-line rationale per entry. Keeps future research from rediscovering the same dead ends.

---

## 7. Out of Scope

- Replacing the hook architecture itself (PreToolUse + fail-open). It works.
- Touching the L1 rule files (`CLAUDE.md`, `.claude/rules/*`). Those are policy, not detection.
- Network-dependent detection at hook time. Latency + privacy + reliability all veto it for the PreToolUse path.
- Replacing the override sentinel UX. The HTML-comment form is well-understood by Claude and humans; the bug is in *what we check*, not *how it's expressed*.

---

**End of brief.** Hand to deep-research with the prompt: "Read this brief end-to-end, then research §6.1, §6.2, §6.3 in parallel and produce §6.5."
