# `ISNUMBER(range)` has no array context in Excel — Actuals EUR fix

**Date:** 2026-07-27
**Status:** diagnosis confirmed against Excel-computed evidence; ready to implement
**Scope:** Actuals EUR path only. Actuals-Hours is a **named, deferred gap** (§Stage B) at the
user's request — hours are not being filled at this stage.
**Supersedes the root cause in:** `2026-07-27-actuals-posted-idx-cse-design.md` (docs/plans/),
whose "`COLUMN(range)` needs CSE" mechanism is **disproven** below. That doc's shipped work (the
literal midx row, the recalc-gate fix) stands; only its causal claim is wrong.

---

## Symptom (unchanged)

After `python -m budget_planner rebuild` + recalc in Excel, every column in the Actuals
`ACTUAL EUR -i` / `+i` blocks renders blank except `Budget -i` / `Budget +i`.

Commit 8ab774c was supposed to fix this and did not. The user rebuilt on 508b852 and the
columns are still empty.

## Root cause — it is `ISNUMBER(range)`, not `COLUMN(range)`

Inside a non-CSE formula, Excel does **not** give `ISNUMBER(range)` array context. It collapses
to a scalar, so `MAX(… * ISNUMBER(range))` returns `0`. LibreOffice array-evaluates it
correctly, which is why every LO-recalced test in the suite is green.

### Evidence — natural experiment inside the user's own rebuilt file

`tmp/budget_plan.xlsx` was last saved by **Excel** (`docProps/app.xml` →
`<Application>Microsoft Excel</Application>`, `lastModifiedBy: Kristjan Sillmann`,
`xl/calcChain.xml` present). Every cached `<v>` in it is Excel's own answer:

| Sheet!Cell | Inner test | Excel computed | Correct? |
|---|---|---|---|
| `Timelog Plan!G33` | `MAX((B20:AS20<>"")*B19:AS19)` | `'M24-Oct28'` | ✅ |
| `Actuals!BG7` (posted_idx) | `MAX((midx<=current_month)*midx*ISNUMBER(C7:AT7))` | `0` | ❌ (should be 9) |
| `Actuals!C1` (auto-derive) | `MAX($C$4:$AT$4*ISNUMBER($C$7:$AT$13))` | `0` | ❌ (should be 9) |
| `Actuals-Hours!AW5…AW31` (24 cells) | `… * ISNUMBER(…)` | all `0` | ❌ |

Same file, same engine, same `SUMPRODUCT(MAX(…))` wrapper, same literal-midx-row technique.
The **only** difference between the working row and the broken ones is `(range<>"")` vs
`ISNUMBER(range)`.

### `COLUMN(range)` is exonerated

From `tmp/budget_plan_JULY_ACTUALS_16.06.2026.xlsx` (the last workbook where these columns
were filled — also Excel-saved, also Excel-computed):

```
AX6 = SUMPRODUCT((COLUMN(C6:AT6)-COLUMN($C$1)+1<=current_month)*(C6:AT6))  →  4884  ✅
```

`COLUMN(range)` array-evaluated correctly inside `SUMPRODUCT` with no CSE. Meanwhile that same
old file's `C1` — `SUMPRODUCT(MAX((COLUMN(...)+1)*ISNUMBER($C$6:$AT$12)))` — computed `0`. The
bug was **already present in the old workbook**; nothing depended on `C1` back then
(`MAX(1, eur_actuals_through_month)` degrades a 0 harmlessly), so it stayed invisible.

**Therefore:** 8ab774c changed a variable that was never the cause. The midx row it introduced
is harmless and stays — it is clearer than the `COLUMN()` arithmetic and is already tested — but
the prose asserting `COLUMN(range)` needs CSE must be corrected wherever it appears.

### Why the whole right-hand block goes blank

`posted_idx = 0` propagates: `Slope`, `Variance`, `Posted`, `% of Budget`, `Breach-In` — both
`-i` and `+i` — are each wrapped `IF(posted_idx=0,"",…)`. `Budget` alone has no guard and no
array expression, which is exactly why it alone survives.

### Why no test caught it, and why no currently-written test could

`scripts/recalc.py` drives **LibreOffice**. LO evaluates `ISNUMBER(range)` in array context
correctly. There is no Excel-engine evaluation anywhere in CI. A recalc-based test therefore
cannot see this defect class at all — `tests/test_actuals_posted.py::gap_recalc_path` seeds a
sparse posting pattern, asserts `posted_idx == 5`, is green, and the bug shipped anyway.

Verified directly (LO 7.3.7.2, via `scripts/recalc.py` on a scratch probe): the exact ported
formula returns `7` in LO where Excel returns `0`.

## The fix

Swap `ISNUMBER(rng)` → `(rng<>"")` inside array scans. This is the form
`timelog_plan.py:1888` already uses and that the user's Excel demonstrably evaluates correctly.

```
=SUMPRODUCT(MAX(($C$4:$AT$4<=current_month)*$C$4:$AT$4*(C7:AT7<>"")))
```

Verified in LO across the edge cases (scratch probe, all four correct):

| Case | idx | value |
|---|---|---|
| two postings at M3, M7 (`current_month`=9) | 7 | 200 |
| posting at M10, beyond `current_month`=9 | 3 | 100 (correctly capped) |
| empty row | 0 | `""` |
| **zero-valued** posting at M3 | 3 | 0 (not treated as blank) |

### Semantic delta, accepted

`(rng<>"")` counts a cell holding **text** as posted, where `ISNUMBER` would not. Inactive month
cells are written as `""` (`actuals.py:387`) → `"" <> ""` is FALSE ✅; active-unfilled cells are
truly blank → FALSE ✅. Only a user typing free text into a EUR grid cell diverges, and it
would then be counted as a posting. `timelog_plan.py` already lives with this. Accepted.

Note `=IF(ISNUMBER(B1),B1,C1)` (`D1` on both sheets) and the `ISNUMBER` uses in
conditional-formatting rules (`trends.py:387,395,403`, `capacity.py:632`,
`timelog_plan.py:1075`) are **scalar** contexts — not affected, do not touch.

---

## Stage A — Actuals EUR path (this session)

| # | Site | Change |
|---|---|---|
| A1 | `budget_planner/actuals.py:405` | per-WP `posted_idx`: `ISNUMBER({wp_range})` → `({wp_range}<>"")` |
| A2 | `budget_planner/actuals.py:790` | `Actuals!C1` auto-derive: same swap on `{wp_scan_range}` |
| A3 | `budget_planner/trends.py:538` | reconciliation scan over `SHEET_ACTUALS` range: same swap |
| A4 | `budget_planner/actuals.py:176, 392-399, 784-786, 864` | correct the comments that assert `COLUMN(range)` needs CSE — replace with the real mechanism. Keep the midx row. |
| A5 | Tests | flip the three assertions that pin the broken form (below) |
| A6 | Guard | new static test, §Guard |
| A7 | `docs/spec.md` | record the invariant (CLAUDE.md rule 6) |
| A8 | `2026-07-27-actuals-posted-idx-cse-design.md` | absorb still-live content (RULING/Shape, midx-row rationale) into this plan or `docs/spec.md`, then `git mv` to `docs/plans/archive/` and repoint refs (both rename halves in the pathspec). |

**Session decisions (qimpag round, 2026-07-27) — these supersede the section bodies below where they differ:**
- **Stage B stays deferred.** Hours sites remain in `KNOWN_CSE_GAPS`.
- **A6 guard detects by ARGUMENT, not wrapper** (see §Guard).
- **A4 scope extended** to `actuals.py:176` and `:864`, which carry the same disproven claim.
- **A8 added**: the superseded design doc is absorbed and archived, not banner-patched.

`trends.py:538` is in Stage A because its scan range is on the **Actuals** sheet
(`act_range_full = f"{SHEET_ACTUALS}!…"`) — it is part of the EUR path the user needs working,
not part of the hours path. Its `COLUMN()` arithmetic stays as-is (exonerated above); only the
`ISNUMBER` term changes.

### Tests to flip (A5)

| Test | Current assertion | Becomes |
|---|---|---|
| `tests/test_actuals_posted.py:68-76` | `assert "ISNUMBER(" in formula` | `assert '<>""' in formula` + `assert "ISNUMBER(" not in formula` |
| `tests/test_actuals_breach_todate.py:150-159` | `assert "MAX(" and "ISNUMBER(" in auto` (Actuals C1) | same swap. **Leave line 178/180 alone** — that is the Actuals-Hours case, still deferred |
| `tests/test_trends.py:193` | `assert "ISNUMBER" in formula` | `assert '<>""' in formula` |

Docstrings in all three carry the old rationale (`test_actuals_posted.py:9` says
"SUMPRODUCT-MAX-ISNUMBER scan") — update them in the same edit, they are the next reader's
first source.

### Guard (A6) — the thing that would actually have caught this

A static assertion over the built workbook: **no cell formula may pass a multi-cell RANGE to
`ISNUMBER(`.** Detect by argument shape — `ISNUMBER(` followed by a range reference
(`$?[A-Z]+$?\d+:$?[A-Z]+$?\d+`) — not by the `SUMPRODUCT(` wrapper. Keying on the wrapper
would let a future `MAX(…*ISNUMBER(range))` written without `SUMPRODUCT` sail through; the
defect is the range argument itself. Scalar uses (`ISNUMBER(B1)`, `IF(ISNUMBER(B1),B1,C1)`)
are permitted by construction. Walks `ws.iter_rows()` on the built (not recalced) workbook;
conditional-formatting rules are stored separately from cell values and are therefore skipped
for free.

It must carry an explicit allowlist for the deferred hours sites:

```python
KNOWN_CSE_GAPS = {  # Stage B — Actuals-Hours ISNUMBER scans, deliberately not fixed yet.
    (SHEET_ACTUALS_HOURS, ...),  # see 2026-07-27-isnumber-array-context-excel-fix.md §Stage B
}
```

so the gap is *named* rather than silently passing. Stage B is then literally "delete entries
from `KNOWN_CSE_GAPS` until it is empty."

**Demonstrate it can fail before trusting it** (a check incapable of returning non-zero is not
a check): run the guard against the pre-fix builder output — it must go red on the Actuals
`posted_idx` rows and `C1`. Record the count in the commit message.

### Verification

1. `make check` (recalc + pytest) — LO-side, proves no regression. **Does not prove the fix.**
2. Guard red on pre-fix build, green on post-fix build.
3. **User-owed, the only real proof:** rebuild, open in Excel, confirm the `ACTUAL EUR -i` block
   populates. Ranked checklist at handoff — top item is `Actuals!BG7` (hidden helper col,
   expect `9` not `0`); if that is 9 the whole block follows.

### Note on the user's current file

`Actuals!D1` in `tmp/budget_plan.xlsx` holds a **literal `9`**, not the builder's
`=IF(ISNUMBER(B1),B1,C1)`. It was hand-typed over — accidentally load-bearing, since the broken
`C1` scan returns 0 and `D1` is the target of `eur_actuals_through_month`. After the fix, `C1`
auto-derives 9 on its own and the rebuild restores `D1` to a formula. No action needed, but the
literal will disappear on rebuild and that is correct, not data loss.

---

## Stage B — Actuals-Hours (deferred, user's call)

Out of scope this session: the user is not filling hours data at this stage.

| Site | Cell |
|---|---|
| `budget_planner/actuals_hours.py:86-89` | per-person-row `INDEX`+scan → currently `0` in all 24 `AW` cells |
| `budget_planner/actuals_hours.py:455` | `Actuals-Hours!C1` auto-derive |
| `tests/test_actuals_hours.py:222-223,246` | assertions pinning the broken form |
| `tests/test_actuals_breach_todate.py:178,180` | hours-side C1 assertion |

These are broken **today** in Excel (all `AW` = 0 in the user's file) and stay broken until
Stage B. They are held in `KNOWN_CSE_GAPS` so no future session mistakes the green suite for
coverage. `hours_actuals_through_month` therefore reads 0 → Capacity "Actual PM" shows 0, which
is correct-by-accident while no hours are posted, and wrong the moment any are.

---

## Open, not chased

`2026-07-27-actuals-posted-idx-cse-design.md` §3 asks why LO no-ops on the user's Excel-saved
workbook. Still undiagnosed, still worth its own investigation, unaffected by this fix.

The precise Excel mechanism for `ISNUMBER(range)` collapsing (implicit intersection vs. scalar
coercion) is **not proven** — only the input/output behaviour and the discriminating variable
are. The fix removes the dependency rather than reasoning about it, so it is safe either way.
