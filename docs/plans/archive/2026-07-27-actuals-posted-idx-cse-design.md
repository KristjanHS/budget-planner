# Actuals `+i` columns render blank — diagnosis + fix design

**Date:** 2026-07-27
**Status:** diagnosis confirmed; recalc-gate fix shipped; formula fix has an OPEN decision (needs Excel verification)
**Evidence file:** `tmp/budget_plan.xlsx` (user's Excel-saved workbook; gitignored)

---

## Symptom

After `python -m budget_planner rebuild`, every column in the Actuals `ACTUAL EUR +i` block
renders empty except `Budget +i`. Same for the collapsed `-i` block.

## Root cause

The monthly EUR data is **present and was not lost on rebuild**. `posted_idx` — the helper
every guarded column anchors on — evaluates to `0`.

Excel's own cached values in the user's file:

| Cell | Formula | Excel computed | Correct |
|---|---|---|---|
| `G13` | `SUM(G6:G12)` | 27 445 | ✅ |
| `AV6` | `Config!E14` | 18 000 | ✅ |
| `C1` (auto-scan) | `SUMPRODUCT(MAX((COLUMN($C$6:$AT$12)-COLUMN($C$1)+1)*ISNUMBER($C$6:$AT$12)))` | **0** | ❌ (should be 9) |
| `BG6` (`posted_idx`) | `SUMPRODUCT(MAX((COLUMN(C6:AT6)-COLUMN($C$1)+1<=current_month)*(COLUMN(C6:AT6)-COLUMN($C$1)+1)*ISNUMBER(C6:AT6)))` | **0** | ❌ (should be 9) |

Data present: WP1 `G6` = €4 884 (M5), `K6` = €7 500 (M9); WP2–WP4 likewise.
`current_month` (`Trends!B1`) = 9, so both postings are inside the window.

**Mechanism.** `COLUMN(range)` only returns the full N-element array under CSE array entry
(Ctrl+Shift+Enter). openpyxl cannot emit the `t="array"` flag — and Excel confirms it treated
these as ordinary formulas by rewriting them as *shared* formulas (`t="shared"`, no array
marker). Under plain entry, legacy `COLUMN(range)` degenerates to its first column, so the
expression collapses to `(3-3+1<=9) * 1 * ISNUMBER(C6)`; `C6` is empty → `0`. Same collapse
in the `C1` scan.

Downstream: `Slope/Variance/Posted/% of Budget/Breach-In`, both `-i` and `+i`, are all wrapped
in `IF(posted_idx=0,"",…)` → blank. `Budget +i` = `INDEX(wp_budget_with_indirect,n)` — no guard,
no array expression — which is why it alone survives. Exactly the reported picture.

**This is the trap already documented in `.claude/rules/openpyxl.md`** ("LibreOffice handles all
approaches; always verify in Excel"). LibreOffice array-evaluates everything, so the idiom
looked correct in every LO-recalced test. `COLUMN(range)` is the one piece of the
`SUMPRODUCT(MAX(…))` idiom that still needs CSE.

### Confirmation test (30 s, in Excel)

Click `BG6`, F2, Ctrl+Shift+Enter. If it flips to 9 and the `+i` block populates → confirmed.
Undo afterwards; the fix belongs in the builder.

### Affected sites (same idiom, same failure expected)

| File:line | Cell |
|---|---|
| `budget_planner/actuals.py:328` | per-WP `posted_idx` |
| `budget_planner/actuals.py:714` | `Actuals!C1` auto-scan |
| `budget_planner/actuals_hours.py:86` | per-row scan |
| `budget_planner/actuals_hours.py:455` | `Actuals-Hours!C1` auto-scan |
| `budget_planner/trends.py:538` | reconciliation scan |

`actuals.py:656` also uses `COLUMN()` but inside a **conditional-formatting** formula, which is
evaluated per cell — not affected, leave it.

`budget_planner/timelog_plan.py:1888` is the counter-example and the proof the fix works: it scans
with `SUMPRODUCT(MAX((grid<>"")*month_idx_range))` against a **literal month-index row**
(`grid_midx_row`) — no `COLUMN()`, no CSE dependency.

---

## Secondary findings

### 1. `Actuals!D1` and `D2` are hardcoded in the user's file

`D1` should hold `=IF(ISNUMBER(B1),B1,C1)`; it holds a literal `9`. `D2` holds a stray literal
`5` in a cell the builder leaves empty. Black-formula cells were typed over.

The `9` is *accidentally load-bearing*: it is the only reason `eur_actuals_through_month` isn't
0 given the broken `C1` scan. After the formula fix, move the `9` into `B1` (the blue override
cell that exists for this) and let `D1` return to being a formula — otherwise the next rebuild
silently resets the trust boundary.

### 2. `scripts/recalc.py` reported success on a workbook it never calculated — **FIXED**

On a cache-stripped copy of the user's file, recalc printed
`{"status": "success", "total_errors": 0, "total_formulas": 4103}` while LibreOffice had
computed **nothing** — every formula cell was an empty `<v></v>` shell, including
`SUM(G6:G12)`. The script counted formulas and scanned for error *strings*; it never asserted
that any value was produced. And `main()` returned exit 0 unconditionally, so
`recalc && pytest` in `make check` never gated at all — even on `errors_found`.

Because `data_only=True` reads `None` for both "never calculated" and "returned empty string",
a green pytest run against such a file proves nothing. This is the documented
"no `<v>` = no calc engine ran" failure, previously unguarded.

**Shipped:**
- `formula_values_present` counted per run; `status: "not_recalculated"` + hint when
  `total_formulas > 0 and values_present == 0`. Deliberately conservative — a formula
  legitimately returning `""` also reads as `None`, so only a workbook where *not one*
  formula produced a value trips it.
- `main()` now exits 1 on `not_recalculated`, `errors_found`, and `error`.

**Demonstrated it can fail** (a gate that cannot return non-zero is not a gate):

| Input | status | exit |
|---|---|---|
| `tmp/probe_real.xlsx` (LO no-ops on it) | `not_recalculated` | **1** |
| a workbook LO computes fine | `success` | 0 |

`make lint` clean, `pytest tests/ -q` → 552 passed.

### 3. Open question, not chased: why LO no-ops on this workbook

LibreOffice computed nothing on the user's Excel-saved file while handling synthetic probes
fine. Not diagnosed. It is the reason `make check` could never have caught the `COLUMN()` bug
on this file — worth its own investigation.

---

## OPEN DECISION — how to remove the CSE dependency

Not implemented, because the choice has real blast-radius differences and **can only be
verified in Excel**, which this environment does not have. Picking wrong reships the same
class of bug (works in LO, blank in Excel).

### Option 1 — literal month-index row (native to this codebase)

Add a hidden row of literals `1..N` under the month header on each affected sheet, exactly as
`timelog_plan.py` already does, and scan against it:
`SUMPRODUCT(MAX((idx_row<=current_month)*idx_row*ISNUMBER(C6:AT6)))`.

- **Buys:** proven in this workbook, in Excel, today. Zero CSE dependence. Same idiom in one
  more place instead of two competing ones.
- **Costs:** inserts a row → `first_wp_row` shifts on `Actuals`, `Actuals-Hours`, `Trends`;
  touches row-position assertions across the test suite. Largest diff of the three.

### Option 2 — inline array constant

Replace `COLUMN(range)-COLUMN($C$1)+1` with a literal `{1,2,…,44}` array constant.

- **Buys:** formula-only change, no layout shift, no row renumbering, smallest diff.
- **Costs:** ~300-character formulas; the constant's length must track `project_months` (a
  silent breakage if they ever drift apart); and array constants written by openpyxl are
  *believed* to evaluate without CSE, but that is precisely the assumption that just failed
  for `COLUMN()`. Unverified in Excel.

### Option 3 — drop the array scan entirely

Postings are cumulative and monotonic, so `Posted` is just the row maximum:
`IF(COUNT(rng)=0,"",MAX(C6:INDEX(C6:AT6,current_month)))` — `INDEX` returning a range
reference needs no CSE, `MAX`/`COUNT` over a plain range works everywhere.

- **Buys:** no array formula anywhere; simplest formulas of the three; nothing to keep in sync.
- **Costs:** loses `posted_idx` as a distinct value, which the slope denominator and the Trends
  chart's `(posted_idx, posted_value)` anchor both consume. Either re-anchor the slope on
  `eur_actuals_through_month` (a sheet-wide cutoff — semantically *different* from a per-WP
  latest posting, and `docs/spec.md` § Two-cadence actuals deliberately distinguishes them) or
  keep one array scan just for the chart anchor and gain little. Contradicts a recorded spec
  invariant; would need Rule 6 sign-off before merge.

**Recommendation: Option 1.** It is the only one with in-workbook evidence behind it, it
converges on an idiom the codebase already proved, and the row-shift cost is mechanical and
test-visible — the failure mode is a red suite, not a silently blank column. Option 2 is the
tempting shortcut that repeats the exact mistake being fixed.

---

## RULING — 2026-07-27 (supersedes the OPEN DECISION block above)

**Option 1 is chosen.** Scope this session is **`actuals.py` only** — Affected-sites rows 1 and 2
(`actuals.py:328` per-WP `posted_idx`, `actuals.py:714` `Actuals!C1` auto-scan).

**Diagnosis is NOT Excel-confirmed.** The Ctrl+Shift+Enter check has not been run; the fix proceeds on
the cached-value evidence alone (Excel's own stored `0` for `BG6`/`C1` with data demonstrably present,
plus `t="shared"` rewriting proving non-array entry). Treat the root cause as strongly-evidenced, not
verified.

### Placement constraint discovered during recon (changes Option 1's shape)

`timelog_plan.py` puts `grid_midx_row` *below* the month-header row. **Actuals cannot do that.**
`budget_planner/readers/actuals.py:71` derives `first_wp_row = _find_header_row(ws, "WP ID") + 1`, so a
literal row inserted directly under the month header would be parsed as the first WP data row on
rebuild — silent data corruption. The index row therefore goes **above the year row** (after the
section-header row). Reader discovery stays header-anchored and is unaffected by the shift.

### Shape

- New `midx_row` written between `_write_section_headers` and `build_year_month_headers` in
  `build_actuals_sheet`; literals `1..N` across `month_col_start..month_col_end`, styled to match
  `timelog_plan.py`'s `grid_midx_row` (grey 8pt Arial, borderless, `"(month idx)"` label in col A)
  and `hidden = True`. Carries an in-cell "do not edit" comment naming the Rule 3 font-colour
  deviation — the cells hold literals but are builder metadata, and the workbook ships unprotected,
  so clearing the row reproduces the exact blank-column symptom it exists to fix. Deliberately NOT an
  outline group: Actuals otherwise has only *column* groups, and a lone grouped row would add a
  row-outline gutter to the sheet.
- Exported as `"midx_row"` in the returned Actuals layout dict; `first_wp_row` and everything below
  shift `+1` automatically (rows are computed sequentially from `r`).
- `posted_idx` → `=SUMPRODUCT(MAX(({midx}<=current_month)*{midx}*ISNUMBER({wp_range})))`
- `Actuals!C1` → `=SUMPRODUCT(MAX({midx}*ISNUMBER({wp_scan_range})))` — the 1×N index row broadcasts
  against the M×N `ISNUMBER` block exactly as legacy `COLUMN()` did, so the scan semantics are a
  faithful port.
- `actuals.py:656` (`COLUMN()` inside conditional formatting) is per-cell evaluated — untouched.

### Consumers verified as shift-safe

`summary.py:105` and `trends.py:505` both read `first_wp_row` out of the returned layout dict; no
module hardcodes an `Actuals!<row>` reference (`grep` for cross-sheet Actuals row literals across
`budget_planner/` returns nothing outside `actuals.py`).

### Explicitly OUT of scope this session

`actuals_hours.py:86`, `actuals_hours.py:455`, `trends.py:538` keep the CSE-dependent `COLUMN()` idiom
and **remain broken in Excel**. Also deferred: the `D1`→`B1` literal move, the static `COLUMN(`-in-
`SUMPRODUCT` guard (Next-action 4a — the sparse-postings *behavioural* test already exists and cannot
catch this class), and the LibreOffice no-op investigation.

## Next actions

1. **User — STILL OWED, not waived.** Run the 30-second Ctrl+Shift+Enter check in Excel. The fix
   shipped without it (see RULING), but it remains the cheapest falsifier of the *mechanism*, and the
   mechanism has an unisolated premise: `SUMPRODUCT` normally propagates array context into a nested
   `COLUMN(range)`, so `=SUMPRODUCT(MAX(COLUMN(A1:E1)))` is generally expected to return 5 without CSE.
   The `t="shared"` evidence proves how Excel *stored* the formula, not how it *evaluated* it. The
   cached `0` is real; "plain entry ⇒ `COLUMN(range)` degenerates *inside* `SUMPRODUCT`" is not yet
   isolated. The fix is structurally safe either way (it removes the dependency rather than reasoning
   about it), but if the premise is wrong the symptom survives and the row shift bought nothing.
2. ~~**User:** pick an option above.~~ Ruled: Option 1 — see RULING.
3. Implement across all five sites; move the `D1` literal back into `B1`.
4. **CORRECTED 2026-07-27 — this item was wrong as originally written.** It claimed the suite has no
   sparse-postings case "which is how this shipped." That is false: `tests/test_actuals_posted.py:128`
   (`gap_recalc_path`) already seeds an explicit `[fill, blank, fill, blank, fill]` at M1/M3/M5, pins
   `current_month = 10`, runs real `scripts/recalc.py`, and asserts `posted_idx == 5` plus
   `posted_value == 500` and `slope == 100`. It was green throughout — and the bug shipped anyway.

   **Why it did not catch it, and why a duplicate would not either:** that fixture recalcs through
   LibreOffice, which array-evaluates regardless of CSE entry. No LO-based assertion can distinguish
   the broken form from the fixed one. Adding another M5+M9 gap test would go green and *close the
   gap on paper while leaving it open*.

   The real owed guard is one of:
   - **(a) static, cheap, pytest-expressible — recommended.** Assert no `COLUMN(` appears in any
     array-forcing cell formula on Actuals: scan every cell whose formula contains `SUMPRODUCT` and
     assert `"COLUMN(" not in it`. This would have gone red on the pre-fix code. It must NOT scan
     conditional-formatting rules — `actuals.py`'s reconciliation-grey CF legitimately uses `COLUMN()`
     and is per-cell evaluated. Extend to `actuals_hours.py` / `trends.py` only once those are fixed,
     or it fails immediately (they still carry the broken idiom).
   - **(b) an Excel-engine check outside pytest.** The only thing that truly closes it. No such
     harness exists in this repo; LibreOffice cannot substitute.
5. Investigate why LibreOffice no-ops on the Excel-saved workbook (finding 3).
6. Then revisit the usability backlog: safety fixes → delta entry → planning block
   (`Remaining` / `Required Run-Rate` / `Forecast at Completion` / `Over-Under at End` /
   `Plan vs Actual`). All of it anchors on `posted_idx`, so none of it can land before this fix.
