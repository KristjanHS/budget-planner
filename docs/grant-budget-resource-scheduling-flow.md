# Programme Actual-Cost Budget & Resource Scheduling Flow

Reference for Claude Code agents working on the programme budget Excel generation app. Describes the full lifecycle of budgeting and resource scheduling activities from proposal to close-out.

---

## Phase 1: PROPOSAL

> Bottom-up estimation → budget tables → submission

1. **Estimate effort per partner** `ACTION`
   Each partner estimates PMs per work package, per staff category (senior researcher, engineer, dev, admin).

2. **Calculate monthly rates** `ACTION`
   Annual gross employment cost ÷ 12 per staff category. Includes salary, social security, employer taxes.

3. **Derive personnel EUR** `CALC`
   ```
   personnel_eur = Σ(PM[i] × rate[i])
   ```
   PMs × monthly rate = personnel costs per partner.

4. **Add other direct costs** `ACTION`
   Travel, equipment, subcontracting, other (software, dissemination, IP, audit cert).

5. **Apply 25% indirect flat rate** `CALC`
   ```
   indirect = 0.25 × (personnel + travel + equip + other)
   ```
   Subcontracting is excluded from the indirect cost base.

6. **Apply programme funding rate** `CALC`
   ```
   eu_grant = eligible_costs × funding_rate
   ```
   Innovation actions: up to 70%. The rest is partner co-financing.

7. **Consolidate into Part A + Part B** `OUTPUT`
   Part A: EUR bottom lines per partner. Part B: PM tables per WP per partner + narrative justification.

---

## Phase 2: GRANT AGREEMENT

> EUR ceilings per partner become legally binding

1. **EUR ceiling locked per partner** `HARD LIMIT`
   Total eligible costs and EU grant amount per partner are fixed. This is the hard constraint for the entire project.

2. **PM distribution recorded** `OUTPUT`
   PMs per WP per partner as stated in Part B. Serves as the baseline plan — flexible within EUR ceiling.

3. **Funding rate confirmed** `HARD LIMIT`
   Per-partner funding rate is locked.

---

## Phase 3: PROJECT KICK-OFF

> Detailed planning within the locked EUR envelopes

1. **Detailed resource scheduling** `ACTION`
   Map named people to tasks within WPs. Account for allocation % (e.g., 50% on this project, 50% on other projects).

2. **Build internal timeline** `ACTION`
   Break WP-level PMs into task-level effort estimates with start/end dates and dependencies.

3. **Validate against ceiling** `CHECK`
   Will planned PMs × actual rates fit within the partner EUR envelope? Leave margin for rate increases over the project lifetime.

4. **Set up cost tracking** `ACTION`
   Time sheets, cost recording system. Each person logs hours/days to specific WPs.

---

## Phase 4: EXECUTION CYCLE

> Repeats per reporting period (typically 12–18 months)

1. **People work, log time to WPs** `ACTION`
   Actual person-days recorded per WP. Monthly declarations or timesheets.

2. **Incur and record other costs** `ACTION`
   Travel, equipment purchases, subcontracts — all attributed to specific WPs.

3. **Monitor burn rate vs ceiling** `CHECK`
   Compare actual cumulative costs to partner EUR ceiling. Flag if on track to overspend or underspend.

4. **Reallocate PMs between WPs** `FLEXIBLE`
   Shift effort between work packages as needed. No amendment needed if within partner ceiling and work plan still delivered.

5. **Adjust resource scheduling** `FLEXIBLE`
   React to delays, new hires, changed allocation %. Rebalance the plan within the ceiling.

---

## Phase 5: PERIODIC REPORTING

> Formal cost claim + technical progress report

1. **Compile financial statements** `OUTPUT`
   Actual costs per partner, per category, with supporting docs. Costs must match time records.

2. **Calculate indirect costs** `CALC`
   ```
   indirect_claim = 0.25 × (personnel + travel + equip + other)
   ```
   Automatic — no actuals tracked for indirect.

3. **Technical report** `OUTPUT`
   Deliverables achieved, deviations from plan explained, PM consumption per WP reported.

4. **Audit certificate if > €430k** `HARD LIMIT`
   Partners claiming over €430k total must submit an external CFS (Certificate on Financial Statements).

5. **Review & payment** `ACTION`
   EC project officer reviews. May question costs. Approved amounts are reimbursed at the funding rate.

---

## Phase 6: AMENDMENT (if needed)

> Required only for changes that cross partner boundaries

1. **Budget transfer between partners** `HARD LIMIT`
   Moving EUR from one partner to another requires formal grant agreement amendment.

2. **Major scope changes** `HARD LIMIT`
   Adding/removing WPs, changing core deliverables, adding new partners — all need amendment.

3. **NOT needed for** `FLEXIBLE`
   Shifting PMs between WPs within same partner. Moving budget between cost categories within same partner (unless structural change like converting personnel to subcontracting).

---

## Phase 7: FINAL REPORTING & CLOSE-OUT

> Final cost claim, reconciliation, audit

1. **Final financial statements** `OUTPUT`
   All remaining costs claimed. Total per partner must not exceed EUR ceiling from grant agreement.

2. **Reconciliation** `CHECK`
   Compare total actual costs to budget. Underspend stays with EC. Overspend is absorbed by partner.

3. **Final technical report** `OUTPUT`
   All deliverables accounted for. PM actuals vs plan comparison.

4. **Final payment** `ACTION`
   EC releases remaining grant after review. Possible ex-post audit within 5 years.

---

## Key Insight

**EUR per partner is the ceiling. PMs are the lever.** During execution, you're free to reallocate PMs across work packages within your partner envelope. The only hard triggers for a grant amendment are moving money between partners or changing total project scope.

---

## Step Type Reference

| Tag | Meaning | App implication |
|---|---|---|
| `ACTION` | Human activity, no computation | May need UI input or trigger |
| `CALC` | Deterministic formula | Must be implemented exactly |
| `OUTPUT` | Produces a deliverable or data artifact | App should generate this |
| `HARD LIMIT` | Cannot be violated without amendment | Enforce as validation rule |
| `CHECK` | Monitoring / warning condition | Implement as warning, not blocker |
| `FLEXIBLE` | Allowed reallocation within constraints | App should permit without warnings |
