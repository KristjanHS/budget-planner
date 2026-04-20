# Programme Budget Domain Reference

Reference for Claude Code agents working on the programme budget Excel generation app. Covers the full programme actual-cost budgeting model: cost categories, funding rates, project lifecycle from proposal to close-out, reallocation rules, and app implementation guidance.

---

## Core Principle

**The EUR budget per partner is the binding financial constraint.** Person-months (PMs) per work package are a planning and justification mechanism --- they derive the EUR amounts, not the other way around.

The causal chain:

```
PM estimate per person x monthly rate = personnel EUR per partner
  + other direct costs (travel, equipment, subcontracting, other)
  = total direct costs per partner
  + indirect costs (25% flat rate on eligible direct costs)
  = total eligible costs per partner
  x funding rate (up to 70%, per the grant agreement)
  = EU grant per partner
```

---

## Cost Categories

### 1. Personnel Costs (typically 50--70% of total budget)

```
personnel_cost = Sum (person_months[i] x monthly_rate[i])
```

Where for each person `i`:
- `person_months[i]` = fraction of time dedicated to the project (max 12 PM/year for full-time)
- `monthly_rate[i]` = annual gross employment cost / 12 (includes salary, social security, employer taxes, benefits)

Staff categories (each has its own rate):
- Senior researcher / principal investigator
- Researcher / engineer
- Developer / technician
- Project manager / admin

**1 PM** = one person working full-time for one month. Actual hours depend on the organisation's working time model (typically 130--150 hours). A person cannot declare more than 12 PMs per year across all EU grants combined.

Monthly rates are organisation-specific. They must reflect actual costs, not inflated project-specific salaries.

**WP active range.** Every WP has a single contiguous active range `[Start, End]` — no multi-segment schedules. Auditors reject PM logged outside this range, so the app treats the intersection of an assignment's window and the WP's active months as the only valid distribution target. The builder redistributes each assignment's `Total FTE` scalar only over in-range months and greys out inactive cells, keeping `Capacity Sum`, `Config Sig Assigned`, and `Plan Assigned PM` reconciled at the user-entered total.

**Vacation months.** Auditors also reject PM logged during vacation. The workbook encodes a single uniform vacation month via the Config `vacation_month` scalar (default `"Jul"`; named range); the derived `num_vacation_months` literal counts calendar-vacation months inside the project span. On Timelog Plan, vacation-month columns are yellow-locked and empty across every person row — the solver never places a WP code there, and the PER-PERSON SUMMARY Pool reduces to `max_fte × (project_months − num_vacation_months)` so the "Over-pool?" check doesn't false-flag people whose Assigned count sits at the available-month ceiling. Heterogeneous per-(person, year) vacation is deferred (Option A); the scalar path is designed to accept an additive override table later with no cell-relocation on Timelog Plan.

### 2. Travel & Subsistence

Direct costs for project-related travel: flights, hotels, per diem. Must be justified per trip or trip category.

### 3. Equipment

Depreciation costs of equipment used for the project (not full purchase price unless 100% dedicated and consumed during project). Must be justified if exceeding 15% of partner's personnel costs.

### 4. Subcontracting

Costs of work outsourced to third parties not part of the consortium. Cannot cover core project tasks. The partner remains responsible for subcontracted work.

### 5. Other Direct Costs

Consumables, software licences, dissemination, IP protection, audit certificates (CFS required if partner claims > EUR 430k). Justify if > 15% of personnel costs.

### 6. Indirect Costs

**25% flat rate** applied to eligible direct costs (excluding subcontracting). This is automatic --- no justification needed, no actual indirect costs tracked.

```
indirect_costs = 0.25 x (personnel + travel + equipment + other_direct)
```

Note: subcontracting is excluded from the indirect cost base.

---

## Funding Rates

| Action Type | Max EU Funding Rate | Notes |
|---|---|---|
| Research actions | Up to 100% | per programme rules |
| Innovation actions | Up to 70% | rest co-financed |

Factors that can affect the funding rate:
- Consortium composition per programme rules

The exact funding rate per partner is specified in the grant agreement.

EU grant calculation:

```
eu_grant = eligible_costs x funding_rate
```

The difference (`eligible_costs - eu_grant`) is the partner's own co-financing obligation.

---

## Project Lifecycle

### Phase 1: PROPOSAL

> Bottom-up estimation -> budget tables -> submission

1. **Estimate effort per partner** `ACTION`
   Each partner estimates PMs per work package, per staff category (see Cost Categories above for staff category list).

2. **Calculate monthly rates** `ACTION`
   Annual gross employment cost / 12 per staff category (see Cost Categories above for rate definition).

3. **Derive personnel EUR** `CALC`
   ```
   personnel_eur = Sum(PM[i] x rate[i])
   ```
   PMs x monthly rate = personnel costs per partner.

4. **Add other direct costs** `ACTION`
   Travel, equipment, subcontracting, other (see Cost Categories above for details on each).

5. **Apply 25% indirect flat rate** `CALC`
   See Indirect Costs under Cost Categories above for the formula and exclusions.

6. **Apply programme funding rate** `CALC`
   See Funding Rates above for the rate table.

7. **Consolidate into Part A + Part B** `OUTPUT`
   Part A: EUR bottom lines per partner. Part B: PM tables per WP per partner + narrative justification.

### Phase 2: GRANT AGREEMENT

> EUR ceilings per partner become legally binding

1. **EUR ceiling locked per partner** `HARD LIMIT`
   Total eligible costs and EU grant amount per partner are fixed. This is the hard constraint for the entire project.

2. **PM distribution recorded** `OUTPUT`
   PMs per WP per partner as stated in Part B. Serves as the baseline plan --- flexible within EUR ceiling.

3. **Funding rate confirmed** `HARD LIMIT`
   Per-partner funding rate is locked.

#### Grant Agreement Structure

- **Part A (binding financials)**: Total EUR per partner for full project duration, broken down by cost category (personnel, travel, equipment, subcontracting, other direct, indirect). This is the ceiling --- cannot be exceeded without amendment.
- **Part B (work plan justification)**: PMs per work package per partner (tables 3.1a, 3.1b), narrative justification of costs. PMs here must be consistent with EUR in Part A.

### Phase 3: PROJECT KICK-OFF

> Detailed planning within the locked EUR envelopes

1. **Detailed resource scheduling** `ACTION`
   Map named people to tasks within WPs. Account for allocation % (e.g., 50% on this project, 50% on other projects).

2. **Build internal timeline** `ACTION`
   Break WP-level PMs into task-level effort estimates with start/end dates and dependencies.

3. **Validate against ceiling** `CHECK`
   Will planned PMs x actual rates fit within the partner EUR envelope? Leave margin for rate increases over the project lifetime.

4. **Set up cost tracking** `ACTION`
   Time sheets, cost recording system. Each person logs hours/days to specific WPs.

### Phase 4: EXECUTION CYCLE

> Repeats per reporting period (typically 12--18 months)

1. **People work, log time to WPs** `ACTION`
   Actual person-days recorded per WP. Monthly declarations or timesheets.

2. **Incur and record other costs** `ACTION`
   Travel, equipment purchases, subcontracts --- all attributed to specific WPs.

3. **Monitor burn rate vs ceiling** `CHECK`
   Compare actual cumulative costs to partner EUR ceiling. Flag if on track to overspend or underspend.

4. **Reallocate PMs between WPs** `FLEXIBLE`
   See Reallocation Rules below for what is allowed without amendment.

5. **Adjust resource scheduling** `FLEXIBLE`
   React to delays, new hires, changed allocation %. Rebalance the plan within the ceiling.

### Phase 5: PERIODIC REPORTING

> Formal cost claim + technical progress report

1. **Compile financial statements** `OUTPUT`
   Actual costs per partner, per category, with supporting docs. Costs must match time records.

2. **Calculate indirect costs** `CALC`
   See Indirect Costs under Cost Categories above --- automatic 25% flat rate, no actuals tracked.

3. **Technical report** `OUTPUT`
   Deliverables achieved, deviations from plan explained, PM consumption per WP reported.

4. **Audit certificate if > EUR 430k** `HARD LIMIT`
   Partners claiming over EUR 430k total must submit an external CFS (Certificate on Financial Statements).

5. **Review & payment** `ACTION`
   EC project officer reviews. May question costs. Approved amounts are reimbursed at the funding rate.

### Phase 6: AMENDMENT (if needed)

> Required only for changes that cross partner boundaries

1. **Budget transfer between partners** `HARD LIMIT`
   Moving EUR from one partner to another requires formal grant agreement amendment.

2. **Major scope changes** `HARD LIMIT`
   Adding/removing WPs, changing core deliverables, adding new partners --- all need amendment.

3. **NOT needed for** `FLEXIBLE`
   See Reallocation Rules below for what is allowed without amendment.

### Phase 7: FINAL REPORTING & CLOSE-OUT

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

## Reallocation Rules (Actual-Cost Model)

- **Within a partner, between cost categories**: Generally flexible, no amendment needed for small shifts. Large structural changes (e.g., moving budget from personnel to subcontracting) may need project officer approval.
- **Between work packages for same partner**: PMs can shift between WPs without amendment, as long as the total EUR stays within the partner ceiling and the work plan is still fulfilled.
- **Between partners**: Requires a formal grant agreement amendment. This is the hard boundary.
- **Increasing total project budget**: Requires amendment and is rarely approved.

---

## Key Constraints for Validation

1. **Partner ceiling**: Sum of all costs per partner <= grant agreement EUR ceiling for that partner
2. **PM consistency**: PMs in budget tables must match PMs in work package descriptions
3. **No double funding**: A person's total declared time across all EU grants <= 12 PM/year
4. **Subcontracting limit**: No core tasks subcontracted; subcontracting must be pre-approved in grant agreement
5. **Purchase justification threshold**: Equipment/other costs > 15% of personnel costs require itemised justification
6. **Audit threshold**: Partners claiming > EUR 430k total need Certificate on Financial Statements (CFS)

---

## Budget Hierarchy in the App

```
Project
+-- Partner (has: total_eur_budget, funding_rate)
|   +-- Cost Category
|   |   +-- Personnel
|   |   |   +-- per Work Package: PMs x rate = EUR
|   |   +-- Travel
|   |   +-- Equipment
|   |   +-- Subcontracting
|   |   +-- Other Direct
|   |   +-- Indirect (calculated: 25% flat rate)
|   +-- Total = Sum all categories <= total_eur_budget
+-- Work Package (has: total PMs across all partners, deliverables, milestones)
    +-- per Partner: PMs allocated, tasks assigned
```

---

## App Implementation Notes

When generating budget Excel sheets:

- The **primary input** is EUR per partner (from grant agreement) --- not PMs
- PMs are derived from: how many person-months at what rate fit within the personnel budget envelope
- Each partner may have different monthly rates (country-specific, seniority-specific)
- The app should validate that: `Sum (PM x rate) + other costs + 25% indirect <= partner EUR ceiling`
- Work package views aggregate PMs across partners to show total effort per WP
- The funding rate converts total eligible costs to the EU grant amount: `eu_grant = eligible_costs x funding_rate`
- The difference (`eligible_costs - eu_grant`) is the partner's own co-financing obligation

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

---

## Key Insight

**EUR per partner is the ceiling. PMs are the lever.** During execution, you're free to reallocate PMs across work packages within your partner envelope. The only hard triggers for a grant amendment are moving money between partners or changing total project scope.
