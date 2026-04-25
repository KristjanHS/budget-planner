# Programme Project Lifecycle — human reference

Background reading for someone learning the programme actual-cost budgeting model. Not loaded into agent context (this directory is `.claudeignore`'d). For the rules that actually gate code, see `docs/spec.md` § programme domain rules.

---

## Core Principle

**The EUR budget per partner is the binding financial constraint.** Person-months (PMs) per work package are a planning and justification mechanism — they derive the EUR amounts, not the other way around.

```
PM estimate per person × monthly rate = personnel EUR per partner
  + other direct costs (travel, equipment, subcontracting, other)
  = total direct costs per partner
  + indirect costs (25% flat rate on eligible direct costs)
  = total eligible costs per partner
  × funding rate (up to 70%, per the grant agreement)
  = EU grant per partner
```

---

## Cost Categories

### 1. Personnel Costs (typically 50–70% of total budget)

```
personnel_cost = Σ(person_months[i] × monthly_rate[i])
```

- `person_months[i]` = fraction of time dedicated to the project (max 12 PM/year for full-time)
- `monthly_rate[i]` = annual gross employment cost / 12 (includes salary, social security, employer taxes, benefits)

Staff categories (each has its own rate):
- Senior researcher / principal investigator
- Researcher / engineer
- Developer / technician
- Project manager / admin

**1 PM** = one person working full-time for one month. Actual hours depend on the organisation's working time model (typically 130–150 hours). A person cannot declare more than 12 PMs per year across all EU grants combined.

Monthly rates are organisation-specific. They must reflect actual costs, not inflated project-specific salaries.

### 2. Travel & Subsistence

Direct costs for project-related travel: flights, hotels, per diem. Must be justified per trip or trip category.

### 3. Equipment

Depreciation costs of equipment used for the project (not full purchase price unless 100% dedicated and consumed during project). Must be justified if exceeding 15% of partner's personnel costs.

### 4. Subcontracting

Costs of work outsourced to third parties not part of the consortium. Cannot cover core project tasks. The partner remains responsible for subcontracted work.

### 5. Other Direct Costs

Consumables, software licences, dissemination, IP protection, audit certificates (CFS required if partner claims > €430k). Justify if > 15% of personnel costs.

### 6. Indirect Costs

**25% flat rate** applied to eligible direct costs (excluding subcontracting). Automatic — no justification needed, no actual indirect costs tracked.

```
indirect_costs = 0.25 × (personnel + travel + equipment + other_direct)
```

Subcontracting is excluded from the indirect cost base.

---

## Funding Rates

| Action Type | Max EU Funding Rate | Notes |
|---|---|---|
| Research actions | Up to 100% | per programme rules |
| Innovation actions | Up to 70% | rest co-financed |

Factors that can affect the funding rate:
- Consortium composition per programme rules

The exact funding rate per partner is specified in the grant agreement.

```
eu_grant = eligible_costs × funding_rate
```

The difference (`eligible_costs − eu_grant`) is the partner's own co-financing obligation.

---

## Project Lifecycle

### Phase 1: PROPOSAL

> Bottom-up estimation → budget tables → submission

1. **Estimate effort per partner** `ACTION` — Each partner estimates PMs per work package, per staff category.
2. **Calculate monthly rates** `ACTION` — Annual gross employment cost / 12 per staff category.
3. **Derive personnel EUR** `CALC` — `personnel_eur = Σ(PM[i] × rate[i])`.
4. **Add other direct costs** `ACTION` — Travel, equipment, subcontracting, other.
5. **Apply 25% indirect flat rate** `CALC`.
6. **Apply programme funding rate** `CALC`.
7. **Consolidate into Part A + Part B** `OUTPUT` — Part A: EUR bottom lines per partner. Part B: PM tables per WP per partner + narrative justification.

### Phase 2: GRANT AGREEMENT

> EUR ceilings per partner become legally binding

1. **EUR ceiling locked per partner** `HARD LIMIT` — Total eligible costs and EU grant amount per partner are fixed for the entire project.
2. **PM distribution recorded** `OUTPUT` — PMs per WP per partner as stated in Part B. Baseline plan, flexible within EUR ceiling.
3. **Funding rate confirmed** `HARD LIMIT` — Per-partner funding rate locked.

#### Grant Agreement Structure

- **Part A (binding financials)**: Total EUR per partner for full project duration, broken down by cost category. This is the ceiling — cannot be exceeded without amendment.
- **Part B (work plan justification)**: PMs per work package per partner (tables 3.1a, 3.1b), narrative justification of costs. PMs here must be consistent with EUR in Part A.

### Phase 3: PROJECT KICK-OFF

> Detailed planning within the locked EUR envelopes

1. **Detailed resource scheduling** `ACTION` — Map named people to tasks within WPs. Account for allocation %.
2. **Build internal timeline** `ACTION` — Break WP-level PMs into task-level estimates with start/end dates and dependencies.
3. **Validate against ceiling** `CHECK` — Will planned PMs × actual rates fit within partner EUR envelope? Leave margin for rate increases.
4. **Set up cost tracking** `ACTION` — Time sheets, cost recording system. Each person logs hours/days to specific WPs.

### Phase 4: EXECUTION CYCLE

> Repeats per reporting period (typically 12–18 months)

1. **People work, log time to WPs** `ACTION` — Actual person-days recorded per WP. Monthly declarations or timesheets.
2. **Incur and record other costs** `ACTION` — Travel, equipment purchases, subcontracts attributed to specific WPs.
3. **Monitor burn rate vs ceiling** `CHECK` — Compare actual cumulative costs to partner EUR ceiling. Flag if on track to over/underspend.
4. **Reallocate PMs between WPs** `FLEXIBLE` — See Reallocation Rules below.
5. **Adjust resource scheduling** `FLEXIBLE` — React to delays, new hires, changed allocation %.

### Phase 5: PERIODIC REPORTING

> Formal cost claim + technical progress report

1. **Compile financial statements** `OUTPUT` — Actual costs per partner, per category, with supporting docs.
2. **Calculate indirect costs** `CALC` — Automatic 25% flat rate.
3. **Technical report** `OUTPUT` — Deliverables, deviations, PM consumption per WP.
4. **Audit certificate if > €430k** `HARD LIMIT` — Partners claiming over €430k total must submit external CFS.
5. **Review & payment** `ACTION` — EC project officer reviews. Approved amounts reimbursed at funding rate.

### Phase 6: AMENDMENT (if needed)

> Required only for changes that cross partner boundaries

1. **Budget transfer between partners** `HARD LIMIT` — Requires formal grant agreement amendment.
2. **Major scope changes** `HARD LIMIT` — Adding/removing WPs, changing core deliverables, adding new partners.
3. **NOT needed for** `FLEXIBLE` — See Reallocation Rules below.

### Phase 7: FINAL REPORTING & CLOSE-OUT

1. **Final financial statements** `OUTPUT` — All remaining costs claimed. Total per partner ≤ EUR ceiling.
2. **Reconciliation** `CHECK` — Compare total actual costs to budget. Underspend stays with EC. Overspend absorbed by partner.
3. **Final technical report** `OUTPUT` — All deliverables accounted for. PM actuals vs plan.
4. **Final payment** `ACTION` — EC releases remaining grant after review. Possible ex-post audit within 5 years.

---

## Reallocation Rules (Actual-Cost Model)

- **Within a partner, between cost categories**: Generally flexible; large structural changes may need project officer approval.
- **Between work packages for same partner**: PMs can shift between WPs without amendment, as long as total EUR stays within partner ceiling.
- **Between partners**: Requires formal grant agreement amendment. **Hard boundary.**
- **Increasing total project budget**: Requires amendment, rarely approved.

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

**EUR per partner is the ceiling. PMs are the lever.** During execution, free to reallocate PMs across work packages within your partner envelope. The only hard triggers for grant amendment are moving money between partners or changing total project scope.
