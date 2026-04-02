# Programme Actual-Cost Budgeting Logic

Reference document for Claude Code agents working on the programme budget Excel generation app.

## Core Principle

**The EUR budget per partner is the binding financial constraint.** Person-months (PMs) per work package are a planning and justification mechanism — they derive the EUR amounts, not the other way around.

The causal chain:

```
PM estimate per person × monthly rate = personnel EUR per partner
  + other direct costs (travel, equipment, subcontracting, other)
  = total direct costs per partner
  + indirect costs (25% flat rate on eligible direct costs)
  = total eligible costs per partner
  × funding rate (up to 70%, per the grant agreement)
  = EU grant per partner
```

## Grant Agreement Structure

### Part A (binding financials)
- Total EUR per partner for full project duration
- Broken down by cost category (personnel, travel, equipment, subcontracting, other direct, indirect)
- This is the ceiling — you cannot exceed it without a grant amendment

### Part B (work plan justification)
- PMs per work package per partner (tables 3.1a, 3.1b)
- Narrative justification of costs
- PMs here must be consistent with EUR in Part A

## Cost Categories

### 1. Personnel Costs (typically 50–70% of total budget)

```
personnel_cost = Σ (person_months[i] × monthly_rate[i])
```

Where for each person `i`:
- `person_months[i]` = fraction of time dedicated to the project (max 12 PM/year for full-time)
- `monthly_rate[i]` = annual gross employment cost ÷ 12 (includes salary, social security, employer taxes, benefits)

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
**25% flat rate** applied to eligible direct costs (excluding subcontracting and some other categories). This is automatic — no justification needed, no actual indirect costs tracked.

```
indirect_costs = 0.25 × (personnel + travel + equipment + other_direct)
```

Note: subcontracting is typically excluded from the indirect cost base.

## Programme-Specific Funding Rates

| Action Type | Max EU Funding Rate | Notes |
|---|---|---|
| Research actions | Up to 100% | per programme rules |
| Innovation actions | Up to 70% | rest co-financed |

Factors that can affect the funding rate:
- Consortium composition per programme rules

The exact funding rate per partner is specified in the grant agreement.

## Budget Hierarchy in the App

```
Project
├── Partner (has: total_eur_budget, funding_rate)
│   ├── Cost Category
│   │   ├── Personnel
│   │   │   └── per Work Package: PMs × rate = EUR
│   │   ├── Travel
│   │   ├── Equipment
│   │   ├── Subcontracting
│   │   ├── Other Direct
│   │   └── Indirect (calculated: 25% flat rate)
│   └── Total = Σ all categories ≤ total_eur_budget
└── Work Package (has: total PMs across all partners, deliverables, milestones)
    └── per Partner: PMs allocated, tasks assigned
```

## Key Constraints for Validation

1. **Partner ceiling**: Sum of all costs per partner ≤ grant agreement EUR ceiling for that partner
2. **PM consistency**: PMs in budget tables must match PMs in work package descriptions
3. **No double funding**: A person's total declared time across all EU grants ≤ 12 PM/year
4. **Subcontracting limit**: No core tasks subcontracted; subcontracting must be pre-approved in grant agreement
5. **Purchase justification threshold**: Equipment/other costs > 15% of personnel costs require itemised justification
6. **Audit threshold**: Partners claiming > €430k total need Certificate on Financial Statements (CFS)

## Reallocation Rules (Actual-Cost Model)

- **Within a partner, between cost categories**: Generally flexible, no amendment needed for small shifts. Large structural changes (e.g., moving budget from personnel to subcontracting) may need project officer approval.
- **Between work packages for same partner**: PMs can shift between WPs without amendment, as long as the total EUR stays within the partner ceiling and the work plan is still fulfilled.
- **Between partners**: Requires a formal grant agreement amendment. This is the hard boundary.
- **Increasing total project budget**: Requires amendment and is rarely approved.

## Reporting (Actual-Cost Model)

Each reporting period (typically 12–18 months):
1. **Technical report**: Work done, deliverables achieved, deviations explained
2. **Financial report**: Actual costs incurred per partner, per category, with supporting documentation
3. Costs are checked against eligibility rules and grant agreement ceilings
4. Partners claiming > €430k submit an external audit certificate (CFS)

## App Implementation Notes

When generating budget Excel sheets:

- The **primary input** is EUR per partner (from grant agreement) — not PMs
- PMs are derived from: how many person-months at what rate fit within the personnel budget envelope
- Each partner may have different monthly rates (country-specific, seniority-specific)
- The app should validate that: `Σ (PM × rate) + other costs + 25% indirect ≤ partner EUR ceiling`
- Work package views aggregate PMs across partners to show total effort per WP
- The funding rate converts total eligible costs to the EU grant amount: `eu_grant = eligible_costs × funding_rate`
- The difference (`eligible_costs - eu_grant`) is the partner's own co-financing obligation
