# Research-Driven Enhancements Plan

**Date**: 2026-04-05
**Status**: Research complete, implementation not started

## Goals

Improve the programme budget planner to:

1. **Prevent audit findings** — build guardrails for compliance risks that a first-time EU PM wouldn't know to watch for
2. **Surface cross-project overcommitment** — the same people split time across programme, local product dev, and other projects; the tool currently only sees the programme slice
3. **Align with EU reporting workflow** — structure data to match what's needed at reporting period boundaries (Form C, deviation narratives, CFS thresholds)

## Context

- The user is a **first-time programme project manager** who also manages the same people across local product development and other parallel projects
- Team members need to know **which WPs/tasks they can log hours to** in any given month — currently they ask the PM
- The biggest unknown-unknown risk is **auditor expectations** around personnel costs (70%+ of EU grant audit errors)
- Programme uses a standard EU Model Grant Agreement for personnel cost rules; HE audit experience transfers directly

## Research Outputs

Two research reports were produced and saved:

| Report | File | Focus |
|--------|------|-------|
| **A: Compliance & Audit Risks** | `docs/research/research-a-grant-compliance-risks.md` | Timesheet evidence, split-time docs, WP eligibility, reallocation rules, common PM mistakes |
| **B: Tools & Workflows** | `docs/research/research-b-eu-pm-tools-workflows.md` | Tool landscape (EMDESK, Kronis, Tempo), capacity views, variance workflows, reporting periods |

Read the specific report section when implementing a related feature — don't load both fully into a single session.

## Prioritised Enhancement List

Ordered by value-to-effort ratio. Each item is scoped for a single implementation session.

### P1: 215-Day Ceiling + Cross-Project Capacity Check

**Risk addressed**: Person declares >215 day-equivalents across all EU grants in a calendar year = costs rejected, possible fraud referral.

**What to build**:
- "Other EU Projects" and "Non-EU Work" input rows per person (monthly PM values, blue font = input)
- Annual total formula per person: programme PM + Other EU PM
- CF: Red if annual EU total > 12 PM (~215 days), Amber if > 11 PM
- Named range for the ceiling value (Config sheet)
- Monthly total row: programme + Other + Non-EU, CF Red if > 1.0 FTE in any month

**Reference**: Research A section 2 (split-time), section 5.2 mistake #4 (215-day ceiling). Research B section 2 (cross-project capacity views).

**Key rule**: 215 days / ~17.92 working days per month = ~12 PM per year. Part-time employees: pro-rate by contract FTE.

### P2: Plan vs. Actual Variance Classification

**Risk addressed**: Deviations creep unnoticed until reporting deadline, when it's too late for corrective action.

**What to build**:
- Variance formula row per WP: `(Actual - Planned) / Planned`
- CF thresholds: Green (<5%), Amber (5-10%), Red (>10%), Critical (>20%)
- Cumulative variance column per reporting period
- The >20% threshold on cost category transfers flags amendment risk

**Reference**: Research B section 3 (commitment vs actuals workflows), section 3.2 (standard thresholds). Research A section 4 (20% budget transfer rule).

### P3: Reporting Period Aggregation

**Risk addressed**: Manual summing of monthly columns for Form C preparation is error-prone and slow.

**What to build**:
- Reporting period boundaries in Config (e.g., RP1: M1-M12, RP2: M13-M24)
- Summary section aggregating per-WP personnel PM and EUR across each reporting period
- Match Form C structure: cost category A (personnel), E (indirect at 25%)
- CFS threshold indicator: flag when cumulative costs approach EUR 430k

**Reference**: Research B section 5 (reporting period structure, Form C, CFS threshold). Research A section 5.3 #3 (indirect cost choice).

### P4: Per-Person WP Eligibility View

**Risk addressed**: Team member logs hours to wrong WP or outside active period = ineligible costs.

**What to build**:
- Read-only reference sheet (or section): Person x Month grid
- Cell formula: person has Capacity allocation > 0 AND month is within WP active period
- CF: Green = can log, Grey = WP inactive, White = not assigned
- Optional: remaining budget indicator (Capacity allocation minus Actuals)

**Reference**: Research B section 4 (per-person eligibility). Research A section 3 (WP eligibility boundaries, Annex 1 link).

### P5: Prior Period Adjustment Support

**Risk addressed**: Cumulative totals in the planner won't match official financial statements if corrections aren't tracked.

**What to build**:
- "Adjustments" row per WP in Actuals (or dedicated section)
- Feeds into cumulative totals
- Distinct styling (italic or different background) to separate from current-period actuals
- Can be positive or negative

**Reference**: Research B section 3.5 (prior period adjustments).

## Anti-Patterns (Do Not Build)

- **Don't replace the timesheet system** — the planner is planning/monitoring, not time recording
- **Don't replicate EMDESK's consortium view** — focus on single-partner internal planning
- **Don't hardcode reporting period dates** — make configurable in Config
- **Don't over-automate Form C** — prepare data, not the official submission form

## How to Use This Plan

1. Pick one P-item per session
2. Read the referenced research report sections (not the full reports)
3. Read the relevant deep-dive docs listed in CLAUDE.md's "Before Modifying Code" table
4. Implement, test, verify
5. Mark the item as done below

## Implementation Status

| Item | Status | Date | Notes |
|------|--------|------|-------|
| P1: 215-Day Ceiling | Done | 2026-04-05 | P1a + P1b implemented, plans in old_already_implemented/ |
| P2: Variance Classification | Won't implement | 2026-04-05 | Low value vs effort; PM can eyeball variance from existing Plan/Actuals columns |
| P3: Reporting Period Aggregation | Plan created | 2026-04-05 | `2026-04-05-p3-reporting-period-aggregation.md` — 9 tasks |
| P4: WP Eligibility View | Not started | | |
| P5: Prior Period Adjustments | Not started | | |
