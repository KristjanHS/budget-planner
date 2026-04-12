# Ideas to Explore (from Research A & B)

**Date**: 2026-04-05
**Status**: Unrefined — brainstorm before implementing any of these

Items below are NOT in the P1-P6 backlog. They surfaced during research review but weren't prioritised. Revisit when current backlog is done.

## From Research A (Compliance)

### Split-Time Totality Check
Beyond the 215-day EU ceiling (P1, done), verify: programme days + non-programme days + leave = total working days per month. Flags >100% allocation including non-project time. Different from capacity overallocation — this is the audit cross-check.
- *Research A, Section 2.3-2.4*

### Assignment Proof Documentation Helper
EU auditors require an "internal management decision" linking each person to the action. A printable summary from the Capacity sheet (person → WPs → periods → FTE) could serve as that proof or at least a starting template.
- *Research A, Section 1.4, item 4*

### Mid-Year Rate Changes
Grade-based rates (P6) use a single annual rate. If someone gets promoted mid-year, the rate should change. Could add a "rate effective from" mechanism per person. Low priority — annual averaging is acceptable.
- *Research A, Section "P2: Rate Calculation Support"*

### Part-Time Factor
Part-time employees have a reduced 215-day ceiling (e.g., 0.5 FTE → 107.5 days max). The EU ceiling check (P1) currently uses a flat 12 PM ceiling. Could scale by Max FTE per person.
- *Research A, Section "P1: Time Allocation Guardrails"*

## From Research B (Tools & Workflows)

### Reallocation Log Sheet
Simple structured table: Date | From WP | To WP | PM Amount | Reason. No formulas — just an audit trail for PM reallocation decisions. Required for periodic reports.
- *Research B, Section 3.6, item 4*

### Indirect Cost Auto-Calculation
Formula: `0.25 × (Personnel + Travel + Equipment + Other)`. Simple but requires non-personnel cost category inputs that don't exist yet. Only useful if user starts tracking other cost categories.
- *Research B, "Lower-Priority Enhancements" table*

### Exhaustion Date Confidence Bands
Enhance existing burn rate projection with optimistic/pessimistic bands based on variance history. More relevant for longer-running projects.
- *Research B, Section 3.6, item 3*

### Earned Value Formulas (SPI, CPI)
Standard in regulated PM but rarely used in EU research projects. programme straddles both worlds. Could add as optional section in Trends. Low priority.
- *Research B, Section 3.4*

### Cost Category Breakdown Beyond Personnel
Add travel, equipment, subcontracting, other direct costs to Actuals. Would enable Form C data prep and indirect cost calculation. High effort — restructures Actuals fundamentally.
- *Research B, "Lower-Priority Enhancements" table*

## Anti-Patterns (from Research B)
Keep in mind when evaluating the above:
- Don't try to replace the timesheet system
- Don't replicate EMDESK's consortium view
- Don't over-automate Form C generation
