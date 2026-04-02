# Budget Planner — Feature Brainstorm Findings

## Session: 2026-04-01 (Deep Mode)

### Context Gathered

1. **Biggest pain point**: Planning phase — distributing FTE across WPs and months
2. **Planning approach**: Bottom-up (people's availability) + milestone-driven (deliverable deadlines)
3. **Team**: 4 people, all cross-cutting multiple WPs
4. **Deliverable deadlines**: In the planner's head/emails + grant agreement (DoA)
5. **Key insight**: All 4 people work across MULTIPLE projects with competing deadlines — planning one project in isolation is insufficient
6. **Project load**: Varies — some people on 2 projects, others on 5+
7. **Redmine**: Used as task management tool but currently has only WP/task structure — no content, sizing, or deadlines defined yet

### Emerging Problem Statement

The real planning problem is **cross-project, per-person capacity management** — not just single-project FTE distribution. The current tool solves the output (a compliant budget spreadsheet) but not the input problem (who can work on what, when, given all their commitments).

Additionally, the Redmine instance is underutilized — it has structure but no actionable data (deadlines, effort estimates). There's a potential connection between budget planning and task management that isn't being exploited.

### Priority

1. **Visibility** — single view of per-person load across all projects, spot conflicts early
2. **Compliance** — spreadsheet satisfies auditors and funders
3. **Redmine guidance** — no automation, but help the planner think through what data should go into Redmine (deadlines, effort estimates, etc.)

### Replanning Frequency

- **Monthly or more** — constant juggling as deadlines shift
- Deadlines shift frequently due to large number of consortium partners
- Current tool doesn't support quick replanning — need to manually redistribute FTE each time

### Core Tension

The tool generates a nice compliant spreadsheet, but the PLANNING INPUTS (who is available when, which deadlines are shifting) live in the planner's head. Every replan means mentally recalculating 4 people × N projects × shifting deadlines, then manually updating the spreadsheet.

### Ideation Results

#### Top 3 Blind Spots (from Inversion)

1. **The unit of planning is the person, not the WP.** The tool organizes around WP×month (compliance artifact), but the actual constraint is person×month capacity. The person-load view isn't a reporting feature — it's the missing planning surface.

2. **The scratchpad lives in the planner's head.** The bottleneck isn't editing speed — it's fear of overwriting a valid plan with an unvalidated one. A draft/baseline toggle externalizes the working memory that currently lives nowhere.

3. **Deadline slips are invisible at decision time.** When a WP date shifts, the FTE redistribution cost is never made concrete. Showing the cell-level diff at change time converts abstract risk into a visible decision.

#### User Validation

- Blind spots 1 and 2 resonate strongly
- Blind spot 3 (deadline slip cost) is less relevant — WP-level deadlines don't shift, but intermediate deliverables cluster toward WP end dates, creating effort spikes
- The real pattern: deliverables pile up near WP deadlines, making certain months disproportionately loaded per person

#### Chosen Direction

- **B (Cross-Project Capacity)** + **C (Draft/Baseline)** — both wanted
- Other projects are simple: single FTE commitment line per person (someone else manages the details)
- All data must live in Excel — no YAML/Python config for other projects
- This means: a "Capacity" sheet (or separate workbook) where other projects appear as simple rows (project name, person, FTE/month, start month, end month), while the programme project pulls its detail from the WP-level Plan sheet

### Approved Design

#### New Sheet: Capacity

**Block 1 — Other Commitments** (per person-project, monthly grid)
- Config columns (blue input): Person, Project, Total FTE, Start Month, End Month
- Monthly grid M1..M60: auto-filled with even distribution on first gen, then user-editable
- `--update` preserves manual edits (same round-trip as Plan)
- Full regen without `--update`: CLI double-confirms before overwriting external commitments

**Block 2 — Person-Load Overview** (formula cells)
- Rows = people, Columns = months
- Each cell = SUM(person's programme WP rows from Plan) + SUM(person's external commitment rows)
- Red fill > 1.0, amber fill > 0.8
- Remaining capacity row per person: 1.0 - total

#### New Sheet: Draft

- Copy of Planned FTE block, pre-filled from baseline, fully editable
- Diff row per WP: Draft - Baseline, yellow fill where non-zero
- Capacity sheet has toggle cell `use_draft` (named range): TRUE reads from Draft, FALSE reads from Plan
- `--promote-draft` CLI flag: copies Draft → Plan baseline, clears Draft
- `--update` preserves both baseline and draft edits

#### Config Change: Person-WP Assignments

- config.py needs person definitions and which WPs each person contributes to + share

#### Perspective Highlights

- **Planner**: Person-centric Load sheet with remaining capacity per month, red on >1.0 FTE
- **Team member**: Per-person commitment card they can review and flag conflicts on
- **Auditor**: _SOURCES sheet with evidence chain + completeness flags for missing actuals
- **Future planner**: Replan log capturing what changed, when, and why (_LOG sheet + --note flag)
