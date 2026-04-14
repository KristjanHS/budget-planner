# Design Invariants

Cross-sheet invariants, trust boundaries, and derived-metric rules that must hold across code, tests, and docs. When a new structural invariant lands, record it here — future agents need it in loaded context, not buried in a sheet builder.

## Two-cadence actuals (hours vs EUR)

The workbook records actuals on two separate cadences. Any derived view that reads actual data must pick the correct trust boundary — mixing them produces wrong answers that look right.

### Hours cadence

Per-person × per-WP × per-month cells on `Actuals-Hours`.

- **Granularity:** monthly, per person per WP.
- **Entry convention:** skip-and-cumulate is valid. A user may leave intermediate months blank and enter the cumulative hours-to-date in a later month. This means a single-month cell is **NOT** authoritative for that month — only the running cumulative total is.
- **Trust boundary:** `hours_actuals_through_month` (`Config!$B$8`, blue input integer). Months `1..hours_actuals_through_month` are closed for hours; later months are forecast-only territory.

### EUR cadence

WP-level totals on the `Actuals` sheet.

- **Granularity:** monthly, per WP (no per-person EUR on `Actuals`).
- **Entry convention:** driven by accountant/finance close, which runs with a 2–3 month lag behind the hours timelog.
- **Trust boundary:** `eur_actuals_through_month` (`Config!$B$7`, blue input integer). Months past this index are not yet reconciled.

### Derived boundary

- `reconciled_through_month = MIN(eur_actuals_through_month, hours_actuals_through_month)` (`Config!$B$11`, black formula).
- This is the authoritative boundary for any **head-to-head comparison** between Actual EUR and Forecast EUR (hours-derived). Past this month at least one side is stale or missing.

### Design principle (enforce on every new actual-side readout)

No derived view reads a single-month cell as authoritative for that month. Every actual-side readout is **either**:

- **(a)** a cumulative-through-a-through-month total, or
- **(b)** a to-date **rate** (cumulative / `months_elapsed_*`).

Raw monthly actual cells surface only in collapsed / clearly-labelled bands ("reliable only when hours were entered per-month"), never as the input to a derived metric.

### Forward-note (hours-timelog v2 rollout)

Stage 1 (complete) wires the Config inputs + named ranges only; no consumer code reads them yet — so if you grep for `hours_actuals_through_month` and find only Config + docs, that is expected mid-rollout, not drift.

Stage 2 (W2) adds a to-date column block on Capacity, switches Breach-In to a cumulative-slope formula, drives per-column partial-year shading on Summary by the stricter of the two through-months, and adds cosmetic grey shading past the through-month on `Actuals-Hours` + `Plan-Forecast`.

Stage 3 (W3) adds a Reconciliation section on Trends: three health cells (EUR through / Hours through / Reconciled through), a per-WP summary table with cumulative Forecast vs Actual EUR + gap %, and a gap-over-time LineChart clipped at `reconciled_through_month`.

## Future invariants (reserved)

Placeholders for invariants that have a home elsewhere today but belong here when the next cross-sheet rule lands. Write them up when the next invariant accretes — don't flesh them out speculatively.

- **Font color convention** (Critical Rule 3 pointer): Blue = input, Black = formula, Green = cross-sheet ref.
- **Named-range contract**: `project_start`, `fte_rate`, `hours_actuals_through_month`, `eur_actuals_through_month`, `reconciled_through_month`, `months_elapsed_*` — what each guarantees, who writes it, who reads it.
- **Reader/builder round-trip rule**: readers use named ranges (not coordinates); tolerate internal blanks; rebuild must preserve user-authored values on blue input cells.
