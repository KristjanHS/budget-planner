# Architecture Diagrams

## System Context (C4 Level 1)

```mermaid
C4Context
  title System Context
  Person(pm, "Project Manager", "Plans FTE, enters actuals")
  Person(accountant, "Accountant", "Provides EUR actuals")
  System(grant, "Budget Planner", "Python CLI + openpyxl")
  System_Ext(excel, "Excel / LibreOffice", "Spreadsheet editor")
  System_Ext(soffice, "LibreOffice headless", "Formula recalculation")
  Rel(pm, grant, "Runs CLI commands")
  Rel(grant, excel, "Produces .xlsx workbook")
  Rel(pm, excel, "Edits Capacity FTE, reviews dashboards")
  Rel(accountant, excel, "Enters Actuals EUR")
  Rel(grant, soffice, "Recalculates formulas")
```

## Component Architecture (C4 Level 3)

```mermaid
graph TB
  subgraph "CLI Layer"
    CLI["__main__.py"]
  end
  subgraph "Orchestrator"
    WB["workbook.py"]
  end
  subgraph "Sheet Builders"
    CFG_S["config_sheet.py"] & PLAN["plan.py"] & CAP["capacity.py"]
    ACT["actuals.py"] & SUM["summary.py"] & TRENDS["trends.py"] & META["meta.py"]
  end
  subgraph "Shared Infrastructure"
    LAYOUT["layout.py"] & STYLES["styles.py"]
    CONFIG["config.py"] & READER["reader.py"]
  end
  CLI --> WB
  WB --> CFG_S & PLAN & CAP & ACT & SUM & TRENDS & META
  WB -->|"init"| CONFIG
  WB -->|"rebuild"| READER
  CFG_S & PLAN & CAP & ACT & SUM & TRENDS --> LAYOUT & STYLES
  PLAN -->|"SUM refs"| CAP
  SUM -->|"refs"| PLAN & ACT
  TRENDS -->|"refs"| PLAN & ACT & CAP & CFG_S
```

## Build Order

Sheets are built sequentially because later builders need row positions from earlier ones.

```mermaid
sequenceDiagram
  participant WB as workbook.py
  participant CS as config_sheet
  participant PL as plan
  participant CP as capacity
  participant TR as trends (stub)
  participant AC as actuals
  participant SM as summary
  participant T2 as trends (content)
  participant MT as meta
  WB->>CS: build_config_sheet()
  Note right of CS: Returns config_layout
  WB->>PL: build_plan_sheet(config_layout)
  Note right of PL: Returns plan_layout
  WB->>CP: build_capacity_sheet(config_layout, plan_layout)
  Note right of CP: Returns capacity_layout
  WB->>PL: _apply_planned_formulas(capacity_layout)
  Note right of PL: Patches Plan with SUM(Capacity!...)
  WB->>TR: Skeleton: current_month input + named range
  WB->>AC: build_actuals_sheet(plan_layout)
  Note right of AC: Returns actuals_layout
  WB->>SM: build_summary_sheet(plan_layout, actuals_layout)
  WB->>T2: build_trends_sheet(all layouts)
  WB->>MT: build_meta_sheet(plan_layout)
```

## Init Workflow

```mermaid
sequenceDiagram
  participant PM as Project Manager
  participant CLI as budget_planner CLI
  participant CFG as config.py
  participant WB as workbook.py
  participant XLSX as .xlsx file
  participant LO as LibreOffice

  PM->>CFG: Edit seed data (WPs, team, rate)
  PM->>CLI: python -m budget_planner init [--demo]
  CLI->>CFG: Read seed data
  CLI->>WB: init_workbook()
  WB->>XLSX: Generate all 7 sheets
  Note over XLSX: Capacity/Actuals cells empty
  alt --demo flag
    CLI->>WB: seed_demo_data()
    WB->>XLSX: Write FTE + EUR values
  end
  PM->>LO: python scripts/recalc.py output/...xlsx
  LO->>XLSX: Evaluate all formulas
  PM->>XLSX: Open in Excel
```

## Rebuild Workflow

```mermaid
sequenceDiagram
  participant PM as Project Manager
  participant XL as Excel/LibreOffice
  participant CLI as budget_planner CLI
  participant RDR as reader.py
  participant WB as workbook.py
  participant XLSX as .xlsx file

  PM->>XL: Edit Config (add WP, change team)
  PM->>XL: Save workbook
  PM->>CLI: python -m budget_planner rebuild
  CLI->>RDR: read_config_sheet() + read_capacity_sheet() + read_actuals_values()
  RDR->>XLSX: Extract params, WPs, team, FTE, EUR
  CLI->>WB: validate_config()
  Note over WB: Check for errors/warnings,<br/>orphaned data
  CLI->>WB: _build_all_sheets() with preserved data
  WB->>XLSX: Regenerate all sheets,<br/>restore Capacity FTE + Actuals EUR
  PM->>CLI: python scripts/recalc.py
  PM->>CLI: pytest tests/ -v
```

## Cross-Sheet Formula References

Directed acyclic graph of cross-sheet `=SheetName!CellRef` formulas.

```mermaid
flowchart TD
  CONFIG["Config<br/>WP table, roster,<br/>named ranges"]
  PLAN["Plan<br/>=Config!A/B/C refs<br/>=SUM(Capacity!...) per WP"]
  CAP["Capacity<br/>=Config!A refs (person names)<br/>=Config!B refs (Max FTE)"]
  ACT["Actuals<br/>=Config!A/B/G refs<br/>uses current_month"]
  SUM["Summary<br/>=SUM(Plan!...) per year<br/>=SUM(Actuals!...) per year<br/>uses fte_rate, partner_ceiling"]
  TRENDS["Trends<br/>=Config! refs (WP start/end)<br/>=Actuals! cum/ceiling<br/>=Capacity! TOTAL rows<br/>=Plan! FTE for chart"]

  CONFIG --> PLAN
  CONFIG --> CAP
  CONFIG --> ACT
  CONFIG --> TRENDS
  CAP --> PLAN
  PLAN --> SUM
  ACT --> SUM
  PLAN --> TRENDS
  ACT --> TRENDS
  CAP --> TRENDS
```

## Design Decisions

- **Even FTE spread**: Default allocation divides total WP FTE evenly across active months; users override individual cells in Capacity
- **Single FTE rate**: All EUR amounts are derived from `FTE × monthly rate` (named range `fte_rate`); no per-person salary data
- **Separate Plan / Actuals sheets**: Plan tracks planned FTE (from Capacity SUM); Actuals tracks EUR from accountant. Different units, different sources, never mixed
