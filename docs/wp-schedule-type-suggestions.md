# WP Type Suggestions

Each WP has a single contiguous active range `[Start, End]` and a single distribution `Type` from `{flat, front-loaded, bell, back-loaded}`. The Type drives the per-month curve shape `capacity.py::distribute` applies to each assignment's `Total FTE` scalar.

## Recommendations

| WP  | Name                             | Type           | Reasoning                                             |
|-----|----------------------------------|----------------|-------------------------------------------------------|
| WP1 | Project Management               | flat           | Steady management effort across the whole range       |
| WP2 | Requirements and concept      | front-loaded   | Define architecture fast upfront; tail is minor       |
| WP3 | Research and analysis | front-loaded   | Heavy research early; tail is follow-up               |
| WP4 | Architecture design              | bell           | Design ramps, peaks mid-range, winds down             |
| WP5 | Component design                | bell           | Iteration peaks mid-range                             |
| WP6 | Subsystem design   | bell           | Same pattern as WP5                                   |
| WP7 | Software development          | back-loaded    | Dev builds toward integration/bugfix push near end    |
| WP8 | Prototyping         | back-loaded    | Assembly ramps up toward the deadline                 |
| WP9 | Validation and demo        | back-loaded    | Final demo push — intensity increases toward end      |

## General Principles

- **flat** — steady-state work (management, ongoing analysis).
- **front-loaded** — work that needs a fast start (early design, research).
- **bell** — work that ramps, peaks, and tapers (iterative design).
- **back-loaded** — deadline-driven work (integration, demo prep).
