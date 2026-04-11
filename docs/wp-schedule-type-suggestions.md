# WP Schedule Type Suggestions

Per-WP distribution types for a two-range split where Range 2 is 4x smaller than Range 1.

## Recommendations

| WP  | Name                           | Range 1 Type   | Range 2 Type   | Reasoning                                                  |
|-----|--------------------------------|----------------|----------------|------------------------------------------------------------|
| WP1 | Project Management             | flat           | flat           | Management is steady in both phases                        |
| WP2 | Requirements and concept    | front-loaded   | flat           | Define architecture fast upfront; tail is minor review     |
| WP3 | Research and analysis | front-loaded | flat           | Heavy research early; tail is follow-up                    |
| WP4 | Architecture design            | bell           | flat           | Design ramps, peaks, winds down; tail is updates           |
| WP5 | Component design              | bell           | flat           | Design iteration peaks mid-range; tail is updates          |
| WP6 | Subsystem design | bell           | flat           | Same pattern as WP5                                        |
| WP7 | Software development        | bell           | back-loaded    | Dev peaks mid-range; tail is integration/bugfix push       |
| WP8 | Prototyping       | back-loaded    | back-loaded    | Assembly builds toward deadline in both phases             |
| WP9 | Validation and demo      | back-loaded    | back-loaded    | Final demo push — intensity increases toward end           |

## General Principles

- **Range 1** keeps the original distribution type — it carries 80% of the effort and reflects the WP's natural work rhythm.
- **Range 2** defaults to **flat** for most WPs — the 20% tail is too small for curve shaping to matter meaningfully.
- **Exception:** Late-stage WPs (WP7-9) use **back-loaded** for Range 2 because the second range is still building toward a project deadline.
