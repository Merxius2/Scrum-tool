# Sprint planning — agent context

## Goal

Excel sprint planner: capacity, Committed?=Y backlog, velocity dashboard, Instructions tab last.

## Important behaviour (2026)

- Default **`sprints`: []** — fresh `--force` creates no sprint tabs until `add_sprints` is used.
- **`add_sprints` + run (no `--force`)** appends new sprint sheets before Instructions; does not modify existing sprint tabs.
- **`update_workbook`** with empty `add_sprints` only rebuilds Velocity links from existing sprint tab names.
- **Velocity**: one row per sprint tab; columns link via `sprint_meta_from_sheet_name` / `build_sprint_sheet` return dict.
- **Sheet protection** on all tabs; sprint tab unlocks input cells only.

## Key files

| File | Role |
| --- | --- |
| `generate_sprint_planner.py` | Generator + incremental update |
| `sprint_planner.config.json` | Team, `sprints`, `add_sprints` |
| `README.md` | Operator docs |

Keep this file updated when CLI, config keys, or layout changes.
