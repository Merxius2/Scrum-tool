# Sprint planning — agent context

## Config

- **`sprints`**: integer count of tabs on fresh `--force` generate (`Sprint 1` … `Sprint N`), not a list of objects.
- **`add_sprints`**: removed — use CLI **`addsprint COUNT [workbook.xlsx]`**.

## CLI

- Default: create workbook; error if output exists unless `--force`.
- **`addsprint`**: append COUNT new sprint sheets; renumber after highest existing `{prefix} N`; rebuild Velocity.

## Key files

| File | Role |
| --- | --- |
| `generate_sprint_planner.py` | Generator |
| `sprint_planner.config.json` | Team + `sprints` count |

Update this file when config keys or CLI change.
