# Sprint planning workbook

Generate an Excel workbook for Scrum sprint planning: capacity per engineer, sprint backlog, results, and a velocity dashboard across multiple sprints.

## Quick start

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Edit `sprint_planner.config.json` (team, sprints, defaults).

3. Generate the workbook:

   ```bash
   python3 generate_sprint_planner.py
   ```

4. Open `Sprint-Planning.xlsx` (or the filename set in config).

## Workbook tabs

| Tab | Purpose |
| --- | --- |
| **Product backlog** | Prioritised items not yet in a sprint; optional target sprint and “Ready?” flag. |
| **Velocity** | Cross-sprint committed/completed SP, carry-over, completion %, rolling average velocity. |
| **One tab per sprint** | Capacity, sprint backlog, and sprint results for that iteration. |

## Using a sprint tab

1. Set **Start** (and **Sprint goal**). **End (auto)** is calculated from the configured sprint length (`WORKDAY`); change Start in Excel and End, working days, and capacity update automatically.
2. **Capacity**: enter **PTO days** and adjust **Availability %** / **Focus %** per person. **Team capacity (SP)** sums individual capacity using working days (`NETWORKDAYS`, Mon–Fri only).
3. **Sprint backlog**: add work items; pick **Story points**, **Type**, **Assignee**, and **Status** from dropdowns.
4. Mark finished items **Done** — the **Done SP** column and **Sprint results** block update automatically.
5. Enter **Planning velocity (SP)** (expected throughput from past sprints). **Suggested story point limit** = `MIN(team capacity, planning velocity)` — leave velocity blank to use capacity only.
6. Before sprint start, watch **Committed vs suggested limit** (and **Load vs capacity**). Green/yellow/red formatting highlights over-commitment.

## Configuration

File: `sprint_planner.config.json`

| Field | Description |
| --- | --- |
| `team_members` | Names used in capacity rows and assignee dropdowns. |
| `sprints` | List of sprint objects (see below). |
| `output_file` | Output `.xlsx` filename (default: `Sprint-Planning.xlsx`). |
| `story_point_scale` | Allowed values in SP dropdowns (default: 1, 2, 3, 5, 8, 13). |
| `work_item_types` | Backlog type dropdown (default: Story, Bug, Task, Spike). |
| `statuses` | Status dropdown (default: To Do, In Progress, Done, Blocked). |
| `default_availability_percent` | Initial availability column (default: 100). |
| `default_focus_percent` | Initial focus column (default: 80). |
| `story_points_per_day` | SP multiplier at 100% avail/focus per working day (default: 1). |
| `sprint_length_weeks` | Sprint span as weeks × 5 working days (default: 2 → 10 days). Ignored if `sprint_length_working_days` is set. |
| `sprint_length_working_days` | Exact Mon–Fri days per sprint, inclusive of start and end (overrides `sprint_length_weeks`). |
| `default_planning_velocity` | Optional starting value for **Planning velocity** on the first sprint tab. |
| `prefill_velocity_from_rolling_average` | When true (default), sprint 2+ prefills planning velocity from the Velocity tab rolling average (editable in Excel). |
| `max_backlog_rows_per_sprint` | Fixed backlog size per sprint tab (default: 30). |
| `product_backlog_rows` | Rows on Product backlog tab (default: 40). |

### Sprint object

```json
{
  "sheet_name": "Sprint 1",
  "title": "Sprint 1 — Kickoff",
  "start": "2026-10-20",
  "goal": "Deliver first increment"
}
```

- `sheet_name` — Excel tab name (max 31 characters).
- `title` — Display name on the sprint tab (optional; defaults to sheet name).
- `start` — ISO date (`YYYY-MM-DD`) for the first day of the sprint; end date is computed in the sheet.
- `goal` — Free-text sprint goal.
- `sprint_length_working_days` or `working_days` (optional) — override global sprint length for this tab only.
- `planning_velocity` (optional) — fixed planning velocity for this sprint instead of default / rolling average.

## Command-line options

```bash
python3 generate_sprint_planner.py
python3 generate_sprint_planner.py -c /path/to/other-config.json
```

The default config path is `sprint_planner.config.json` in the same folder as the script.

## Project layout

```
sprint-planning/
├── AGENTS.md                      # agent context (Cursor auto-loads in this folder)
├── README.md
├── requirements.txt
├── sprint_planner.config.json
├── generate_sprint_planner.py
└── Sprint-Planning.xlsx           # created when you run the script
```

## Requirements

- Python 3.9+
- [openpyxl](https://openpyxl.readthedocs.io/)

## Related tools

- [holiday-planner](../holiday-planner/) — team time-off grid; PTO on sprint tabs can mirror days planned there manually.

## Ideas for future versions

- Burndown remaining-SP column per sprint day.
- Import PTO from the holiday planner workbook.
- Definition-of-ready checklist on product backlog items.
- WIP limit warnings per assignee.
