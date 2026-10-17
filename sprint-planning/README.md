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

Layout follows common sprint-planning templates: **overview KPIs** (top-right), **capacity & load**, **backlog with commit flag**, then **results**.

1. Set **Start** and **Sprint goal**. **End (auto)** follows configured sprint length (`WORKDAY`).
2. **Sprint overview** (columns F–G): enter **Planning velocity**; see **Previous sprint velocity** (from the prior tab), **Suggested SP limit**, **Committed SP**, **Buffer**, and **Commitment status** (`OVER COMMITTED` when above limit).
3. **Team capacity & load**: **PTO** and **FTE %** per person; **Assigned SP** rolls up backlog rows where **Committed? = Y**; **Headroom** turns red if someone is overloaded.
4. **Sprint backlog**: add items; set **Committed?** to **Y** only for work the team agrees to take (candidates can stay blank/N). Use **Priority**, **Dependency / blocker**, **Assignee**, **Status**.
5. After the sprint, set **Status** to **Done** — **Sprint results** and the **Velocity** tab update from completed points.

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
| `default_fte_percent` | Initial FTE % column per member (default: 100). Legacy key `default_availability_percent` still accepted. |
| `story_points_per_day` | SP per full-time working day at 100% FTE (default: 1). |
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
