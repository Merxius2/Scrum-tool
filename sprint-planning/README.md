# Sprint planning workbook

Generate an Excel workbook for Scrum sprint planning: capacity, committed backlog, velocity, and instructions.

## Quick start

```bash
pip install -r requirements.txt
python3 generate_sprint_planner.py --force   # first time: empty workbook (no sprint tabs)
```

Open **Instructions** (last tab) for a fill-in guide.

## Add sprint tabs without losing data

1. Edit `sprint_planner.config.json` — add entries to **`add_sprints`** (not `sprints`):

```json
"add_sprints": [
  { "sheet_name": "Sprint 1", "title": "Sprint 1", "goal": "" }
]
```

2. Run (no `--force`):

```bash
python3 generate_sprint_planner.py
```

Existing sprint sheets and backlog data are kept. The **Velocity** tab is rebuilt so **one row per sprint tab** links to committed/completed SP on that sprint (formulas — do not type in Velocity).

3. Clear `add_sprints` after adding (optional) to avoid skipping duplicates on the next run.

## Commands

| Command | Behaviour |
| --- | --- |
| `python3 generate_sprint_planner.py` | If `.xlsx` exists: apply `add_sprints` and refresh Velocity; if `add_sprints` is empty, only refresh Velocity links. |
| `python3 generate_sprint_planner.py --force` | Recreate entire workbook from config (`sprints` list only). **Overwrites** the file. |

## Config highlights

| Field | Purpose |
| --- | --- |
| `sprints` | Used only with **`--force`** (initial sprint tabs). Default sample config uses `[]` (empty). |
| `add_sprints` | Sprint tabs to append when the file already exists. |
| `sheet_name` | Required. Excel tab name (max 31 chars). |
| `start`, `title`, `goal` | Optional; leave empty and fill in Excel. |

## Sheet protection

Formula and label cells are **locked**; editable fields (start date, goal, PTO, FTE, planning velocity, backlog columns except Done SP) are unlocked. Velocity and Instructions tabs are read-only (linked formulas).

## Workbook tabs

**Product backlog** → **Velocity** → **Sprint …** → **Instructions**
