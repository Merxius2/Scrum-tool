# Sprint planning workbook

Generate an Excel workbook for Scrum sprint planning.

## Quick start

```bash
pip install -r requirements.txt
```

Set **`"sprints": 5`** in `sprint_planner.config.json` to create five empty tabs (`Sprint 1` … `Sprint 5`) on first generate.

```bash
python3 generate_sprint_planner.py --force
```

## Commands

| Command | Description |
| --- | --- |
| `python3 generate_sprint_planner.py` | Create `output_file` from config. Fails if the file already exists. |
| `python3 generate_sprint_planner.py --force` | Recreate the workbook (uses **`sprints`** count from config). |
| `python3 generate_sprint_planner.py addsprint 4` | Append 4 sprint tabs to the config `output_file`. |
| `python3 generate_sprint_planner.py addsprint 4 Sprint-Planning.xlsx` | Append 4 tabs to a specific workbook. |

New sprint names continue the sequence (`Sprint 3`, `Sprint 4`, …) based on existing tab names. The **Velocity** sheet is rebuilt with one linked row per sprint tab.

Before **`addsprint`** or **`generate --force`** overwrites an existing file, the script saves a timestamped backup next to it (e.g. `Sprint-Planning.backup-20261018-120000.xlsx`).

## Config

| Field | Description |
| --- | --- |
| `sprints` | **Integer** — number of sprint tabs to create on generate (0 = no sprint tabs). |
| `sprint_name_prefix` | Tab names `{prefix} 1`, `{prefix} 2`, … (default: `Sprint`). |
| `output_file` | Output `.xlsx` filename. |
| `sprint_length` | Sprint span: `{"weeks": 2}` or `{"working_days": 10}`, or an integer (working days). Legacy: `sprint_length_weeks` / `sprint_length_working_days`. Drives end date, capacity, and **Sprint length** on each sprint tab. |

See **Instructions** tab in the workbook for fill-in guidance.
