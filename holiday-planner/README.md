# Team holiday planner

Generate an Excel workbook for planning team time off. Each team member gets a row; weekdays are columns grouped by ISO week. A bottom row counts how many people are away each day, with a color scale (green → yellow → red).

## Quick start

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Edit `holiday_planner.config.json` (names and which periods to generate).

3. Generate the workbook:

   ```bash
   python3 generate_holiday_planner.py
   ```

4. Open `Team-Holiday-Planner.xlsx` (or the filename set in config).

## Using the spreadsheet

- Enter **`H`** (or your configured marker) in a cell when someone is on holiday; leave the cell blank when they are working.
- Each engineer gets a **random colour** when the workbook is generated. Column A uses that colour as the row label background; holiday cells (`H`) use the same colour.
- **Week boundaries** use a stronger vertical line on the first weekday column of each ISO week.
- The **Total on holiday** row uses formulas to count markers per day.
- Conditional formatting on that row: **0** = green, **1–2** = yellow, **3+** = red (more than two engineers away).

Column A names can be edited in Excel after generation; re-run the script if you change names in config and want a fresh file (and new colours, unless you set `color_seed`).

## Configuration

File: `holiday_planner.config.json`

| Field | Description |
| --- | --- |
| `team_members` | List of names (one row per person). |
| `years` | List of sheets to create (see below). |
| `output_file` | Output `.xlsx` filename (default: `Team-Holiday-Planner.xlsx`). |
| `holiday_marker` | Character(s) marked in cells for a day off (default: `H`). |
| `color_seed` | Optional integer for reproducible engineer colours (omit for random each run). |

### Period definitions

Each item in `years` is either a **number** (shorthand for a full calendar year) or an **object**.

**Full year** (weekdays only, January–December):

```json
{ "sheet_name": "2027", "year": 2027 }
```

Shorthand:

```json
2028
```

**Custom date range:**

```json
{
  "sheet_name": "Rest of 2026",
  "from": "remainder",
  "through": "2026-12-31"
}
```

- `"from": "remainder"` — first weekday on or after **today** (when you run the script).
- Or use fixed dates: `"from": "2026-11-01"`, `"through": "2026-12-31"`.

Alternative keys `start` / `end` / `to` are accepted instead of `from` / `through`.

### Example config

```json
{
  "output_file": "Team-Holiday-Planner.xlsx",
  "holiday_marker": "H",
  "team_members": [
    "Alice",
    "Bob",
    "Charlie",
    "Diana",
    "Eve"
  ],
  "years": [
    {
      "sheet_name": "Rest of 2026",
      "from": "remainder",
      "through": "2026-12-31"
    },
    { "sheet_name": "2027", "year": 2027 }
  ]
}
```

## Command-line options

```bash
python3 generate_holiday_planner.py
python3 generate_holiday_planner.py -c /path/to/other-config.json
```

The default config path is `holiday_planner.config.json` in the same folder as the script.

## Project layout

```
team-holiday-planner/
├── README.md
├── requirements.txt
├── holiday_planner.config.json
├── generate_holiday_planner.py
└── Team-Holiday-Planner.xlsx   # created/updated when you run the script
```

## Requirements

- Python 3.9+
- [openpyxl](https://openpyxl.readthedocs.io/)
