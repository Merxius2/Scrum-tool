# Holiday planner — agent context

Read this file before changing code, config, or docs in this folder.

## Goal

Help a Scrum development team plan time off in one shared Excel workbook so capacity is visible at a glance.

The tool should:

- Put **one engineer per row** and **weekdays as columns**, grouped by **ISO week**.
- Support **multiple sheets** (e.g. remainder of current year + full calendar years) from JSON config.
- Let planners mark days off with a simple cell marker (default **`H`**).
- Show a **daily total** of engineers on holiday on the bottom row.
- Highlight risk on the total row: **green** when nobody is away, **yellow** for 1–2, **red** when more than two are away.
- Assign each engineer a **distinct colour** on generate; use it in column A and for **`H`** cells via conditional formatting.
- Make **week boundaries** easy to see (stronger border on the first weekday of each ISO week).
- Column **B**: per-engineer **total days off** on the sheet (COUNTIF row), colour scale red (min) → green (max) within the team.

Generated output is **`Team-Holiday-Planner.xlsx`** (gitignored; recreate with the script). User-facing usage lives in **`README.md`**.

## What is already done

- **`generate_holiday_planner.py`** — Python 3 generator using **openpyxl**.
- **`holiday_planner.config.json`** — team member names, `years` / periods, `output_file`, `holiday_marker`, optional **`color_seed`**.
- **Config-driven periods**: full `year`, or custom range with `"from": "remainder"` (first weekday on/after run date) through a fixed end date.
- **Sheet layout**: merged week header row, day row (`Mon dd-mm`), frozen panes from first data column, COUNTIF totals, data validation hint for marker.
- **Styling**: random pastel engineer colours (or seeded), FormulaRule fill when cell equals marker, ColorScaleRule on total row, medium left border on week-start columns.
- **`requirements.txt`** — `openpyxl>=3.1.0`.
- **`README.md`** — setup, config reference, CLI (`-c` for alternate config path).
- **CLI**: `generate` (default), `--force`, **`addyear YEAR …`** (append calendar year sheets; backup before overwrite). Header dates **`dd-mm-yyyy`**. **Sheet protection**: lock formula/header cells; unlock names and day markers.
- Repo home: **`Documents/Git/Scrum-tool/holiday-planner/`** inside [Scrum-tool](https://github.com/Merxius2/Scrum-tool.git); root README links here.

## Key files (do not break without reason)

| File | Role |
| --- | --- |
| `generate_holiday_planner.py` | Single source of truth for workbook structure and behaviour |
| `holiday_planner.config.json` | Default team and periods; safe to edit for local team |
| `README.md` | Human documentation; keep in sync with behaviour |
| `AGENTS.md` | This file — project memory for humans and coding agents |

## Conventions for future changes

- Prefer **config over hardcoding** (team size, years, marker, output name).
- Keep **weekdays only** (Mon–Fri) unless explicitly requested otherwise.
- Preserve **Excel formulas** in the total row so the sheet stays useful after manual edits.
- Engineer colours: one palette per **`generate()`** run, shared across all sheets in that workbook.
- Minimize dependencies; stay on **openpyxl** unless there is a strong reason to add more.

## Keeping this file up to date

**Who:** Anyone (human or agent) who merges a meaningful change to this tool.

**When to update `AGENTS.md`:**

- New feature or changed behaviour (layout, colours, config keys, CLI, output format).
- New or removed files that agents should know about.
- Shift in product goal or target users.
- Known limitations, tradeoffs, or intentional non-goals worth remembering.

**When you can skip:** Typos only, dependency pin bumps with no behaviour change, or regenerating the xlsx locally.

**How to update:**

1. Edit the **Goal** section if intent changed.
2. Append or revise bullets under **What is already done** (dated bullet optional, e.g. `2026-10-17: …`).
3. Update **Key files** if layout changes.
4. Update **Conventions** if new rules apply.
5. Mirror user-visible details in **`README.md`**; `AGENTS.md` is for context and decisions, README for operators.

**Do not** duplicate the full README here — summarize decisions and history so the next session can continue without re-reading the whole codebase.

## Out of scope (unless asked)

- Web UI, authentication, or multi-team databases.
- Automatic sync from HR/calendar systems.
- Committing generated `.xlsx` files (remain gitignored).
