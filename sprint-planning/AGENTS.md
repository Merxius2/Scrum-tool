# Sprint planning workbook — agent context

Read this file before changing code, config, or docs in this folder.

## Goal

Help a Scrum team plan sprints in one shared Excel workbook: **capacity**, **committed backlog**, **results**, and **velocity** across multiple sprints.

The tool should:

- Create **one tab per sprint** (from JSON config) plus **Product backlog** and **Velocity** summary tabs.
- On each sprint tab: sprint metadata (dates, goal), **capacity calculator** per team member (PTO, availability %, focus % → story points), and a **sprint backlog** table with dropdowns.
- Compute **sprint results** with Excel formulas: committed SP, completed SP, velocity, carry-over, completion %, load vs team capacity.
- **Velocity** tab links to each sprint’s result cells; includes rolling average velocity and a suggested next-sprint capacity hint.
- **Product backlog** tab for ranked items not yet assigned to a sprint (target sprint dropdown points at configured sprint sheet names).
- Use **formulas** (NETWORKDAYS, SUM, SUMIF-style Done SP column) so the file stays live after manual edits.
- Keep dependencies minimal (**openpyxl** only).

Generated output is **`Sprint-Planning.xlsx`** (gitignored; recreate with the script). User-facing usage lives in **`README.md`**.

## What is already done

- **`generate_sprint_planner.py`** — Python 3 generator using **openpyxl**.
- **`sprint_planner.config.json`** — team members, sprint list (sheet name, start date, goal), global **`sprint_length_weeks`** or **`sprint_length_working_days`**, story-point scale, statuses/types, capacity defaults, row counts.
- **Sprint end date**: `End (auto)` cell uses `=WORKDAY(Start, length−1)` so editing Start in Excel recalculates end, `NETWORKDAYS`, and capacity. Per-sprint length override via `working_days` / `sprint_length_working_days` on a sprint entry.
- **Sprint sheet layout**: fixed sections so summary cell addresses are stable for cross-sheet references on **Velocity**.
- **Capacity**: `NETWORKDAYS` between Start and End (auto); per-member capacity = `(sprint working days − PTO) × availability × focus × story_points_per_day`.
- **Backlog**: ID, title, type, SP, assignee, status, notes; **Done SP** column formula when status is Done.
- **Validation lists** for SP scale, type, status, assignee; conditional formatting on load vs capacity and velocity column.
- **`requirements.txt`** — `openpyxl>=3.1.0`.
- **`README.md`** — setup, config reference, CLI (`-c` for alternate config path).
- Repo home: **`sprint-planning/`** inside [Scrum-tool](https://github.com/Merxius2/Scrum-tool.git); root README links here.

## Suggested extensions (not implemented unless requested)

- Burndown chart data column (remaining SP per day).
- Automatic import of PTO from the **holiday-planner** workbook.
- Definition-of-ready checklist column on product backlog.
- Blocked-item ageing / WIP limits per assignee.
- Sprint retrospective notes section per tab.

## Key files (do not break without reason)

| File | Role |
| --- | --- |
| `generate_sprint_planner.py` | Workbook structure, formulas, and layout constants |
| `sprint_planner.config.json` | Default team and sprints; safe to edit for local team |
| `README.md` | Human documentation; keep in sync with behaviour |
| `AGENTS.md` | This file — project memory for humans and coding agents |

## Conventions for future changes

- Prefer **config over hardcoding** (team, sprints, scales, defaults, output filename).
- If you change sprint sheet **row layout**, update **Velocity** link generation and document the new anchor rows here.
- Preserve **Excel formulas** in summary and Done SP columns.
- Sheet tab order: **Product backlog**, **Velocity**, then sprints in config order.
- Minimize dependencies; stay on **openpyxl** unless there is a strong reason to add more.

## Keeping this file up to date

**Who:** Anyone (human or agent) who merges a meaningful change to this tool.

**When to update `AGENTS.md`:**

- New feature or changed behaviour (layout, config keys, CLI, output format, formula semantics).
- New or removed files that agents should know about.
- Shift in product goal or target users.
- Known limitations, tradeoffs, or intentional non-goals worth remembering.

**When you can skip:** Typos only, dependency pin bumps with no behaviour change, or regenerating the xlsx locally.

**How to update:**

1. Edit the **Goal** section if intent changed.
2. Append or revise bullets under **What is already done** (dated bullet optional, e.g. `2026-10-17: …`).
3. Update **Key files** if layout changes.
4. Update **Conventions** if new rules apply.
5. Move items from **Suggested extensions** into **What is already done** when implemented.
6. Mirror user-visible details in **`README.md`**; `AGENTS.md` is for context and decisions, README for operators.

**Do not** duplicate the full README here — summarize decisions and history so the next session can continue without re-reading the whole codebase.

## Out of scope (unless asked)

- Web UI, Jira/Azure DevOps sync, or multi-team databases.
- Committing generated `.xlsx` files (remain gitignored).
- Macro-enabled `.xlsm` or chart objects (data-only sheets for now).
