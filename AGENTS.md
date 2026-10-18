# Scrum-tool — agent context

## Start here

1. **Load [`repo-map.json`](./repo-map.json)** — paths, CLI, config keys, tests.
2. **Read this file** — product goals, conventions, and what to keep updated.
3. When editing a tool, also read that tool’s **`README.md`** (`holiday-planner/README.md` or `sprint-planning/README.md`).

Human overview: **[README.md](./README.md)**.

## Repository layout

| Path | Role |
| --- | --- |
| `repo-map.json` | Machine-readable map (keep current) |
| `AGENTS.md` | This file — all agent context for the repo |
| `holiday-planner/` | Holiday Excel generator |
| `sprint-planning/` | Sprint planning Excel generator |
| `tests/` | Pytest suite (`tests/conftest.py` imports generators) |
| `requirements-dev.txt` | pytest + pytest-cov + planner deps |

## Unit tests

From repo root:

```bash
pip install -r requirements-dev.txt
python3 -m pytest tests/ -v
python3 -m pytest tests/ -q --cov=generate_holiday_planner --cov=generate_sprint_planner --cov-report=term-missing
```

Modules load as `generate_holiday_planner` / `generate_sprint_planner` in `tests/conftest.py` — use those names with `--cov`, not file paths.

| Tool | Test file | Run / update when |
| --- | --- | --- |
| Holiday | `tests/test_holiday_planner.py` | Changes to `generate_holiday_planner.py`, config, CLI, backups, roster commands, layout / CF / protection |
| Sprint | `tests/test_sprint_planner.py` | Changes to `generate_sprint_planner.py`, config, CLI, backups, layout, velocity/backlog formulas, protection |

Prefer pure helpers + `tmp_path` integration tests; do not depend on gitignored `.xlsx` files.

## Keeping `AGENTS.md` and `repo-map.json` current

**Who:** Anyone who merges a meaningful change.

**Update [`repo-map.json`](./repo-map.json) when:** new/removed tools or paths; generator/config/output/backup paths; CLI commands or flags; config keys agents need; test paths or import module names; gitignore rules for artifacts.

**Update this file (`AGENTS.md`) when:** product goal, behaviour, CLI, config semantics, conventions, test workflow, or out-of-scope boundaries change for either tool. Edit the matching section below and mirror user-facing details in the tool **`README.md`**.

**When you can skip:** Typos only, dependency pin bumps with no behaviour change, or local xlsx regeneration.

**Repo map mechanics:** Edit the relevant `tools[]` entry; bump `schemaVersion` only on breaking top-level shape changes. Keep `repo-map.json` factual and minimal; keep narrative and decisions here.

## Repo-wide conventions

- Do not commit generated `.xlsx` or `backups/` (see `.gitignore`, `repo-map.json` → `gitignoreOutputs`).
- Prefer config over hardcoding team size, sheet counts, output names.
- Only create git commits when the user asks.

---

## Holiday planner

**Folder:** `holiday-planner/` · **Operator docs:** `holiday-planner/README.md`

### Goal

Help a Scrum team plan time off in one shared Excel workbook so capacity is visible at a glance.

- One engineer per row; weekdays as columns, grouped by **ISO week**.
- Multiple sheets from JSON config (e.g. remainder of year + full calendar years).
- Day off marker in cells (default **`H`**).
- Bottom row: daily count on holiday; CF **green** (0) → **yellow** (1–2) → **red** (3+).
- Distinct engineer colours on generate (column A + **`H`** cells via **CellIsRule**).
- Week boundaries: stronger border on first weekday of each ISO week.
- Column **B**: per-engineer days off (COUNTIF), colour scale min→max within team.

Output: **`Team-Holiday-Planner.xlsx`** (gitignored).

### Implemented behaviour

- **`generate_holiday_planner.py`** — openpyxl generator.
- **`holiday_planner.config.json`** — `team_members`, `years`, `output_file`, `holiday_marker`, optional **`color_seed`**.
- Periods: full `year`, or `"from": "remainder"` through a fixed end date.
- Header dates **`dd-mm-yyyy`**; sheet protection (lock formulas/headers; unlock names and day cells).
- **CLI:** `generate` (default), `--force`, **`add-year`**, **`add-engineer` / `remove-engineer`** (updates config + all sheets). **Backups:** `holiday-planner/backups/{name}.backup-YYYYMMDD-HHMMSS.xlsx`.

### Key files

| File | Role |
| --- | --- |
| `generate_holiday_planner.py` | Workbook structure and behaviour |
| `holiday_planner.config.json` | Default team and periods |
| `README.md` | Human documentation |

### Conventions

- Weekdays only (Mon–Fri) unless asked otherwise.
- Preserve formulas in total row after manual edits.
- Engineer colours: one palette per generate run, shared across sheets in that workbook.
- Stay on **openpyxl** unless there is strong reason to add deps.

### Out of scope (unless asked)

- Web UI, auth, multi-team databases, HR/calendar sync.

---

## Sprint planning

**Folder:** `sprint-planning/` · **Operator docs:** `sprint-planning/README.md`

### Goal

Excel workbook for sprint capacity, per-sprint backlog, velocity history, and product backlog.

### Config

- **`sprints`**: integer tab count on fresh `--force` (`Sprint 1` … `Sprint N`), not a list of objects.
- **`sprint_length`**: `{"weeks": N}` or `{"working_days": N}` — end date, capacity, meta row.
- **`add_sprints`**: removed — use **`add-sprint`**.

### CLI

- Default `generate`; fails if output exists unless **`--force`**.
- **`add-sprint COUNT [workbook.xlsx]`**: append sprint sheets; renumber after highest `{prefix} N`; rebuild Velocity.
- **`add-engineer` / `remove-engineer NAME`**: update `team_members`; rebuild capacity on each sprint tab (keep backlog, dates, PTO/FTE by name); rebuild Velocity.
- **Backups:** `sprint-planning/backups/{name}.backup-YYYYMMDD-HHMMSS.xlsx` before `--force`, `add-sprint`, or team commands.

System sheets: **Product backlog**, **Velocity**, **Instructions** (regenerate with `--force` to refresh Instructions text).

### Key files

| File | Role |
| --- | --- |
| `generate_sprint_planner.py` | Generator |
| `sprint_planner.config.json` | Team, `sprints`, sprint length, scales |
