# Scrum-tool

Tools that help teams working in a Scrum configuration.

Open this folder (`…/Scrum-tool/scrum-tool`) as your Cursor/IDE workspace — it is the git repository root.

For coding agents: start with [`repo-map.json`](./repo-map.json) and [`AGENTS.md`](./AGENTS.md).

## Tools

| Tool | Description |
| --- | --- |
| [holiday-planner](./holiday-planner/) | Generate an Excel team holiday planner (weeks, capacity row, colour-coded engineers). |
| [sprint-planning](./sprint-planning/) | Generate an Excel sprint planner (capacity, backlog, velocity, one tab per sprint). |

## Tests

```bash
pip install -r requirements-dev.txt
python3 -m pytest tests/ -v
python3 -m pytest tests/ -q --cov=generate_holiday_planner --cov=generate_sprint_planner --cov-report=term-missing
```
