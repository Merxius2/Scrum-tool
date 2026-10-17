#!/usr/bin/env python3
"""Generate Scrum sprint planning Excel workbook from JSON config."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DEFAULT_CONFIG_NAME = "sprint_planner.config.json"

THIN = Side(style="thin", color="CCCCCC")
HEADER_FILL = PatternFill("solid", fgColor="E8EEF4")
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
SUMMARY_FILL = PatternFill("solid", fgColor="FFF2CC")
DONE_FILL = PatternFill("solid", fgColor="C6EFCE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")

# Sprint sheet layout (1-based rows)
META_TITLE_ROW = 1
META_START_ROW = 2
META_END_ROW = 3
META_GOAL_ROW = 4
META_WORKING_DAYS_ROW = 5

# Sprint overview KPIs (labels / values) — columns F–G, clear of backlog table width
KPI_LABEL_COL = 6
KPI_VALUE_COL = 7
KPI_PLANNING_VEL_ROW = 2
KPI_SUGGESTED_LIMIT_ROW = 3
KPI_PREV_VELOCITY_ROW = 4
KPI_COMMITTED_ROW = 5
KPI_BUFFER_ROW = 6
KPI_STATUS_ROW = 7

CAPACITY_TITLE_ROW = 8
CAPACITY_HEADER_ROW = 9
CAPACITY_FIRST_DATA_ROW = 10

COL_MEMBER = 1
COL_PTO = 2
COL_FTE = 3
COL_CAPACITY = 4
COL_ASSIGNED = 5
COL_HEADROOM = 6

BACKLOG_ID = 1
BACKLOG_TITLE = 2
BACKLOG_PRIORITY = 3
BACKLOG_TYPE = 4
BACKLOG_SP = 5
BACKLOG_ASSIGNEE = 6
BACKLOG_COMMITTED = 7
BACKLOG_STATUS = 8
BACKLOG_DEPENDENCY = 9
BACKLOG_NOTES = 10
BACKLOG_DONE_SP = 11


def parse_date(value: str):
    return datetime.strptime(value, "%Y-%m-%d").date()


def resolve_sprint_length_working_days(
    sprint: dict[str, Any] | None, cfg: dict[str, Any]
) -> int:
    """Working days in a sprint (Mon–Fri), inclusive of start and end."""
    if sprint is not None:
        for key in ("sprint_length_working_days", "working_days"):
            if key in sprint:
                days = int(sprint[key])
                if days < 1:
                    raise ValueError(f"Sprint {sprint.get('sheet_name', '?')}: {key} must be >= 1.")
                return days

    if cfg.get("sprint_length_working_days") is not None:
        days = int(cfg["sprint_length_working_days"])
    elif cfg.get("sprint_length_weeks") is not None:
        days = int(cfg["sprint_length_weeks"]) * 5
    else:
        days = 10

    if days < 1:
        raise ValueError("sprint_length_working_days / sprint_length_weeks must yield at least 1 day.")
    return days


def sprint_end_formula(working_days: int) -> str:
    """Excel end date from start in B{META_START_ROW}, inclusive working-day span."""
    if working_days <= 1:
        return f"=B{META_START_ROW}"
    return f"=WORKDAY(B{META_START_ROW},{working_days - 1})"


def load_config(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        data = json.load(f)

    members = data.get("team_members")
    if not members or not isinstance(members, list):
        raise ValueError('"team_members" must be a non-empty list of names.')

    sprints = data.get("sprints")
    if not sprints or not isinstance(sprints, list):
        raise ValueError('"sprints" must be a non-empty list of sprint definitions.')

    return {
        "output_file": data.get("output_file", "Sprint-Planning.xlsx"),
        "team_members": [str(m) for m in members],
        "story_point_scale": [int(x) for x in data.get("story_point_scale", [1, 2, 3, 5, 8, 13])],
        "work_item_types": [str(t) for t in data.get("work_item_types", ["Story", "Bug", "Task", "Spike"])],
        "statuses": [str(s) for s in data.get("statuses", ["To Do", "In Progress", "Done", "Blocked"])],
        "priorities": [str(p) for p in data.get("priorities", ["1 - Must", "2 - Should", "3 - Could"])],
        "default_fte_percent": float(
            data.get("default_fte_percent", data.get("default_availability_percent", 100))
        ),
        "story_points_per_day": float(data.get("story_points_per_day", 1)),
        "max_backlog_rows_per_sprint": int(data.get("max_backlog_rows_per_sprint", 30)),
        "product_backlog_rows": int(data.get("product_backlog_rows", 40)),
        "sprint_length_working_days": data.get("sprint_length_working_days"),
        "sprint_length_weeks": data.get("sprint_length_weeks"),
        "default_planning_velocity": data.get("default_planning_velocity"),
        "prefill_velocity_from_rolling_average": bool(
            data.get("prefill_velocity_from_rolling_average", True)
        ),
        "sprints": sprints,
    }


def list_formula(values: list[str]) -> str:
    escaped = [v.replace('"', '""') for v in values]
    return '"' + ",".join(escaped) + '"'


def apply_table_borders(ws, min_row: int, max_row: int, min_col: int, max_col: int):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).border = Border(
                left=THIN, right=THIN, top=THIN, bottom=THIN
            )


def style_header_row(ws, row: int, min_col: int, max_col: int):
    for c in range(min_col, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def velocity_rolling_average_row(sprint_count: int) -> int:
    """Row on Velocity tab with rolling average formula (must match build_velocity_sheet)."""
    header_row = 4
    data_start = header_row + 1
    last_data = data_start + sprint_count - 1
    return last_data + 2


def initial_planning_velocity(
    sprint_index: int, sprint: dict[str, Any], cfg: dict[str, Any]
) -> float | str | None:
    if "planning_velocity" in sprint:
        return float(sprint["planning_velocity"])
    if sprint_index == 0:
        default = cfg.get("default_planning_velocity")
        return default if default is not None else None
    if cfg.get("prefill_velocity_from_rolling_average", True):
        row = velocity_rolling_average_row(len(cfg["sprints"]))
        return f"=Velocity!$B${row}"
    default = cfg.get("default_planning_velocity")
    return default if default is not None else None


def suggested_limit_formula(
    team_capacity_row: int,
    planning_velocity_row: int,
    velocity_col: int = KPI_VALUE_COL,
) -> str:
    cap = f"D{team_capacity_row}"
    vel = f"{get_column_letter(velocity_col)}{planning_velocity_row}"
    return f'=IF({cap}=0,"",IF({vel}="",{cap},MIN({cap},{vel})))'


def committed_sp_sum_formula(first_row: int, last_row: int) -> str:
    """Sum story points where Committed? = Y."""
    sp = get_column_letter(BACKLOG_SP)
    com = get_column_letter(BACKLOG_COMMITTED)
    return f'=SUMPRODUCT(({com}{first_row}:{com}{last_row}="Y")*({sp}{first_row}:{sp}{last_row}))'


def member_assigned_sp_formula(member_row: int, first_row: int, last_row: int) -> str:
    sp = get_column_letter(BACKLOG_SP)
    asn = get_column_letter(BACKLOG_ASSIGNEE)
    com = get_column_letter(BACKLOG_COMMITTED)
    return (
        f'=SUMPRODUCT(({asn}${first_row}:${asn}${last_row}=A{member_row})*'
        f'({com}${first_row}:${com}${last_row}="Y")*({sp}${first_row}:${sp}${last_row}))'
    )


def sprint_sheet_layout(team_count: int, max_backlog_rows: int) -> dict[str, int]:
    capacity_last = CAPACITY_FIRST_DATA_ROW + team_count - 1
    team_total_row = capacity_last + 1
    backlog_title_row = team_total_row + 2
    backlog_header_row = backlog_title_row + 1
    backlog_first_data = backlog_header_row + 1
    backlog_last_data = backlog_first_data + max_backlog_rows - 1
    summary_title_row = backlog_last_data + 2
    summary_first_row = summary_title_row + 1
    summary_last_row = summary_first_row + 5
    return {
        "capacity_last": capacity_last,
        "team_total_row": team_total_row,
        "backlog_title_row": backlog_title_row,
        "backlog_header_row": backlog_header_row,
        "backlog_first_data": backlog_first_data,
        "backlog_last_data": backlog_last_data,
        "summary_title_row": summary_title_row,
        "summary_first_row": summary_first_row,
        "summary_last_row": summary_last_row,
    }


def build_sprint_overview_kpis(
    ws,
    tr: int,
    bf: int,
    bl: int,
    sprint: dict[str, Any],
    cfg: dict[str, Any],
    sprint_index: int,
    previous_sprint_meta: dict[str, Any] | None,
):
    vel_col = get_column_letter(KPI_VALUE_COL)
    overview_fill = PatternFill("solid", fgColor="F2F7FB")

    ws.cell(row=META_TITLE_ROW, column=KPI_LABEL_COL, value="Sprint overview")
    ws.cell(row=META_TITLE_ROW, column=KPI_LABEL_COL).font = Font(bold=True)

    kpi_labels = [
        (KPI_PLANNING_VEL_ROW, "Planning velocity (SP)"),
        (KPI_SUGGESTED_LIMIT_ROW, "Suggested SP limit"),
        (KPI_PREV_VELOCITY_ROW, "Previous sprint velocity"),
        (KPI_COMMITTED_ROW, "Committed SP (Y)"),
        (KPI_BUFFER_ROW, "Buffer to limit (SP)"),
        (KPI_STATUS_ROW, "Commitment status"),
    ]
    for row, label in kpi_labels:
        lbl = ws.cell(row=row, column=KPI_LABEL_COL, value=label)
        lbl.font = Font(bold=True)
        lbl.fill = overview_fill

    vel_initial = initial_planning_velocity(sprint_index, sprint, cfg)
    vel_cell = ws.cell(
        row=KPI_PLANNING_VEL_ROW,
        column=KPI_VALUE_COL,
        value=vel_initial if vel_initial is not None else "",
    )
    vel_cell.fill = SUMMARY_FILL

    limit_cell = ws.cell(
        row=KPI_SUGGESTED_LIMIT_ROW,
        column=KPI_VALUE_COL,
        value=suggested_limit_formula(tr, KPI_PLANNING_VEL_ROW),
    )
    limit_cell.fill = WARN_FILL

    prev_cell = ws.cell(row=KPI_PREV_VELOCITY_ROW, column=KPI_VALUE_COL)
    if previous_sprint_meta:
        prev_cell.value = f"={previous_sprint_meta['completed_cell']}"
    else:
        prev_cell.value = "n/a (first sprint)"
    prev_cell.fill = overview_fill

    committed = committed_sp_sum_formula(bf, bl)
    ws.cell(row=KPI_COMMITTED_ROW, column=KPI_VALUE_COL, value=committed).fill = SECTION_FILL

    lim = f"{vel_col}{KPI_SUGGESTED_LIMIT_ROW}"
    com = f"{vel_col}{KPI_COMMITTED_ROW}"
    ws.cell(
        row=KPI_BUFFER_ROW,
        column=KPI_VALUE_COL,
        value=f'=IF({lim}="","",{lim}-{com})',
    ).fill = SECTION_FILL

    ws.cell(
        row=KPI_STATUS_ROW,
        column=KPI_VALUE_COL,
        value=f'=IF({lim}="","",IF({com}>{lim},"OVER COMMITTED","On track"))',
    ).font = Font(bold=True)

    ws.conditional_formatting.add(
        f"{vel_col}{KPI_STATUS_ROW}",
        FormulaRule(
            formula=[f'${vel_col}${KPI_STATUS_ROW}="OVER COMMITTED"'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(bold=True, color="9C0006"),
            stopIfTrue=True,
        ),
    )


def build_sprint_sheet(
    ws,
    sprint: dict[str, Any],
    cfg: dict[str, Any],
    sprint_index: int = 0,
    previous_sprint_meta: dict[str, Any] | None = None,
) -> dict[str, Any]:
    team = cfg["team_members"]
    max_rows = cfg["max_backlog_rows_per_sprint"]
    layout = sprint_sheet_layout(len(team), max_rows)

    sheet_name = str(sprint.get("sheet_name", "Sprint"))[:31]
    ws.title = sheet_name

    title = sprint.get("title") or sheet_name
    start = parse_date(str(sprint["start"]))
    working_days = resolve_sprint_length_working_days(sprint, cfg)
    goal = sprint.get("goal", "")

    ws.cell(row=META_TITLE_ROW, column=1, value="Sprint")
    ws.cell(row=META_TITLE_ROW, column=2, value=title)
    ws.cell(row=META_START_ROW, column=1, value="Start")
    ws.cell(row=META_START_ROW, column=2, value=start)
    ws.cell(row=META_START_ROW, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=META_END_ROW, column=1, value="End (auto)")
    end_cell = ws.cell(row=META_END_ROW, column=2, value=sprint_end_formula(working_days))
    end_cell.number_format = "yyyy-mm-dd"
    ws.cell(
        row=META_END_ROW,
        column=3,
        value=f"{working_days} working days from start",
    )
    ws.cell(row=META_GOAL_ROW, column=1, value="Sprint goal")
    ws.cell(row=META_GOAL_ROW, column=2, value=goal)
    ws.merge_cells(
        start_row=META_GOAL_ROW,
        start_column=2,
        end_row=META_GOAL_ROW,
        end_column=4,
    )
    ws.cell(row=META_WORKING_DAYS_ROW, column=1, value="Working days (Mon–Fri)")
    ws.cell(
        row=META_WORKING_DAYS_ROW,
        column=2,
        value=f'=NETWORKDAYS(B{META_START_ROW},B{META_END_ROW})',
    )

    bf = layout["backlog_first_data"]
    bl = layout["backlog_last_data"]
    tr = layout["team_total_row"]

    ws.cell(row=CAPACITY_TITLE_ROW, column=1, value="Team capacity & load")
    ws.cell(row=CAPACITY_TITLE_ROW, column=1).font = Font(bold=True, size=12)

    headers = [
        "Team member",
        "PTO days",
        "FTE %",
        "Capacity (SP)",
        "Assigned SP (Y)",
        "Headroom (SP)",
    ]
    for col, label in enumerate(headers, start=COL_MEMBER):
        ws.cell(row=CAPACITY_HEADER_ROW, column=col, value=label)
    style_header_row(ws, CAPACITY_HEADER_ROW, COL_MEMBER, COL_HEADROOM)

    sp_day = cfg["story_points_per_day"]
    default_fte = cfg["default_fte_percent"]
    working_days_ref = f"$B${META_WORKING_DAYS_ROW}"

    for i, name in enumerate(team):
        row = CAPACITY_FIRST_DATA_ROW + i
        ws.cell(row=row, column=COL_MEMBER, value=name)
        ws.cell(row=row, column=COL_PTO, value=0)
        ws.cell(row=row, column=COL_FTE, value=default_fte)
        ws.cell(
            row=row,
            column=COL_CAPACITY,
            value=(
                f"=MAX(0,ROUND(({working_days_ref}-B{row})*C{row}/100*{sp_day},1))"
            ),
        )
        ws.cell(
            row=row,
            column=COL_ASSIGNED,
            value=member_assigned_sp_formula(row, bf, bl),
        )
        ws.cell(
            row=row,
            column=COL_HEADROOM,
            value=f"=D{row}-E{row}",
        )

    ws.cell(row=tr, column=COL_MEMBER, value="Team totals")
    ws.cell(row=tr, column=COL_MEMBER).font = Font(bold=True)
    ws.cell(
        row=tr,
        column=COL_CAPACITY,
        value=f"=SUM(D{CAPACITY_FIRST_DATA_ROW}:D{layout['capacity_last']})",
    )
    ws.cell(
        row=tr,
        column=COL_ASSIGNED,
        value=f"=SUM(E{CAPACITY_FIRST_DATA_ROW}:E{layout['capacity_last']})",
    )
    ws.cell(
        row=tr,
        column=COL_HEADROOM,
        value=f"=SUM(F{CAPACITY_FIRST_DATA_ROW}:F{layout['capacity_last']})",
    )
    for col in (COL_CAPACITY, COL_ASSIGNED, COL_HEADROOM):
        cell = ws.cell(row=tr, column=col)
        cell.font = Font(bold=True)
        cell.fill = SECTION_FILL

    build_sprint_overview_kpis(
        ws, tr, bf, bl, sprint, cfg, sprint_index, previous_sprint_meta
    )

    bt = layout["backlog_title_row"]
    ws.cell(row=bt, column=1, value="Sprint backlog — set Committed? to Y for planning")
    ws.cell(row=bt, column=1).font = Font(bold=True, size=12)

    bh = layout["backlog_header_row"]
    backlog_headers = [
        "ID",
        "Title",
        "Priority",
        "Type",
        "Story points",
        "Assignee",
        "Committed?",
        "Status",
        "Dependency / blocker",
        "Notes",
        "Done SP",
    ]
    for col, label in enumerate(backlog_headers, start=1):
        ws.cell(row=bh, column=col, value=label)
    style_header_row(ws, bh, 1, BACKLOG_DONE_SP)

    status_col = get_column_letter(BACKLOG_STATUS)
    for r in range(bf, bl + 1):
        ws.cell(
            row=r,
            column=BACKLOG_DONE_SP,
            value=f'=IF({status_col}{r}="Done",E{r},0)',
        )

    st = layout["summary_title_row"]
    ws.cell(row=st, column=1, value="Sprint results (after sprint)")
    ws.cell(row=st, column=1).font = Font(bold=True, size=12)

    committed_formula = committed_sp_sum_formula(bf, bl)
    done_col = get_column_letter(BACKLOG_DONE_SP)
    done_range = f"{done_col}{bf}:{done_col}{bl}"
    kpi_committed = f"{get_column_letter(KPI_VALUE_COL)}{KPI_COMMITTED_ROW}"

    summary_rows = [
        ("Committed story points", committed_formula),
        ("Completed story points", f"=SUM({done_range})"),
        ("Velocity (completed SP)", f"=B{layout['summary_first_row'] + 1}"),
        (
            "Carry-over (committed − completed)",
            f"=B{layout['summary_first_row']}-B{layout['summary_first_row'] + 1}",
        ),
        (
            "Completion %",
            f'=IF(B{layout["summary_first_row"]}=0,"",B{layout["summary_first_row"] + 1}/B{layout["summary_first_row"]})',
        ),
        ("Committed SP (overview link)", f"={kpi_committed}"),
    ]

    for offset, (label, formula) in enumerate(summary_rows):
        row = layout["summary_first_row"] + offset
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=1).font = Font(bold=True)
        cell = ws.cell(row=row, column=2, value=formula)
        cell.fill = SUMMARY_FILL
        if "Completion" in label:
            cell.number_format = "0%"

    ws.cell(
        row=layout["summary_last_row"] + 1,
        column=1,
        value="Tip: mark candidates with Committed?=Y; watch overview status and per-person headroom.",
    )
    ws.merge_cells(
        start_row=layout["summary_last_row"] + 1,
        start_column=1,
        end_row=layout["summary_last_row"] + 1,
        end_column=11,
    )

    sp_list = list_formula([str(x) for x in cfg["story_point_scale"]])
    type_list = list_formula(cfg["work_item_types"])
    status_list = list_formula(cfg["statuses"])
    assignee_list = list_formula(team)
    priority_list = list_formula(cfg["priorities"])
    committed_list = list_formula(["Y", "N"])

    dv_sp = DataValidation(type="list", formula1=sp_list, allow_blank=True)
    dv_type = DataValidation(type="list", formula1=type_list, allow_blank=True)
    dv_status = DataValidation(type="list", formula1=status_list, allow_blank=True)
    dv_assignee = DataValidation(type="list", formula1=assignee_list, allow_blank=True)
    dv_priority = DataValidation(type="list", formula1=priority_list, allow_blank=True)
    dv_committed = DataValidation(type="list", formula1=committed_list, allow_blank=True)

    if bl >= bf:
        cols = {
            BACKLOG_SP: dv_sp,
            BACKLOG_TYPE: dv_type,
            BACKLOG_STATUS: dv_status,
            BACKLOG_ASSIGNEE: dv_assignee,
            BACKLOG_PRIORITY: dv_priority,
            BACKLOG_COMMITTED: dv_committed,
        }
        for col_idx, dv in cols.items():
            letter = get_column_letter(col_idx)
            dv.add(f"{letter}{bf}:{letter}{bl}")
            ws.add_data_validation(dv)

    title_col = get_column_letter(BACKLOG_TITLE)
    headroom_col = get_column_letter(COL_HEADROOM)
    ws.conditional_formatting.add(
        f"{title_col}{bf}:{title_col}{bl}",
        FormulaRule(
            formula=[f'${status_col}{bf}="Done"'],
            fill=DONE_FILL,
            stopIfTrue=True,
        ),
    )
    ws.conditional_formatting.add(
        f"{headroom_col}{CAPACITY_FIRST_DATA_ROW}:{headroom_col}{layout['capacity_last']}",
        FormulaRule(
            formula=[f"${headroom_col}{CAPACITY_FIRST_DATA_ROW}<0"],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            stopIfTrue=True,
        ),
    )

    apply_table_borders(ws, CAPACITY_HEADER_ROW, tr, COL_MEMBER, COL_HEADROOM)
    apply_table_borders(ws, KPI_PLANNING_VEL_ROW, KPI_STATUS_ROW, KPI_LABEL_COL, KPI_VALUE_COL)
    apply_table_borders(ws, bh, bl, 1, BACKLOG_DONE_SP)
    apply_table_borders(
        ws,
        layout["summary_first_row"],
        layout["summary_first_row"] + len(summary_rows) - 1,
        1,
        2,
    )

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 11
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 18
    ws.column_dimensions["J"].width = 22
    ws.column_dimensions["K"].width = 10

    ws.freeze_panes = ws.cell(row=bf, column=BACKLOG_TITLE).coordinate

    start_cell = f"'{sheet_name}'!$B${META_START_ROW}"
    end_cell_ref = f"'{sheet_name}'!$B${META_END_ROW}"
    vel_col = get_column_letter(KPI_VALUE_COL)

    return {
        "sheet_name": sheet_name,
        "goal": goal,
        "start_cell": start_cell,
        "end_cell": end_cell_ref,
        "committed_cell": f"'{sheet_name}'!$B${layout['summary_first_row']}",
        "completed_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 1}",
        "velocity_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 1}",
        "carryover_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 3}",
        "completion_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 4}",
        "capacity_cell": f"'{sheet_name}'!$D${tr}",
        "load_cell": f"='{sheet_name}'!${vel_col}${KPI_COMMITTED_ROW}/'{sheet_name}'!$D${tr}",
        "suggested_limit_cell": f"'{sheet_name}'!${vel_col}${KPI_SUGGESTED_LIMIT_ROW}",
        "planning_velocity_cell": f"'{sheet_name}'!${vel_col}${KPI_PLANNING_VEL_ROW}",
        "working_days_cell": f"'{sheet_name}'!$B${META_WORKING_DAYS_ROW}",
    }


def build_velocity_sheet(ws, sprint_meta: list[dict[str, Any]], cfg: dict[str, Any]):
    ws.title = "Velocity"

    ws.cell(row=1, column=1, value="Velocity & sprint history")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.cell(
        row=2,
        column=1,
        value="Completed metrics pull from sprint tabs. Enter Planning velocity on each sprint for commit limits.",
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)

    header_row = 4
    headers = [
        "Sprint sheet",
        "Start",
        "End",
        "Working days",
        "Sprint goal",
        "Committed SP",
        "Completed SP",
        "Velocity",
        "Carry-over SP",
        "Completion %",
        "Load / capacity",
    ]
    for col, label in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col, value=label)
    style_header_row(ws, header_row, 1, len(headers))

    data_start = header_row + 1
    for i, meta in enumerate(sprint_meta):
        row = data_start + i
        ws.cell(row=row, column=1, value=meta["sheet_name"])
        ws.cell(row=row, column=2, value=f"={meta['start_cell']}")
        ws.cell(row=row, column=2).number_format = "yyyy-mm-dd"
        ws.cell(row=row, column=3, value=f"={meta['end_cell']}")
        ws.cell(row=row, column=3).number_format = "yyyy-mm-dd"
        ws.cell(row=row, column=4, value=f"={meta['working_days_cell']}")
        ws.cell(row=row, column=5, value=meta["goal"])
        ws.cell(row=row, column=6, value=f"={meta['committed_cell']}")
        ws.cell(row=row, column=7, value=f"={meta['completed_cell']}")
        ws.cell(row=row, column=8, value=f"={meta['velocity_cell']}")
        ws.cell(row=row, column=9, value=f"={meta['carryover_cell']}")
        ws.cell(row=row, column=10, value=f"={meta['completion_cell']}")
        ws.cell(row=row, column=10).number_format = "0%"
        ws.cell(row=row, column=11, value=f"={meta['load_cell']}")
        ws.cell(row=row, column=11).number_format = "0%"

    if sprint_meta:
        last_data = data_start + len(sprint_meta) - 1
        stats_row = last_data + 2
        ws.cell(row=stats_row, column=1, value="Rolling average velocity (completed SP)")
        ws.cell(row=stats_row, column=1).font = Font(bold=True)
        ws.cell(
            row=stats_row,
            column=2,
            value=f"=AVERAGE(H{data_start}:H{last_data})",
        )
        ws.cell(row=stats_row, column=2).fill = SECTION_FILL

        ws.cell(row=stats_row + 1, column=1, value="Total carry-over (all sprints)")
        ws.cell(row=stats_row + 1, column=1).font = Font(bold=True)
        ws.cell(
            row=stats_row + 1,
            column=2,
            value=f"=SUM(I{data_start}:I{last_data})",
        )

        ws.cell(row=stats_row + 2, column=1, value="Suggested next-sprint capacity (avg velocity)")
        ws.cell(row=stats_row + 2, column=1).font = Font(bold=True)
        ws.cell(row=stats_row + 2, column=2, value=f"=B{stats_row}")
        ws.cell(row=stats_row + 2, column=2).fill = WARN_FILL

        apply_table_borders(ws, header_row, last_data, 1, len(headers))
        vel_range = f"H{data_start}:H{last_data}"
        ws.conditional_formatting.add(
            vel_range,
            ColorScaleRule(
                start_type="min",
                start_color="F8696B",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFEB84",
                end_type="max",
                end_color="63BE7B",
            ),
        )

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 36
    for col in "FGHIJK":
        ws.column_dimensions[col].width = 14


def build_product_backlog_sheet(ws, cfg: dict[str, Any]):
    ws.title = "Product backlog"
    rows = cfg["product_backlog_rows"]
    team = cfg["team_members"]

    ws.cell(row=1, column=1, value="Product backlog (not yet in a sprint)")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.cell(
        row=2,
        column=1,
        value="Prioritise top-down. Move items into a sprint sheet when committed.",
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    header_row = 4
    headers = [
        "Rank",
        "ID",
        "Title",
        "Type",
        "Story points",
        "Target sprint",
        "Notes",
        "Ready?",
    ]
    for col, label in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col, value=label)
    style_header_row(ws, header_row, 1, len(headers))

    first = header_row + 1
    last = first + rows - 1
    sprint_names = [str(s.get("sheet_name", "")) for s in cfg["sprints"]]
    target_list = list_formula([""] + sprint_names)

    for r in range(first, last + 1):
        ws.cell(row=r, column=1, value=r - first + 1)

    dv_type = DataValidation(
        type="list", formula1=list_formula(cfg["work_item_types"]), allow_blank=True
    )
    dv_sp = DataValidation(
        type="list",
        formula1=list_formula([str(x) for x in cfg["story_point_scale"]]),
        allow_blank=True,
    )
    dv_target = DataValidation(type="list", formula1=target_list, allow_blank=True)
    dv_ready = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)

    dv_type.add(f"C{first}:C{last}")
    dv_sp.add(f"E{first}:E{last}")
    dv_target.add(f"F{first}:F{last}")
    dv_ready.add(f"H{first}:H{last}")
    for dv in (dv_type, dv_sp, dv_target, dv_ready):
        ws.add_data_validation(dv)

    apply_table_borders(ws, header_row, last, 1, len(headers))
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 8
    ws.freeze_panes = ws.cell(row=first, column=2).coordinate


def generate(config_path: Path) -> Path:
    cfg = load_config(config_path)
    base_dir = config_path.resolve().parent

    wb = Workbook()
    wb.remove(wb.active)

    build_product_backlog_sheet(wb.create_sheet(), cfg)

    sprint_meta: list[dict[str, Any]] = []
    previous_meta: dict[str, Any] | None = None
    for i, entry in enumerate(cfg["sprints"]):
        if not isinstance(entry, dict):
            raise ValueError('Each "sprints" item must be an object.')
        for key in ("sheet_name", "start"):
            if key not in entry:
                raise ValueError(f'Sprint entry missing "{key}".')
        ws = wb.create_sheet()
        meta = build_sprint_sheet(
            ws, entry, cfg, sprint_index=i, previous_sprint_meta=previous_meta
        )
        sprint_meta.append(meta)
        previous_meta = meta

    build_velocity_sheet(wb.create_sheet(), sprint_meta, cfg)
    if sprint_meta:
        wb.move_sheet(wb["Velocity"], offset=-len(sprint_meta))

    out = base_dir / cfg["output_file"]
    wb.save(out)
    return out


def main():
    parser = argparse.ArgumentParser(
        description="Generate Scrum sprint planning Excel workbook."
    )
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=None,
        help=f"Path to JSON config (default: {DEFAULT_CONFIG_NAME} next to this script)",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    config_path = args.config or (script_dir / DEFAULT_CONFIG_NAME)
    if not config_path.is_file():
        print(f"Config not found: {config_path}", file=sys.stderr)
        sys.exit(1)

    try:
        out = generate(config_path)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Created: {out}")


if __name__ == "__main__":
    main()
