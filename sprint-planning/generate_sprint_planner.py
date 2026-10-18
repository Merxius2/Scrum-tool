#!/usr/bin/env python3
"""Generate Scrum sprint planning Excel workbook from JSON config."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DEFAULT_CONFIG_NAME = "sprint_planner.config.json"
INSTRUCTIONS_SHEET_NAME = "Instructions"
PRODUCT_BACKLOG_SHEET = "Product backlog"
VELOCITY_SHEET = "Velocity"

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
    if sprints is None:
        sprints = []
    if not isinstance(sprints, list):
        raise ValueError('"sprints" must be a list of sprint definitions.')

    add_sprints = data.get("add_sprints")
    if add_sprints is None:
        add_sprints = []
    if not isinstance(add_sprints, list):
        raise ValueError('"add_sprints" must be a list of sprint definitions.')

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
        "add_sprints": add_sprints,
    }


def system_sheet_names() -> set[str]:
    return {PRODUCT_BACKLOG_SHEET, VELOCITY_SHEET, INSTRUCTIONS_SHEET_NAME}


def ordered_sprint_sheet_names(wb) -> list[str]:
    return [name for name in wb.sheetnames if name not in system_sheet_names()]


def sprint_meta_from_sheet_name(sheet_name: str, cfg: dict[str, Any]) -> dict[str, Any]:
    """Build Velocity cross-references for an existing sprint tab (fixed layout)."""
    layout = sprint_sheet_layout(len(cfg["team_members"]), cfg["max_backlog_rows_per_sprint"])
    tr = layout["team_total_row"]
    sf = layout["summary_first_row"]
    vel_col = get_column_letter(KPI_VALUE_COL)
    quoted = sheet_name.replace("'", "''")
    return {
        "sheet_name": sheet_name,
        "goal_cell": f"'{quoted}'!$B${META_GOAL_ROW}",
        "start_cell": f"'{quoted}'!$B${META_START_ROW}",
        "end_cell": f"'{quoted}'!$B${META_END_ROW}",
        "committed_cell": f"'{quoted}'!$B${sf}",
        "completed_cell": f"'{quoted}'!$B${sf + 1}",
        "velocity_cell": f"'{quoted}'!$B${sf + 1}",
        "carryover_cell": f"'{quoted}'!$B${sf + 3}",
        "completion_cell": f"'{quoted}'!$B${sf + 4}",
        "capacity_cell": f"'{quoted}'!$D${tr}",
        "load_cell": (
            f"=IF('{quoted}'!$D${tr}=0,\"\","
            f"'{quoted}'!${vel_col}${KPI_COMMITTED_ROW}/'{quoted}'!$D${tr})"
        ),
        "working_days_cell": f"'{quoted}'!$B${META_WORKING_DAYS_ROW}",
    }


def lock_all_cells(ws) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == "MergedCell":
                continue
            cell.protection = Protection(locked=True)


def unlock_cell(ws, row: int, col: int) -> None:
    ws.cell(row=row, column=col).protection = Protection(locked=False)


def enable_sheet_protection(ws) -> None:
    ws.protection.sheet = True
    ws.protection.enable()


def protect_sprint_worksheet(
    ws,
    layout: dict[str, int],
    bf: int,
    bl: int,
) -> None:
    lock_all_cells(ws)
    unlock_cell(ws, META_TITLE_ROW, 2)
    unlock_cell(ws, META_START_ROW, 2)
    unlock_cell(ws, META_GOAL_ROW, 2)
    unlock_cell(ws, KPI_PLANNING_VEL_ROW, KPI_VALUE_COL)
    for row in range(CAPACITY_FIRST_DATA_ROW, layout["capacity_last"] + 1):
        unlock_cell(ws, row, COL_PTO)
        unlock_cell(ws, row, COL_FTE)
    for row in range(bf, bl + 1):
        for col in range(BACKLOG_ID, BACKLOG_NOTES + 1):
            unlock_cell(ws, row, col)
    enable_sheet_protection(ws)


def protect_product_backlog_worksheet(ws, first_row: int, last_row: int) -> None:
    lock_all_cells(ws)
    for row in range(first_row, last_row + 1):
        for col in range(1, 9):
            unlock_cell(ws, row, col)
    enable_sheet_protection(ws)


def protect_readonly_worksheet(ws) -> None:
    lock_all_cells(ws)
    enable_sheet_protection(ws)


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
    sprint_index: int,
    sprint: dict[str, Any],
    cfg: dict[str, Any],
    total_sprint_count: int,
) -> float | str | None:
    if "planning_velocity" in sprint:
        return float(sprint["planning_velocity"])
    if sprint_index == 0:
        default = cfg.get("default_planning_velocity")
        return default if default is not None else None
    if cfg.get("prefill_velocity_from_rolling_average", True) and total_sprint_count > 0:
        row = velocity_rolling_average_row(total_sprint_count)
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
    total_sprint_count: int,
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

    vel_initial = initial_planning_velocity(
        sprint_index, sprint, cfg, total_sprint_count
    )
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
    total_sprint_count: int | None = None,
) -> dict[str, Any]:
    team = cfg["team_members"]
    max_rows = cfg["max_backlog_rows_per_sprint"]
    layout = sprint_sheet_layout(len(team), max_rows)

    sheet_name = str(sprint.get("sheet_name", "Sprint"))[:31]
    ws.title = sheet_name

    title = sprint.get("title") or sheet_name
    working_days = resolve_sprint_length_working_days(sprint, cfg)
    goal = sprint.get("goal", "")
    if total_sprint_count is None:
        total_sprint_count = sprint_index + 1

    ws.cell(row=META_TITLE_ROW, column=1, value="Sprint")
    ws.cell(row=META_TITLE_ROW, column=2, value=title)
    ws.cell(row=META_START_ROW, column=1, value="Start")
    start_cell = ws.cell(row=META_START_ROW, column=2)
    start_raw = sprint.get("start")
    if start_raw:
        start_cell.value = parse_date(str(start_raw))
    else:
        start_cell.value = None
    start_cell.number_format = "yyyy-mm-dd"
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
        ws,
        tr,
        bf,
        bl,
        sprint,
        cfg,
        sprint_index,
        previous_sprint_meta,
        total_sprint_count,
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

    protect_sprint_worksheet(ws, layout, bf, bl)

    quoted = sheet_name.replace("'", "''")
    start_cell = f"'{quoted}'!$B${META_START_ROW}"
    end_cell_ref = f"'{quoted}'!$B${META_END_ROW}"
    vel_col = get_column_letter(KPI_VALUE_COL)

    return {
        "sheet_name": sheet_name,
        "goal": goal,
        "goal_cell": f"'{quoted}'!$B${META_GOAL_ROW}",
        "start_cell": start_cell,
        "end_cell": end_cell_ref,
        "committed_cell": f"'{quoted}'!$B${layout['summary_first_row']}",
        "completed_cell": f"'{quoted}'!$B${layout['summary_first_row'] + 1}",
        "velocity_cell": f"'{quoted}'!$B${layout['summary_first_row'] + 1}",
        "carryover_cell": f"'{quoted}'!$B${layout['summary_first_row'] + 3}",
        "completion_cell": f"'{quoted}'!$B${layout['summary_first_row'] + 4}",
        "capacity_cell": f"'{quoted}'!$D${tr}",
        "load_cell": (
            f"=IF('{quoted}'!$D${tr}=0,\"\","
            f"'{quoted}'!${vel_col}${KPI_COMMITTED_ROW}/'{quoted}'!$D${tr})"
        ),
        "suggested_limit_cell": f"'{quoted}'!${vel_col}${KPI_SUGGESTED_LIMIT_ROW}",
        "planning_velocity_cell": f"'{quoted}'!${vel_col}${KPI_PLANNING_VEL_ROW}",
        "working_days_cell": f"'{quoted}'!$B${META_WORKING_DAYS_ROW}",
    }


def build_velocity_sheet(ws, sprint_meta: list[dict[str, Any]], cfg: dict[str, Any]):
    ws.title = "Velocity"

    ws.cell(row=1, column=1, value="Velocity & sprint history")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.cell(
        row=2,
        column=1,
        value="One row per sprint tab. Committed and completed SP are formulas linked to each sprint — do not type here.",
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
        ws.cell(row=row, column=5, value=f"={meta['goal_cell']}")
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

    protect_readonly_worksheet(ws)


def rebuild_velocity_tab(wb, sprint_meta: list[dict[str, Any]], cfg: dict[str, Any]) -> None:
    if VELOCITY_SHEET in wb.sheetnames:
        idx = wb.sheetnames.index(VELOCITY_SHEET)
        del wb[VELOCITY_SHEET]
    else:
        idx = 1
    build_velocity_sheet(wb.create_sheet(VELOCITY_SHEET, idx), sprint_meta, cfg)


def update_product_backlog_sprint_list(ws, sprint_names: list[str]) -> None:
    """Refresh Target sprint dropdown to include all sprint tabs."""
    header_row = 4
    first = header_row + 1
    rows = ws.max_row
    if rows < first:
        return
    target_list = list_formula([""] + sprint_names)
    dv_target = DataValidation(type="list", formula1=target_list, allow_blank=True)
    dv_target.add(f"F{first}:F{rows}")
    ws.add_data_validation(dv_target)


def build_product_backlog_sheet(
    ws, cfg: dict[str, Any], sprint_names: list[str] | None = None
):
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
    if sprint_names is None:
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
    protect_product_backlog_worksheet(ws, first, last)


def _write_instructions_block(
    ws,
    row: int,
    heading: str,
    lines: list[str],
) -> int:
    """Write a heading and bullet lines; return next free row."""
    ws.cell(row=row, column=1, value=heading).font = Font(bold=True, size=12)
    row += 1
    wrap = Alignment(wrap_text=True, vertical="top")
    for text in lines:
        cell = ws.cell(row=row, column=1, value=text)
        cell.alignment = wrap
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        row += 1
    return row + 1


def build_instructions_sheet(ws):
    ws.title = INSTRUCTIONS_SHEET_NAME[:31]

    ws.cell(row=1, column=1, value="How to use this sprint planning workbook")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws.cell(
        row=2,
        column=1,
        value="Fill in yellow/editable cells in Excel. Formulas and dropdowns update capacity, "
        "commitments, and velocity automatically. Regenerate the file from sprint_planner.config.json "
        "only when team names, sprint tabs, or layout change.",
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    ws.cell(row=2, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 36

    row = 4
    row = _write_instructions_block(
        ws,
        row,
        "Workbook tabs (left to right)",
        [
            "• Product backlog — ranked ideas not yet in a sprint; set Target sprint when ready.",
            "• Velocity — history across sprints; rolling average helps set the next planning velocity.",
            "• One tab per sprint — plan capacity, commit backlog items, record results.",
            "• Instructions (this tab) — always last; safe to hide in Excel if you prefer.",
        ],
    )
    row = _write_instructions_block(
        ws,
        row,
        "Before sprint planning",
        [
            "• On the sprint tab, set Start date and Sprint goal. End date and working days calculate automatically.",
            "• Enter PTO days and FTE % for each team member in Team capacity & load.",
            "• In Sprint overview (columns F–G), enter Planning velocity (SP). Use Previous sprint velocity "
            "as a guide; Sprint 2+ may prefill from the Velocity tab average.",
            "• Check Suggested SP limit (MIN of team capacity and planning velocity) before committing work.",
        ],
    )
    row = _write_instructions_block(
        ws,
        row,
        "Committing work (sprint backlog)",
        [
            "• Add rows: ID, Title, Priority, Type, Story points, Assignee.",
            "• Set Committed? to Y only for items the team agrees to deliver this sprint. Leave N or blank for candidates.",
            "• Note dependencies or blockers in the Dependency column.",
            "• Committed SP in the overview counts only rows with Committed? = Y.",
            "• Watch Assigned SP and Headroom per person — red headroom means that person is over their capacity.",
            "• Commitment status shows OVER COMMITTED if committed SP exceeds the suggested limit.",
        ],
    )
    row = _write_instructions_block(
        ws,
        row,
        "During and after the sprint",
        [
            "• Update Status (To Do → In Progress → Done). Done SP and Sprint results update from Done items.",
            "• Velocity (completed SP) on the tab and Velocity sheet reflects completed story points.",
            "• Use carry-over and completion % to discuss unfinished committed work in the retrospective.",
        ],
    )
    row = _write_instructions_block(
        ws,
        row,
        "Product backlog tab",
        [
            "• Keep the list ordered by rank (top = highest priority).",
            "• Set Ready? when an item meets your definition of ready before pulling it into a sprint.",
            "• Copy or re-enter items on a sprint tab when committing; link Target sprint for traceability.",
        ],
    )
    _write_instructions_block(
        ws,
        row,
        "Regenerate vs edit in Excel",
        [
            "• Edit names, dates, backlog, and statuses directly in Excel day to day.",
            "• Add sprint tabs: put entries in add_sprints in sprint_planner.config.json, then run "
            "python3 generate_sprint_planner.py (existing sprint data is kept). Use --force only to "
            "recreate the whole workbook from scratch.",
            "• Formula cells are locked; editable fields (dates, capacity inputs, backlog, planning velocity) "
            "stay unlocked. Unprotect a sheet in Excel if you must change structure.",
        ],
    )

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.sheet_view.showGridLines = True
    protect_readonly_worksheet(ws)


def validate_sprint_entry(entry: dict[str, Any]) -> None:
    if "sheet_name" not in entry:
        raise ValueError('Each sprint entry needs a "sheet_name".')


def create_fresh_workbook(cfg: dict[str, Any]) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    build_product_backlog_sheet(wb.create_sheet(), cfg, sprint_names=[])

    sprint_meta: list[dict[str, Any]] = []
    previous_meta: dict[str, Any] | None = None
    total = len(cfg["sprints"])
    for i, entry in enumerate(cfg["sprints"]):
        if not isinstance(entry, dict):
            raise ValueError('Each "sprints" item must be an object.')
        validate_sprint_entry(entry)
        ws = wb.create_sheet()
        meta = build_sprint_sheet(
            ws,
            entry,
            cfg,
            sprint_index=i,
            previous_sprint_meta=previous_meta,
            total_sprint_count=total,
        )
        sprint_meta.append(meta)
        previous_meta = meta

    if PRODUCT_BACKLOG_SHEET in wb.sheetnames:
        update_product_backlog_sprint_list(
            wb[PRODUCT_BACKLOG_SHEET],
            [m["sheet_name"] for m in sprint_meta],
        )

    rebuild_velocity_tab(wb, sprint_meta, cfg)
    build_instructions_sheet(wb.create_sheet())
    return wb


def add_sprints_to_workbook(wb, cfg: dict[str, Any], entries: list[dict[str, Any]]) -> int:
    existing = set(ordered_sprint_sheet_names(wb))
    added = 0
    sprint_meta = [sprint_meta_from_sheet_name(n, cfg) for n in ordered_sprint_sheet_names(wb)]
    previous_meta = sprint_meta[-1] if sprint_meta else None
    total_after = len(existing) + len(
        [e for e in entries if str(e.get("sheet_name", ""))[:31] not in existing]
    )

    instr_idx = (
        wb.sheetnames.index(INSTRUCTIONS_SHEET_NAME)
        if INSTRUCTIONS_SHEET_NAME in wb.sheetnames
        else len(wb.sheetnames)
    )

    for entry in entries:
        if not isinstance(entry, dict):
            raise ValueError('Each sprint entry must be an object.')
        validate_sprint_entry(entry)
        sheet_name = str(entry["sheet_name"])[:31]
        if sheet_name in existing:
            continue
        ws = wb.create_sheet(sheet_name, instr_idx)
        instr_idx += 1
        sprint_index = len(sprint_meta)
        meta = build_sprint_sheet(
            ws,
            entry,
            cfg,
            sprint_index=sprint_index,
            previous_sprint_meta=previous_meta,
            total_sprint_count=total_after,
        )
        sprint_meta.append(meta)
        previous_meta = meta
        existing.add(sheet_name)
        added += 1

    if PRODUCT_BACKLOG_SHEET in wb.sheetnames:
        update_product_backlog_sprint_list(
            wb[PRODUCT_BACKLOG_SHEET], ordered_sprint_sheet_names(wb)
        )

    rebuild_velocity_tab(wb, sprint_meta, cfg)
    return added


def generate(config_path: Path, force: bool = False) -> Path:
    cfg = load_config(config_path)
    base_dir = config_path.resolve().parent
    out = base_dir / cfg["output_file"]

    if out.is_file() and not force:
        return update_workbook(config_path, cfg, out)

    wb = create_fresh_workbook(cfg)
    wb.save(out)
    return out


def update_workbook(config_path: Path, cfg: dict[str, Any], out: Path) -> Path:
    wb = load_workbook(out)
    entries = cfg.get("add_sprints") or []
    if not entries:
        sprint_meta = [
            sprint_meta_from_sheet_name(n, cfg) for n in ordered_sprint_sheet_names(wb)
        ]
        rebuild_velocity_tab(wb, sprint_meta, cfg)
        wb.save(out)
        return out

    added = add_sprints_to_workbook(wb, cfg, entries)
    if added == 0:
        print("No new sprint tabs added (sheet names already exist).", file=sys.stderr)
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
    parser.add_argument(
        "--force",
        action="store_true",
        help="Recreate the entire workbook (default: update existing file, add sprints from add_sprints)",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    config_path = args.config or (script_dir / DEFAULT_CONFIG_NAME)
    if not config_path.is_file():
        print(f"Config not found: {config_path}", file=sys.stderr)
        sys.exit(1)

    try:
        out = generate(config_path, force=args.force)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
