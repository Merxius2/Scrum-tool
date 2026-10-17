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

CAPACITY_TITLE_ROW = 7
CAPACITY_HEADER_ROW = 8
CAPACITY_FIRST_DATA_ROW = 9

COL_MEMBER = 1
COL_PTO = 2
COL_AVAIL = 3
COL_FOCUS = 4
COL_CAPACITY = 5

BACKLOG_ID = 1
BACKLOG_TITLE = 2
BACKLOG_TYPE = 3
BACKLOG_SP = 4
BACKLOG_ASSIGNEE = 5
BACKLOG_STATUS = 6
BACKLOG_NOTES = 7
BACKLOG_DONE_SP = 8


def parse_date(value: str):
    return datetime.strptime(value, "%Y-%m-%d").date()


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
        "default_availability_percent": float(data.get("default_availability_percent", 100)),
        "default_focus_percent": float(data.get("default_focus_percent", 80)),
        "story_points_per_day": float(data.get("story_points_per_day", 1)),
        "max_backlog_rows_per_sprint": int(data.get("max_backlog_rows_per_sprint", 30)),
        "product_backlog_rows": int(data.get("product_backlog_rows", 40)),
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


def sprint_sheet_layout(team_count: int, max_backlog_rows: int) -> dict[str, int]:
    capacity_last = CAPACITY_FIRST_DATA_ROW + team_count - 1
    team_total_row = capacity_last + 1
    backlog_title_row = team_total_row + 2
    backlog_header_row = backlog_title_row + 1
    backlog_first_data = backlog_header_row + 1
    backlog_last_data = backlog_first_data + max_backlog_rows - 1
    summary_title_row = backlog_last_data + 2
    summary_first_row = summary_title_row + 1
    summary_last_row = summary_first_row + 7
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


def build_sprint_sheet(
    ws,
    sprint: dict[str, Any],
    cfg: dict[str, Any],
) -> dict[str, Any]:
    team = cfg["team_members"]
    max_rows = cfg["max_backlog_rows_per_sprint"]
    layout = sprint_sheet_layout(len(team), max_rows)

    sheet_name = str(sprint.get("sheet_name", "Sprint"))[:31]
    ws.title = sheet_name

    title = sprint.get("title") or sheet_name
    start = parse_date(str(sprint["start"]))
    end = parse_date(str(sprint["end"]))
    goal = sprint.get("goal", "")

    ws.cell(row=META_TITLE_ROW, column=1, value="Sprint")
    ws.cell(row=META_TITLE_ROW, column=2, value=title)
    ws.cell(row=META_START_ROW, column=1, value="Start")
    ws.cell(row=META_START_ROW, column=2, value=start)
    ws.cell(row=META_START_ROW, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=META_END_ROW, column=1, value="End")
    ws.cell(row=META_END_ROW, column=2, value=end)
    ws.cell(row=META_END_ROW, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=META_GOAL_ROW, column=1, value="Sprint goal")
    ws.cell(row=META_GOAL_ROW, column=2, value=goal)
    ws.merge_cells(
        start_row=META_GOAL_ROW,
        start_column=2,
        end_row=META_GOAL_ROW,
        end_column=8,
    )
    ws.cell(row=META_WORKING_DAYS_ROW, column=1, value="Working days (Mon–Fri)")
    ws.cell(
        row=META_WORKING_DAYS_ROW,
        column=2,
        value=f'=NETWORKDAYS(B{META_START_ROW},B{META_END_ROW})',
    )

    ws.cell(row=CAPACITY_TITLE_ROW, column=1, value="Sprint capacity")
    ws.cell(row=CAPACITY_TITLE_ROW, column=1).font = Font(bold=True, size=12)

    headers = [
        "Team member",
        "PTO days",
        "Availability %",
        "Focus %",
        "Capacity (SP)",
    ]
    for col, label in enumerate(headers, start=COL_MEMBER):
        ws.cell(row=CAPACITY_HEADER_ROW, column=col, value=label)
    style_header_row(ws, CAPACITY_HEADER_ROW, COL_MEMBER, COL_CAPACITY)

    sp_day = cfg["story_points_per_day"]
    default_avail = cfg["default_availability_percent"]
    default_focus = cfg["default_focus_percent"]
    working_days_ref = f"$B${META_WORKING_DAYS_ROW}"

    for i, name in enumerate(team):
        row = CAPACITY_FIRST_DATA_ROW + i
        ws.cell(row=row, column=COL_MEMBER, value=name)
        ws.cell(row=row, column=COL_PTO, value=0)
        ws.cell(row=row, column=COL_AVAIL, value=default_avail)
        ws.cell(row=row, column=COL_FOCUS, value=default_focus)
        ws.cell(
            row=row,
            column=COL_CAPACITY,
            value=(
                f"=MAX(0,ROUND(({working_days_ref}-B{row})*C{row}/100*D{row}/100*{sp_day},1))"
            ),
        )

    tr = layout["team_total_row"]
    ws.cell(row=tr, column=COL_MEMBER, value="Team capacity (SP)")
    ws.cell(row=tr, column=COL_MEMBER).font = Font(bold=True)
    ws.cell(
        row=tr,
        column=COL_CAPACITY,
        value=f"=SUM(E{CAPACITY_FIRST_DATA_ROW}:E{layout['capacity_last']})",
    )
    ws.cell(row=tr, column=COL_CAPACITY).font = Font(bold=True)
    ws.cell(row=tr, column=COL_CAPACITY).fill = SECTION_FILL

    bt = layout["backlog_title_row"]
    ws.cell(row=bt, column=1, value="Sprint backlog")
    ws.cell(row=bt, column=1).font = Font(bold=True, size=12)

    bh = layout["backlog_header_row"]
    backlog_headers = [
        "ID",
        "Title",
        "Type",
        "Story points",
        "Assignee",
        "Status",
        "Notes",
        "Done SP",
    ]
    for col, label in enumerate(backlog_headers, start=1):
        ws.cell(row=bh, column=col, value=label)
    style_header_row(ws, bh, 1, BACKLOG_DONE_SP)

    bf = layout["backlog_first_data"]
    bl = layout["backlog_last_data"]
    for r in range(bf, bl + 1):
        ws.cell(
            row=r,
            column=BACKLOG_DONE_SP,
            value=f'=IF(F{r}="Done",D{r},0)',
        )

    st = layout["summary_title_row"]
    ws.cell(row=st, column=1, value="Sprint results")
    ws.cell(row=st, column=1).font = Font(bold=True, size=12)

    sp_range = f"D{bf}:D{bl}"
    done_range = f"H{bf}:H{bl}"
    status_range = f"F{bf}:F{bl}"
    team_cap = f"E{tr}"

    summary_rows = [
        ("Committed story points", f"=SUM({sp_range})"),
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
        ("Team capacity (SP)", f"={team_cap}"),
        (
            "Load vs capacity",
            f'=IF({team_cap}=0,"",B{layout["summary_first_row"]}/{team_cap})',
        ),
    ]

    for offset, (label, formula) in enumerate(summary_rows):
        row = layout["summary_first_row"] + offset
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=1).font = Font(bold=True)
        cell = ws.cell(row=row, column=2, value=formula)
        cell.fill = SUMMARY_FILL
        if "Completion" in label or "Load vs" in label:
            cell.number_format = "0%"

    ws.cell(
        row=layout["summary_last_row"] + 1,
        column=1,
        value="Tip: compare Load vs capacity before sprint start; velocity feeds the Velocity sheet.",
    )
    ws.merge_cells(
        start_row=layout["summary_last_row"] + 1,
        start_column=1,
        end_row=layout["summary_last_row"] + 1,
        end_column=8,
    )

    # Data validation
    sp_list = list_formula([str(x) for x in cfg["story_point_scale"]])
    type_list = list_formula(cfg["work_item_types"])
    status_list = list_formula(cfg["statuses"])
    assignee_list = list_formula(team)

    dv_sp = DataValidation(type="list", formula1=sp_list, allow_blank=True)
    dv_type = DataValidation(type="list", formula1=type_list, allow_blank=True)
    dv_status = DataValidation(type="list", formula1=status_list, allow_blank=True)
    dv_assignee = DataValidation(type="list", formula1=assignee_list, allow_blank=True)

    if bl >= bf:
        col_d = get_column_letter(BACKLOG_SP)
        col_c = get_column_letter(BACKLOG_TYPE)
        col_f = get_column_letter(BACKLOG_STATUS)
        col_e = get_column_letter(BACKLOG_ASSIGNEE)
        dv_sp.add(f"{col_d}{bf}:{col_d}{bl}")
        dv_type.add(f"{col_c}{bf}:{col_c}{bl}")
        dv_status.add(f"{col_f}{bf}:{col_f}{bl}")
        dv_assignee.add(f"{col_e}{bf}:{col_e}{bl}")
        for dv in (dv_sp, dv_type, dv_status, dv_assignee):
            ws.add_data_validation(dv)

    # Conditional formatting: Done rows green tint
    title_col = get_column_letter(BACKLOG_TITLE)
    ws.conditional_formatting.add(
        f"{title_col}{bf}:{title_col}{bl}",
        FormulaRule(
            formula=[f'$F{bf}="Done"'],
            fill=DONE_FILL,
            stopIfTrue=True,
        ),
    )

    load_row = layout["summary_first_row"] + 6
    ws.conditional_formatting.add(
        f"B{load_row}",
        ColorScaleRule(
            start_type="num",
            start_value=0,
            start_color="63BE7B",
            mid_type="num",
            mid_value=1,
            mid_color="FFEB84",
            end_type="num",
            end_value=1.15,
            end_color="F8696B",
        ),
    )

    apply_table_borders(ws, CAPACITY_HEADER_ROW, tr, COL_MEMBER, COL_CAPACITY)
    apply_table_borders(ws, bh, bl, 1, BACKLOG_DONE_SP)
    apply_table_borders(
        ws,
        layout["summary_first_row"],
        layout["summary_first_row"] + len(summary_rows) - 1,
        1,
        2,
    )

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["H"].width = 10

    ws.freeze_panes = ws.cell(row=bf, column=BACKLOG_TITLE).coordinate

    return {
        "sheet_name": sheet_name,
        "start": start,
        "end": end,
        "goal": goal,
        "committed_cell": f"'{sheet_name}'!$B${layout['summary_first_row']}",
        "completed_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 1}",
        "velocity_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 2}",
        "carryover_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 3}",
        "completion_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 4}",
        "capacity_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 5}",
        "load_cell": f"'{sheet_name}'!$B${layout['summary_first_row'] + 6}",
        "working_days_cell": f"'{sheet_name}'!$B${META_WORKING_DAYS_ROW}",
    }


def build_velocity_sheet(ws, sprint_meta: list[dict[str, Any]], cfg: dict[str, Any]):
    ws.title = "Velocity"

    ws.cell(row=1, column=1, value="Velocity & sprint history")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.cell(
        row=2,
        column=1,
        value="Metrics pull from each sprint tab. Update backlog status to Done to refresh completed points.",
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
        ws.cell(row=row, column=2, value=meta["start"])
        ws.cell(row=row, column=2).number_format = "yyyy-mm-dd"
        ws.cell(row=row, column=3, value=meta["end"])
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
    for entry in cfg["sprints"]:
        if not isinstance(entry, dict):
            raise ValueError('Each "sprints" item must be an object.')
        for key in ("sheet_name", "start", "end"):
            if key not in entry:
                raise ValueError(f'Sprint entry missing "{key}".')
        ws = wb.create_sheet()
        meta = build_sprint_sheet(ws, entry, cfg)
        sprint_meta.append(meta)

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
