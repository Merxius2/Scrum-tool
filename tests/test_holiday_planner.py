"""Unit tests for holiday-planner/generate_holiday_planner.py."""

from __future__ import annotations

import json
import random
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from conftest import holiday_planner as hp


def test_assign_engineer_colors_respects_seed():
    rng1 = random.Random(42)
    rng2 = random.Random(42)
    assert hp.assign_engineer_colors(5, rng1) == hp.assign_engineer_colors(5, rng2)
    colors = hp.assign_engineer_colors(3, random.Random(1))
    assert len(colors) == 3
    assert len(set(colors)) == 3


def test_contrasting_font_color():
    assert hp.contrasting_font_color("FFFFFF") == "1F1F1F"
    assert hp.contrasting_font_color("1A1A1A") == "FFFFFF"


def test_parse_day_header():
    assert hp.parse_day_header("Mon 06-01-2027") == date(2027, 1, 6)
    assert hp.parse_day_header(None) is None
    assert hp.parse_day_header("bad") is None


def test_normalize_year_entry():
    assert hp.normalize_year_entry(2028) == {"sheet_name": "2028", "year": 2028}
    entry = {"sheet_name": "Custom", "year": 2029}
    assert hp.normalize_year_entry(entry) == entry


def test_resolve_period_full_year():
    entry = {"sheet_name": "2027", "year": 2027}
    name, start, end = hp.resolve_period(entry, date(2026, 10, 18))
    assert name == "2027"
    assert start == date(2027, 1, 1)
    assert end == date(2027, 12, 31)


def test_update_team_members_in_config(tmp_path: Path):
    cfg_path = tmp_path / "holiday_planner.config.json"
    cfg_path.write_text(
        json.dumps(
            {
                "output_file": "out.xlsx",
                "team_members": ["Alice", "Bob"],
                "years": [2027],
            }
        ),
        encoding="utf-8",
    )
    members = hp.update_team_members_in_config(cfg_path, add="Carol", remove=None)
    assert members == ["Alice", "Bob", "Carol"]
    hp.update_team_members_in_config(cfg_path, add=None, remove="Bob")
    data = json.loads(cfg_path.read_text(encoding="utf-8"))
    assert data["team_members"] == ["Alice", "Carol"]

    with pytest.raises(ValueError, match="already on the team"):
        hp.update_team_members_in_config(cfg_path, add="Alice", remove=None)

    solo_path = tmp_path / "solo.json"
    solo_path.write_text(
        json.dumps({"team_members": ["Only"], "years": [2027]}),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="last team member"):
        hp.update_team_members_in_config(solo_path, add=None, remove="Only")


def test_backup_workbook_if_exists_creates_backups_dir(tmp_path: Path):
    workbook = tmp_path / "Team-Holiday-Planner.xlsx"
    workbook.write_bytes(b"not-real-xlsx-but-copied")
    backup = hp.backup_workbook_if_exists(workbook)
    assert backup is not None
    assert backup.parent == tmp_path / hp.BACKUPS_DIR_NAME
    assert backup.is_file()
    assert hp.backup_workbook_if_exists(tmp_path / "missing.xlsx") is None


def test_generate_and_mutate_team(tmp_path: Path):
    cfg_path = tmp_path / "holiday_planner.config.json"
    cfg_path.write_text(
        json.dumps(
            {
                "output_file": "Team-Holiday-Planner.xlsx",
                "holiday_marker": "H",
                "color_seed": 99,
                "team_members": ["Engineer 1", "Engineer 2"],
                "years": [{"sheet_name": "2027", "year": 2027}],
            }
        ),
        encoding="utf-8",
    )
    out, _ = hp.generate(cfg_path, force=False, today=date(2026, 1, 5))
    assert out.is_file()

    wb = load_workbook(out)
    ws = wb["2027"]
    ws.cell(row=hp.FIRST_DATA_ROW, column=hp.FIRST_DAY_COL, value="H")
    wb.save(out)

    out2, backup, _action = hp.mutate_holiday_team(
        cfg_path, add="Engineer 3", remove=None, workbook=None
    )
    assert backup is not None
    assert (tmp_path / hp.BACKUPS_DIR_NAME).is_dir()
    assert "Engineer 3" in json.loads(cfg_path.read_text())["team_members"]

    wb2 = load_workbook(out2)
    ws2 = wb2["2027"]
    assert ws2.cell(row=hp.FIRST_DATA_ROW, column=hp.FIRST_DAY_COL).value == "H"
    names = []
    row = hp.FIRST_DATA_ROW
    while ws2.cell(row=row, column=hp.NAME_COL).value != "Total on holiday":
        names.append(ws2.cell(row=row, column=hp.NAME_COL).value)
        row += 1
    assert names == ["Engineer 1", "Engineer 2", "Engineer 3"]
