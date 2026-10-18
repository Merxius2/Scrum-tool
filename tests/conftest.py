"""Load planner scripts by path (they are not installed packages)."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load_script(relative: str, module_name: str):
    path = ROOT / relative
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


holiday_planner = _load_script(
    "holiday-planner/generate_holiday_planner.py",
    "generate_holiday_planner",
)
sprint_planner = _load_script(
    "sprint-planning/generate_sprint_planner.py",
    "generate_sprint_planner",
)
