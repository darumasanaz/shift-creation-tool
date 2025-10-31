#!/usr/bin/env python3
"""Simple heuristic solver for generating monthly shift assignments.

This script is intentionally lightweight so that sample schedules can be
produced inside the repository without relying on heavy external solvers.
It understands the JSON format that powers the sample inputs bundled with
this project.  The solver focuses on building a plausible rota that respects
basic constraints such as fixed weekly days off and monthly workday targets.

The produced output mirrors the structure expected by the viewer: assignments
are stored per staff member, shortages are summarised per day, and shift codes
are converted to their Japanese names before writing the final JSON file.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Sequence

JAPANESE_WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土']

CODE_TO_NAME = {
    "EA": "早番",
    "DA": "日勤A",
    "DB": "日勤B",
    "LA": "遅番",
    "NA": "夜勤A",
    "NB": "夜勤B",
    "NC": "夜勤C",
}


@dataclass
class PersonConfig:
    identifier: str
    can_work_codes: Sequence[str]
    fixed_off: Sequence[str]
    monthly_min: int | None
    monthly_max: int | None


@dataclass
class ProblemDefinition:
    year: int
    month: int
    days: int
    weekday_of_day1: int
    people: List[PersonConfig]

    @classmethod
    def from_json(cls, data: Dict) -> "ProblemDefinition":
        people = [
            PersonConfig(
                identifier=person.get("id", ""),
                can_work_codes=list(person.get("canWork", [])),
                fixed_off=list(person.get("fixedOffWeekdays", [])),
                monthly_min=_coerce_int(person.get("monthlyMin")),
                monthly_max=_coerce_int(person.get("monthlyMax")),
            )
            for person in data.get("people", [])
        ]
        return cls(
            year=int(data.get("year")),
            month=int(data.get("month")),
            days=int(data.get("days")),
            weekday_of_day1=int(data.get("weekdayOfDay1", 0)),
            people=people,
        )

    def weekday_for_day(self, day_index: int) -> int:
        """Return weekday index (0=Sunday) for the given zero-based day index."""
        return (self.weekday_of_day1 + day_index) % 7


def _coerce_int(value) -> int | None:
    if value is None:
        return None
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return parsed


def _determine_target_days(person: PersonConfig, available_slots: int) -> int:
    """Determine how many working days to assign to *person*.

    The target favours the midpoint between the monthly minimum and maximum when
    both are provided.  When only a single bound is defined we bias towards that
    bound while ensuring the value stays inside the feasible range defined by
    ``available_slots``.
    """

    minimum = person.monthly_min if person.monthly_min is not None else 0
    maximum = person.monthly_max if person.monthly_max is not None else available_slots

    if person.monthly_max is None:
        maximum = available_slots
    else:
        maximum = min(person.monthly_max, available_slots)

    if person.monthly_min is None:
        minimum = 0
    else:
        minimum = min(person.monthly_min, available_slots)

    if maximum < minimum:
        maximum = minimum

    if person.monthly_min is not None and person.monthly_max is not None:
        target = round((person.monthly_min + person.monthly_max) / 2)
    elif person.monthly_max is not None:
        target = round(maximum * 0.75)
    elif person.monthly_min is not None:
        target = person.monthly_min
    else:
        target = round(available_slots * 0.5)

    target = max(minimum, target)
    target = min(maximum, target)
    return max(0, min(target, available_slots))


def _expand_fixed_off_indices(fixed_off: Sequence[str]) -> List[int]:
    indices: List[int] = []
    for token in fixed_off:
        if token in JAPANESE_WEEKDAYS:
            indices.append(JAPANESE_WEEKDAYS.index(token))
            continue
        try:
            indices.append(int(token) % 7)
        except (TypeError, ValueError):
            continue
    return indices


def _select_working_days(available_days: List[int], target: int) -> List[int]:
    """Pick *target* elements from ``available_days`` while keeping them spaced."""
    if target <= 0:
        return []
    if target >= len(available_days):
        return list(available_days)

    if target == 1:
        return [available_days[len(available_days) // 2]]

    step = (len(available_days) - 1) / (target - 1)
    selected: List[int] = []
    used_indices: set[int] = set()
    for i in range(target):
        raw_index = round(i * step)
        idx = min(len(available_days) - 1, max(0, int(raw_index)))
        while idx in used_indices and idx + 1 < len(available_days):
            idx += 1
        while idx in used_indices and idx > 0:
            idx -= 1
        used_indices.add(idx)
        selected.append(available_days[idx])
    selected.sort()
    return selected


def _build_assignment_for_person(problem: ProblemDefinition, person: PersonConfig) -> Dict:
    available_codes = [code for code in person.can_work_codes if code in CODE_TO_NAME]
    weekday_off_indices = set(_expand_fixed_off_indices(person.fixed_off))
    workdays: List[str] = []
    available_day_slots: List[int] = []

    for day_idx in range(problem.days):
        weekday = problem.weekday_for_day(day_idx)
        if weekday in weekday_off_indices:
            workdays.append("休み")
        else:
            workdays.append("__PENDING__")
            available_day_slots.append(day_idx)

    target_days = _determine_target_days(person, len(available_day_slots))
    selected_slots = set(_select_working_days(available_day_slots, target_days))

    rotation_index = 0
    for day_idx in range(problem.days):
        if workdays[day_idx] == "休み":
            continue
        if day_idx not in selected_slots or not available_codes:
            workdays[day_idx] = "休み"
            continue
        assigned_code = available_codes[rotation_index % len(available_codes)]
        rotation_index += 1
        workdays[day_idx] = assigned_code

    return {
        "staffId": person.identifier,
        "displayName": person.identifier,
        "shiftCodes": workdays,
    }


def _convert_codes_to_names(assignment: Dict) -> None:
    shifts = assignment.get("shiftCodes") or []
    converted: List[str] = []
    for code in shifts:
        if code == "休み":
            converted.append("休み")
        else:
            converted.append(CODE_TO_NAME.get(code, code))
    assignment["shifts"] = converted


def build_shortage_rows(days: int) -> List[Dict[str, int]]:
    rows: List[Dict[str, int]] = []
    for day in range(1, days + 1):
        rows.append({
            "day": day,
            "7-9": 0,
            "9-15": 0,
            "16-18": 0,
            "18-21": 0,
            "21-24": 0,
            "0-7": 0,
            "total": 0,
        })
    return rows


def solve(problem: ProblemDefinition) -> Dict:
    assignments = [_build_assignment_for_person(problem, person) for person in problem.people]
    for assignment in assignments:
        _convert_codes_to_names(assignment)

    return {
        "year": problem.year,
        "month": problem.month,
        "days": problem.days,
        "assignments": assignments,
        "shortageSummary": build_shortage_rows(problem.days),
        "metadata": {
            "generatedBy": "heuristic-solver",
            "notes": "Auto-generated sample schedule with Japanese shift labels.",
        },
    }


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate sample shift output from JSON input.")
    parser.add_argument("--in", dest="input_path", required=True, help="Path to input JSON file")
    parser.add_argument("--out", dest="output_path", required=True, help="Path to write output JSON")
    parser.add_argument("--time_limit", dest="time_limit", default=60, help="Unused compatibility flag")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    input_path = Path(args.input_path)
    output_path = Path(args.output_path)

    with input_path.open("r", encoding="utf-8") as fp:
        data = json.load(fp)

    problem = ProblemDefinition.from_json(data)
    output_data = solve(problem)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fp:
        json.dump(output_data, fp, ensure_ascii=False, indent=2)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
