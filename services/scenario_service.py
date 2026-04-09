from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from typing import Dict, List

from models import ProgressEntry, Room, ScheduleSettings, StaffingDay
from scheduler import run_scheduler


@dataclass
class ScenarioInput:
    settings: ScheduleSettings
    rooms: List[Room]
    progress_entries: List[ProgressEntry] = field(default_factory=list)

    cleaning_staff_by_day: Dict[int, int] = field(default_factory=dict)
    carpet_staff_by_day: Dict[int, int] = field(default_factory=dict)
    outside_help_by_day: Dict[int, int] = field(default_factory=dict)
    absences_by_day: Dict[int, int] = field(default_factory=dict)


def _build_staffing_days(scenario: ScenarioInput) -> List[StaffingDay]:
    all_days = set()
    all_days.update(scenario.cleaning_staff_by_day.keys())
    all_days.update(scenario.carpet_staff_by_day.keys())
    all_days.update(scenario.outside_help_by_day.keys())
    all_days.update(scenario.absences_by_day.keys())

    start_day = scenario.settings.current_day
    end_day = scenario.settings.target_end_day

    if not all_days:
        all_days.update(range(start_day, end_day + 1))
    else:
        min_day = min(min(all_days), start_day)
        max_day = max(max(all_days), end_day)
        all_days.update(range(min_day, max_day + 1))

    staffing_days: List[StaffingDay] = []

    for day in sorted(all_days):
        cleaning_staff = scenario.cleaning_staff_by_day.get(day, 0)
        carpet_staff = scenario.carpet_staff_by_day.get(day, 0)
        outside_help = scenario.outside_help_by_day.get(day, 0)
        absences = scenario.absences_by_day.get(day, 0)

        staffing_days.append(
            StaffingDay(
                day=day,
                available_staff=cleaning_staff + carpet_staff,
                carpet_staff_reserved=carpet_staff,
                absences=absences,
                temporary_help=outside_help,
            )
        )

    return staffing_days


def run_scenario(scenario: ScenarioInput):
    """
    Single entry point for UI -> service -> engine.

    UI should build a ScenarioInput and call this function.
    The service layer translates UI-shaped scenario data into
    engine-shaped inputs, then calls the scheduler engine.
    """
    settings = deepcopy(scenario.settings)
    staffing_days = _build_staffing_days(scenario)

    for d in staffing_days[:5]:
        print(f"DAY {d.day}: total={d.available_staff}, carpet={d.carpet_staff_reserved}, absences={d.absences}, outside={d.temporary_help}")

    return run_scheduler(
        rooms=scenario.rooms,
        settings=settings,
        staffing_days=staffing_days,
        progress_entries=scenario.progress_entries,
    )