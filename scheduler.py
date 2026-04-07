from collections import defaultdict
from typing import Dict, List

from models import (
    Room,
    School,
    ScheduleDayResult,
    ScheduleResult,
    ScheduleSettings,
    StaffingDay,
    TaskItem,
    WorkLogEntry,
)


def build_task_items(rooms: List[Room], settings: ScheduleSettings) -> List[TaskItem]:
    task_items: List[TaskItem] = []

    for room in sorted(
        rooms,
        key=lambda r: (r.school_order, r.school_name, r.room_order, r.room_name),
    ):
        total_sqft = room.carpet_sqft + room.tile_sqft

        if settings.include_deep_clean and room.include_deep_clean and total_sqft > 0:
            hours = total_sqft / settings.deep_clean_rate_sqft_per_hour
            task_items.append(
                TaskItem(
                    school_name=room.school_name,
                    school_order=room.school_order,
                    building_name=room.building_name,
                    zone_name=room.zone_name,
                    room_name=room.room_name,
                    room_order=room.room_order,
                    phase_name="Deep Clean",
                    available_day=room.available_day,
                    total_hours=hours,
                    remaining_hours=hours,
                    notes=room.notes,
                )
            )

        if settings.include_strip and room.include_strip and room.tile_sqft > 0:
            hours = room.tile_sqft / settings.strip_rate_sqft_per_hour
            task_items.append(
                TaskItem(
                    school_name=room.school_name,
                    school_order=room.school_order,
                    building_name=room.building_name,
                    zone_name=room.zone_name,
                    room_name=room.room_name,
                    room_order=room.room_order,
                    phase_name="Strip",
                    available_day=room.available_day,
                    total_hours=hours,
                    remaining_hours=hours,
                    notes=room.notes,
                )
            )

        if settings.include_wax and room.include_wax and room.tile_sqft > 0:
            hours = (room.tile_sqft / settings.wax_rate_sqft_per_hour) * settings.wax_coats
            task_items.append(
                TaskItem(
                    school_name=room.school_name,
                    school_order=room.school_order,
                    building_name=room.building_name,
                    zone_name=room.zone_name,
                    room_name=room.room_name,
                    room_order=room.room_order,
                    phase_name=f"Wax ({settings.wax_coats} coats)",
                    available_day=room.available_day,
                    total_hours=hours,
                    remaining_hours=hours,
                    notes=room.notes,
                )
            )

        if settings.include_carpet and room.include_carpet and room.carpet_sqft > 0:
            hours = room.carpet_sqft / settings.carpet_rate_sqft_per_hour
            task_items.append(
                TaskItem(
                    school_name=room.school_name,
                    school_order=room.school_order,
                    building_name=room.building_name,
                    zone_name=room.zone_name,
                    room_name=room.room_name,
                    room_order=room.room_order,
                    phase_name="Carpet",
                    available_day=room.available_day,
                    total_hours=hours,
                    remaining_hours=hours,
                    notes=room.notes,
                )
            )

        if settings.include_exterior and room.include_exterior and total_sqft > 0:
            hours = total_sqft / settings.exterior_rate_sqft_per_hour
            task_items.append(
                TaskItem(
                    school_name=room.school_name,
                    school_order=room.school_order,
                    building_name=room.building_name,
                    zone_name=room.zone_name,
                    room_name=room.room_name,
                    room_order=room.room_order,
                    phase_name="Exterior",
                    available_day=room.available_day,
                    total_hours=hours,
                    remaining_hours=hours,
                    notes=room.notes,
                )
            )

    return task_items


def staffing_map(staffing_days: List[StaffingDay]) -> Dict[int, StaffingDay]:
    return {item.day: item for item in staffing_days}


def calculate_school_transition_hours(
    tasks_done_today: List[WorkLogEntry],
    settings: ScheduleSettings,
) -> float:
    if settings.transition_hours_per_school <= 0:
        return 0.0

    touched_schools = {item.school_name for item in tasks_done_today}
    school_count = len(touched_schools)

    if school_count <= 1:
        return 0.0

    return (school_count - 1) * settings.transition_hours_per_school


def run_scheduler(
    rooms: List[Room],
    schools: List[School],
    settings: ScheduleSettings,
    staffing_days: List[StaffingDay],
) -> ScheduleResult:
    tasks = build_task_items(rooms, settings)
    staffing_by_day = staffing_map(staffing_days)

    total_planned_hours = sum(task.total_hours for task in tasks)
    total_used_hours = 0.0
    day_results: List[ScheduleDayResult] = []

    current_day = 1
    max_day = max(
        settings.target_end_day,
        max((item.day for item in staffing_days), default=settings.target_end_day),
    ) + 365

    while current_day <= max_day:
        remaining_backlog = sum(task.remaining_hours for task in tasks)
        if remaining_backlog <= 0:
            break

        staff_info = staffing_by_day.get(
            current_day,
            StaffingDay(
                day=current_day,
                available_staff=0,
                carpet_staff_reserved=0,
                absences=0,
                temporary_help=0,
            ),
        )

        effective_staff = staff_info.effective_cleaning_staff()
        daily_capacity = effective_staff * settings.hours_per_staff_per_day
        remaining_capacity = daily_capacity
        work_log: List[WorkLogEntry] = []

        available_tasks = [
            task
            for task in tasks
            if task.available_day <= current_day and task.remaining_hours > 0
        ]

        available_tasks.sort(
            key=lambda t: (
                t.school_order,
                t.school_name,
                t.room_order,
                t.room_name,
                phase_sort_key(t.phase_name),
            )
        )

        for task in available_tasks:
            if remaining_capacity <= 0:
                break

            hours_done = min(task.remaining_hours, remaining_capacity)
            task.remaining_hours -= hours_done
            remaining_capacity -= hours_done
            total_used_hours += hours_done

            work_log.append(
                WorkLogEntry(
                    school_name=task.school_name,
                    room_name=task.room_name,
                    phase_name=task.phase_name,
                    hours_done=hours_done,
                )
            )

        transition_hours = calculate_school_transition_hours(work_log, settings)
        if transition_hours > 0:
            actual_transition = min(transition_hours, remaining_capacity)
            remaining_capacity -= actual_transition
            total_used_hours += actual_transition

            if actual_transition > 0:
                work_log.append(
                    WorkLogEntry(
                        school_name="MULTI-SCHOOL",
                        room_name="TRANSITION",
                        phase_name="Transition / Logistics",
                        hours_done=actual_transition,
                    )
                )

        used_capacity = daily_capacity - remaining_capacity

        day_results.append(
            ScheduleDayResult(
                day=current_day,
                effective_staff=effective_staff,
                daily_capacity=daily_capacity,
                used_capacity=used_capacity,
                unused_capacity=remaining_capacity,
                work_log=work_log,
            )
        )

        current_day += 1

    finish_day = day_results[-1].day if day_results else 0
    remaining_backlog_hours = sum(task.remaining_hours for task in tasks)
    met_deadline = remaining_backlog_hours <= 0 and finish_day <= settings.target_end_day

    return ScheduleResult(
        schedule_name=settings.schedule_name,
        target_end_day=settings.target_end_day,
        finish_day=finish_day,
        met_deadline=met_deadline,
        total_planned_hours=total_planned_hours,
        total_used_hours=total_used_hours,
        days=day_results,
        remaining_backlog_hours=remaining_backlog_hours,
        task_items=tasks,
    )


def phase_sort_key(phase_name: str) -> int:
    order = {
        "Deep Clean": 1,
        "Strip": 2,
        "Wax (1 coats)": 3,
        "Wax (2 coats)": 3,
        "Wax (3 coats)": 3,
        "Wax (4 coats)": 3,
        "Carpet": 4,
        "Exterior": 5,
        "Transition / Logistics": 6,
    }
    return order.get(phase_name, 999)


def summarize_hours_by_school(tasks: List[TaskItem]) -> Dict[str, float]:
    totals: Dict[str, float] = defaultdict(float)
    for task in tasks:
        totals[task.school_name] += task.total_hours
    return dict(sorted(totals.items()))


def summarize_hours_by_phase(tasks: List[TaskItem]) -> Dict[str, float]:
    totals: Dict[str, float] = defaultdict(float)
    for task in tasks:
        totals[task.phase_name] += task.total_hours
    return dict(sorted(totals.items()))