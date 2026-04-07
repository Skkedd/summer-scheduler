from collections import defaultdict
from typing import Dict, List, Tuple

from data_loader import load_progress, load_rooms, load_settings, load_staffing
from exporter import export_daily_schedule, export_summary, export_task_summary
from scheduler import (
    apply_progress_to_tasks,
    build_task_items,
    run_scheduler,
    summarize_hours_by_phase,
    summarize_hours_by_school,
    summarize_remaining_by_phase,
    summarize_remaining_by_school,
)


TaskKey = Tuple[str, str, str, str, str]


def task_key(task) -> TaskKey:
    return (
        task.school_name,
        task.building_name,
        task.zone_name,
        task.room_name,
        task.phase_name,
    )


def build_remaining_map(tasks) -> Dict[TaskKey, float]:
    remaining: Dict[TaskKey, float] = {}
    for task in tasks:
        remaining[task_key(task)] = task.remaining_hours
    return remaining


def build_task_lookup(tasks) -> Dict[TaskKey, object]:
    lookup: Dict[TaskKey, object] = {}
    for task in tasks:
        lookup[task_key(task)] = task
    return lookup


def school_order_map(schools) -> Dict[str, int]:
    return {school.name: school.order for school in schools}


def parse_active_school_labels(active_school_name: str) -> List[str]:
    if not active_school_name:
        return []

    if " | " not in active_school_name and ": " not in active_school_name:
        return [active_school_name]

    parsed: List[str] = []
    for part in active_school_name.split(" | "):
        if ": " in part:
            parsed.append(part.split(": ", 1)[1].strip())
        else:
            parsed.append(part.strip())

    return [item for item in parsed if item]


def school_sequence_from_result(result) -> List[str]:
    sequence: List[str] = []
    seen = set()

    for day in result.days:
        parsed_active = parse_active_school_labels(day.active_school_name)
        for school_name in parsed_active:
            if school_name not in seen:
                seen.add(school_name)
                sequence.append(school_name)

        for item in day.work_log:
            if item.school_name == "MULTI-SCHOOL":
                continue
            if item.school_name not in seen:
                seen.add(item.school_name)
                sequence.append(item.school_name)

    return sequence


def build_start_of_run_snapshot(rooms, settings, progress_entries):
    starting_tasks = build_task_items(rooms, settings)
    completed_before_run = apply_progress_to_tasks(starting_tasks, progress_entries)

    total_by_school = summarize_hours_by_school(starting_tasks)
    remaining_by_school = summarize_remaining_by_school(starting_tasks)

    blocked_by_school: Dict[str, float] = defaultdict(float)
    available_now_by_school: Dict[str, float] = defaultdict(float)

    blocked_tasks = []
    available_now_tasks = []

    for task in starting_tasks:
        if task.remaining_hours <= 0:
            continue

        if task.available_day > settings.current_day:
            blocked_by_school[task.school_name] += task.remaining_hours
            blocked_tasks.append(task)
        else:
            available_now_by_school[task.school_name] += task.remaining_hours
            available_now_tasks.append(task)

    return {
        "tasks": starting_tasks,
        "completed_before_run": completed_before_run,
        "total_by_school": total_by_school,
        "remaining_by_school": remaining_by_school,
        "blocked_by_school": dict(sorted(blocked_by_school.items())),
        "available_now_by_school": dict(sorted(available_now_by_school.items())),
        "blocked_tasks": blocked_tasks,
        "available_now_tasks": available_now_tasks,
    }


def print_schedule_capacity_breakdown(settings) -> None:
    print("\n--- DAILY STAFF TIME MODEL ---")
    print(f"Scheduled shift per employee: {settings.scheduled_shift_hours_per_day:.2f} hrs")
    print(f"Lunch per employee: {settings.lunch_hours_per_day:.2f} hrs")
    print(f"Breaks per employee: {settings.break_hours_per_day:.2f} hrs")
    print(f"Setup/opening per employee: {settings.setup_hours_per_day:.2f} hrs")
    print(f"Cleanup/lockup per employee: {settings.cleanup_hours_per_day:.2f} hrs")
    print(
        f"Real productive time per employee: "
        f"{settings.productive_hours_per_staff_per_day:.2f} hrs"
    )

    print("\n--- DAILY TIMING TEMPLATE ---")
    print(f"Arrival / open: {settings.day_start_time}")
    print(f"Work starts: {settings.work_start_time}")
    print(f"First break: {settings.first_break_time}")
    print(f"Lunch: {settings.lunch_time}")
    print(f"Second break: {settings.second_break_time}")
    print(f"Cleanup starts: {settings.cleanup_start_time}")
    print(f"Day ends: {settings.day_end_time}")


def print_school_readiness(snapshot, schools):
    print("\n--- SCHOOL START STATUS ---")

    total_by_school = snapshot["total_by_school"]
    remaining_by_school = snapshot["remaining_by_school"]
    blocked_by_school = snapshot["blocked_by_school"]
    available_now_by_school = snapshot["available_now_by_school"]

    for school in schools:
        total_hours = total_by_school.get(school.name, 0.0)
        remaining_hours = remaining_by_school.get(school.name, 0.0)
        blocked_hours = blocked_by_school.get(school.name, 0.0)
        available_now_hours = available_now_by_school.get(school.name, 0.0)

        if remaining_hours <= 0:
            status = "Already complete before rerun"
        elif available_now_hours > 0 and blocked_hours > 0:
            status = "Active now, but some rooms blocked for later"
        elif available_now_hours > 0:
            status = "Ready to work now"
        else:
            status = "Blocked until later release"

        print(
            f"{school.order}. {school.name} | total {total_hours:.2f} hrs | "
            f"remaining {remaining_hours:.2f} hrs | available now {available_now_hours:.2f} hrs | "
            f"blocked later {blocked_hours:.2f} hrs | {status}"
        )


def print_blocked_rooms(snapshot):
    print("\n--- BLOCKED / LATER-RELEASE ROOMS AT START ---")

    blocked_tasks = snapshot["blocked_tasks"]

    if not blocked_tasks:
        print("None")
        return

    blocked_tasks = sorted(
        blocked_tasks,
        key=lambda t: (
            t.school_order,
            t.school_name,
            t.available_day,
            t.building_name,
            t.zone_name,
            t.room_order,
            t.room_name,
            t.phase_name,
        ),
    )

    for task in blocked_tasks:
        location_bits = [task.building_name, task.zone_name, task.room_name]
        location_text = " | ".join(bit for bit in location_bits if bit)

        print(
            f"{task.school_name} | {location_text} | {task.phase_name} | "
            f"release day {task.available_day} | remaining {task.remaining_hours:.2f} hrs"
        )


def simulate_remaining_state_by_day(starting_tasks, result):
    remaining_map = build_remaining_map(starting_tasks)
    task_lookup = build_task_lookup(starting_tasks)

    snapshots_by_day = {}

    for day in result.days:
        snapshots_by_day[day.day] = {
            "remaining_map": dict(remaining_map),
            "task_lookup": task_lookup,
        }

        for item in day.work_log:
            if item.school_name == "MULTI-SCHOOL":
                continue

            key = (
                item.school_name,
                item.building_name,
                item.zone_name,
                item.room_name,
                item.phase_name,
            )

            if key in remaining_map:
                remaining_map[key] = max(0.0, remaining_map[key] - item.hours_done)

    return snapshots_by_day


def explain_idle_day(day, day_snapshot) -> str:
    remaining_map = day_snapshot["remaining_map"]
    task_lookup = day_snapshot["task_lookup"]

    available_general_now = 0.0
    available_carpet_now = 0.0
    blocked_later = 0.0
    blocked_schools = set()

    for key, remaining_hours in remaining_map.items():
        if remaining_hours <= 0:
            continue

        task = task_lookup[key]

        if task.available_day <= day.day:
            if task.phase_name == "Carpet":
                available_carpet_now += remaining_hours
            else:
                available_general_now += remaining_hours
        else:
            blocked_later += remaining_hours
            blocked_schools.add(task.school_name)

    if available_general_now <= 0 and available_carpet_now > 0 and day.carpet_staff <= 0:
        return (
            f"Idle because only carpet work remained available now "
            f"({available_carpet_now:.2f} hrs) and no carpet staff were reserved today."
        )

    if available_general_now <= 0 and available_carpet_now <= 0 and blocked_later > 0:
        blocked_schools_text = ", ".join(sorted(blocked_schools))
        return (
            f"Idle because no remaining rooms are released yet. "
            f"Blocked-later backlog: {blocked_later:.2f} hrs"
            + (f" at {blocked_schools_text}" if blocked_schools_text else "")
        )

    if available_general_now <= 0 and available_carpet_now <= 0 and blocked_later <= 0:
        return "Idle because all work is complete"

    return (
        f"Idle even though {available_general_now + available_carpet_now:.2f} hrs were available now. "
        f"This suggests a scheduling rule or data issue worth checking."
    )


def print_daily_schedule_with_reasons(result, starting_tasks):
    print("\n--- DAILY REFORECAST SCHEDULE ---")

    day_snapshots = simulate_remaining_state_by_day(starting_tasks, result)

    for day in result.days:
        print(
            f"\nDay {day.day} | total staff {day.effective_staff} | "
            f"general crew {day.general_staff} ({day.general_capacity:.2f} hrs) | "
            f"carpet crew {day.carpet_staff} ({day.carpet_capacity:.2f} hrs) | "
            f"total capacity {day.daily_capacity:.2f} | used {day.used_capacity:.2f} | "
            f"unused {day.unused_capacity:.2f}"
        )
        print(
            f"  Crew usage: general used {day.general_used_capacity:.2f} / unused {day.general_unused_capacity:.2f} | "
            f"carpet used {day.carpet_used_capacity:.2f} / unused {day.carpet_unused_capacity:.2f}"
        )

        if day.active_school_name:
            print(f"  Active school: {day.active_school_name}")

        if day.status_note:
            print(f"  Day note: {day.status_note}")

        if not day.work_log:
            reason = explain_idle_day(day, day_snapshots[day.day])
            print("  No work scheduled")
            print(f"  Reason: {reason}")
            continue

        for item in day.work_log:
            location_bits = [item.building_name, item.zone_name, item.room_name]
            location_text = " | ".join(bit for bit in location_bits if bit)

            if item.school_name == "MULTI-SCHOOL":
                print(
                    f"  [{item.crew_type}] {item.school_name} | {item.room_name} | "
                    f"{item.phase_name} | {item.hours_done:.2f} hrs"
                )
                continue

            print(
                f"  [{item.crew_type}] {item.school_name} | {location_text} | "
                f"{item.phase_name} | {item.hours_done:.2f} hrs"
            )


def print_school_sequence(result, schools):
    print("\n--- ACTIVE SCHOOL SEQUENCE ---")

    sequence = school_sequence_from_result(result)
    if not sequence:
        print("No schools scheduled")
        return

    order_map = school_order_map(schools)
    formatted = [f"{order_map.get(name, '?')}. {name}" for name in sequence]
    print(" -> ".join(formatted))


def print_recommendation(result):
    rec = result.recommendation

    print("\n--- RECOMMENDATION ---")
    print(f"Status: {rec.status_label}")
    print(f"Bottleneck: {rec.bottleneck_type}")
    print(
        f"Available-now backlog at start: {rec.available_backlog_hours:.2f} hrs"
    )
    print(
        f"Blocked-later backlog at start: {rec.blocked_backlog_hours:.2f} hrs"
    )
    print(f"Total backlog at start: {rec.total_backlog_hours:.2f} hrs")
    print(
        f"Productive capacity to deadline: {rec.capacity_to_deadline_hours:.2f} hrs"
    )
    print(
        f"Extra staff-days needed: {rec.extra_staff_days_needed:.2f}"
    )
    print(f"Action: {rec.recommended_action}")


def main():
    settings = load_settings("data/settings.csv")
    rooms, schools = load_rooms("data/rooms.csv")
    staffing_days = load_staffing("data/staffing.csv")
    progress_entries = load_progress("data/progress.csv")

    start_snapshot = build_start_of_run_snapshot(
        rooms=rooms,
        settings=settings,
        progress_entries=progress_entries,
    )

    result = run_scheduler(
        rooms=rooms,
        settings=settings,
        staffing_days=staffing_days,
        progress_entries=progress_entries,
    )

    print("\n--- SCHEDULE INFO ---")
    print(f"Schedule name: {settings.schedule_name}")
    print(f"Current day: {settings.current_day}")
    print(f"Target end day: {settings.target_end_day}")
    print(
        f"Real productive hours per staff per day: "
        f"{settings.productive_hours_per_staff_per_day:.2f}"
    )

    print("\n--- ACTIVE PHASES ---")
    print(f"Deep Clean: {'ON' if settings.include_deep_clean else 'OFF'}")
    print(f"Strip: {'ON' if settings.include_strip else 'OFF'}")
    print(f"Wax: {'ON' if settings.include_wax else 'OFF'}")
    print(f"Carpet: {'ON' if settings.include_carpet else 'OFF'}")
    print(f"Exterior: {'ON' if settings.include_exterior else 'OFF'}")

    print_schedule_capacity_breakdown(settings)

    print("\n--- SCHOOL ORDER ---")
    for school in schools:
        print(f"{school.order}. {school.name}")

    print_school_readiness(start_snapshot, schools)
    print_blocked_rooms(start_snapshot)
    print_school_sequence(result, schools)

    print("\n--- HOURS BY SCHOOL ---")
    by_school = summarize_hours_by_school(result.task_items)
    for school_name, hours in by_school.items():
        print(f"{school_name}: {hours:.2f} hrs")

    print("\n--- HOURS BY PHASE ---")
    by_phase = summarize_hours_by_phase(result.task_items)
    for phase_name, hours in by_phase.items():
        print(f"{phase_name}: {hours:.2f} hrs")

    print("\n--- REMAINING BY SCHOOL ---")
    remaining_by_school = summarize_remaining_by_school(result.task_items)
    if remaining_by_school:
        for school_name, hours in remaining_by_school.items():
            print(f"{school_name}: {hours:.2f} hrs")
    else:
        print("None")

    print("\n--- REMAINING BY PHASE ---")
    remaining_by_phase = summarize_remaining_by_phase(result.task_items)
    if remaining_by_phase:
        for phase_name, hours in remaining_by_phase.items():
            print(f"{phase_name}: {hours:.2f} hrs")
    else:
        print("None")

    print("\n--- REFORECAST RESULT ---")
    print(f"Total planned hours: {result.total_planned_hours:.2f}")
    print(f"Completed before rerun: {result.completed_hours_before_run:.2f}")
    print(f"Remaining at start of rerun: {result.remaining_hours_at_start:.2f}")
    print(f"Projected finish day: {result.finish_day}")
    print(f"Deadline met: {'YES' if result.met_deadline else 'NO'}")
    print(f"Remaining backlog hours: {result.remaining_backlog_hours:.2f}")

    print_recommendation(result)

    print_daily_schedule_with_reasons(
        result=result,
        starting_tasks=start_snapshot["tasks"],
    )

    daily_path = export_daily_schedule(result)
    task_path = export_task_summary(result)
    summary_path = export_summary(result)

    print("\n--- EXPORTED FILES ---")
    print(daily_path)
    print(task_path)
    print(summary_path)


if __name__ == "__main__":
    main()