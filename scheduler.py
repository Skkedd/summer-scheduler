from collections import defaultdict
from typing import Dict, List, Set, Tuple

from models import (
    ProgressEntry,
    RecommendationSummary,
    Room,
    ScheduleDayResult,
    ScheduleResult,
    ScheduleSettings,
    StaffingDay,
    TaskItem,
    WorkLogEntry,
)

EPSILON = 0.005


def is_effectively_zero(value: float) -> bool:
    return abs(value) < EPSILON


def clean_hours(value: float) -> float:
    return 0.0 if is_effectively_zero(value) else value


def sentence_case(text: str) -> str:
    if not text:
        return text
    return text[0].upper() + text[1:]


def is_carpet_task(task: TaskItem) -> bool:
    return task.phase_name == "Carpet"


def is_non_carpet_task(task: TaskItem) -> bool:
    return task.phase_name != "Carpet"


def is_wax_task(task: TaskItem) -> bool:
    return task.phase_name.startswith("Wax")


def is_strip_wax_task(task: TaskItem) -> bool:
    return task.phase_name == "Strip" or is_wax_task(task)


def room_key(task: TaskItem) -> Tuple[str, str, str, str]:
    return (
        task.school_name,
        task.building_name,
        task.zone_name,
        task.room_name,
    )


def build_task_items(rooms: List[Room], settings: ScheduleSettings) -> List[TaskItem]:
    task_items: List[TaskItem] = []

    for room in sorted(
        rooms,
        key=lambda r: (
            r.school_order,
            r.school_name,
            r.room_order,
            r.building_name,
            r.zone_name,
            r.room_name,
        ),
    ):
        total_sqft = room.total_room_sqft
        strip_wax_sqft = room.tile_strip_wax_sqft
        carpet_sqft = room.carpet_sqft

        if settings.include_deep_clean and room.include_deep_clean and total_sqft > 0:
            room_use = (getattr(room, "room_use", "") or "").strip().lower()
            restroom_rate = getattr(
                settings,
                "restroom_deep_clean_rate_sqft_per_hour",
                settings.deep_clean_rate_sqft_per_hour,
            )
            deep_clean_rate = (
                restroom_rate
                if room_use in {"restroom", "bathroom", "small restroom", "large restroom"}
                else settings.deep_clean_rate_sqft_per_hour
            )
            hours = total_sqft / deep_clean_rate if deep_clean_rate > 0 else 0.0
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

        if settings.include_strip and room.include_strip and strip_wax_sqft > 0:
            hours = strip_wax_sqft / settings.strip_rate_sqft_per_hour
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

        if settings.include_wax and room.include_wax and strip_wax_sqft > 0:
            hours = (strip_wax_sqft / settings.wax_rate_sqft_per_hour) * settings.wax_coats
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

        if settings.include_carpet and room.include_carpet and carpet_sqft > 0:
            hours = carpet_sqft / settings.carpet_rate_sqft_per_hour
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
        "Site Transition": 6,
    }
    return order.get(phase_name, 999)


def task_sort_key(task: TaskItem) -> Tuple[int, str, int, int, str, str, str]:
    return (
        task.school_order,
        task.school_name,
        task.room_order,
        phase_sort_key(task.phase_name),
        task.building_name,
        task.zone_name,
        task.room_name,
    )


def progress_identity_key(entry: ProgressEntry) -> Tuple[str, str, str, str, str]:
    return (
        entry.school_name,
        entry.building_name,
        entry.zone_name,
        entry.room_name,
        entry.phase_name,
    )


def progress_legacy_key(entry: ProgressEntry) -> Tuple[str, str, str]:
    return (
        entry.school_name,
        entry.room_name,
        entry.phase_name,
    )


def calculate_school_transition_hours(
    tasks_done_today: List[WorkLogEntry],
    settings: ScheduleSettings,
) -> float:
    if settings.transition_hours_per_school <= 0:
        return 0.0

    touched_schools = {
        item.school_name
        for item in tasks_done_today
        if item.school_name != "MULTI-SCHOOL"
    }

    if len(touched_schools) <= 1:
        return 0.0

    return (len(touched_schools) - 1) * settings.transition_hours_per_school


def apply_progress_to_tasks(
    tasks: List[TaskItem],
    progress_entries: List[ProgressEntry],
) -> float:
    progress_map_precise: Dict[Tuple[str, str, str, str, str], float] = defaultdict(float)
    progress_map_legacy: Dict[Tuple[str, str, str], float] = defaultdict(float)

    for entry in progress_entries:
        if entry.building_name.strip() or entry.zone_name.strip():
            progress_map_precise[progress_identity_key(entry)] += entry.hours_completed
        else:
            progress_map_legacy[progress_legacy_key(entry)] += entry.hours_completed

    completed_hours_total = 0.0

    for task in tasks:
        precise_key = task.identity_key()
        legacy_key = task.legacy_progress_key()

        if precise_key in progress_map_precise:
            completed = min(progress_map_precise[precise_key], task.total_hours)
        else:
            completed = min(progress_map_legacy.get(legacy_key, 0.0), task.total_hours)

        completed = clean_hours(completed)
        task.remaining_hours = clean_hours(max(0.0, task.total_hours - completed))
        completed_hours_total = clean_hours(completed_hours_total + completed)

    return clean_hours(completed_hours_total)


def backlog_split(tasks: List[TaskItem], current_day: int) -> Tuple[float, float]:
    available = 0.0
    blocked = 0.0

    for task in tasks:
        if task.remaining_hours <= 0:
            continue
        if task.available_day <= current_day:
            available += task.remaining_hours
        else:
            blocked += task.remaining_hours

    return clean_hours(available), clean_hours(blocked)


def remaining_capacity_through_deadline(
    staffing_days: List[StaffingDay],
    settings: ScheduleSettings,
) -> float:
    staffing_by_day = staffing_map(staffing_days)
    total = 0.0

    for day in range(settings.current_day, settings.target_end_day + 1):
        staff_info = staffing_by_day.get(
            day,
            StaffingDay(
                day=day,
                available_staff=0,
                carpet_staff_reserved=0,
                absences=0,
                temporary_help=0,
            ),
        )
        total += (
            staff_info.effective_cleaning_staff()
            * settings.productive_hours_per_staff_per_day
        )

    return clean_hours(total)


def build_recommendation(
    start_available_backlog_hours: float,
    start_blocked_backlog_hours: float,
    capacity_to_deadline_hours: float,
    finish_day: int,
    settings: ScheduleSettings,
    remaining_backlog_hours: float,
) -> RecommendationSummary:
    total_backlog_hours = clean_hours(
        start_available_backlog_hours + start_blocked_backlog_hours
    )

    if total_backlog_hours <= 0:
        return RecommendationSummary(
            status_label="COMPLETE",
            bottleneck_type="None",
            available_backlog_hours=0.0,
            blocked_backlog_hours=0.0,
            extra_staff_days_needed=0.0,
            recommended_action="All scheduled work was already complete before this rerun.",
            based_on_start_of_run=True,
            capacity_to_deadline_hours=capacity_to_deadline_hours,
            total_backlog_hours=0.0,
        )

    if remaining_backlog_hours > 0:
        shortfall = clean_hours(max(remaining_backlog_hours, total_backlog_hours - capacity_to_deadline_hours))
        return RecommendationSummary(
            status_label="BEHIND",
            bottleneck_type=(
                "Mixed (Access + Capacity)"
                if start_blocked_backlog_hours > 0
                else "Staffing Capacity"
            ),
            available_backlog_hours=start_available_backlog_hours,
            blocked_backlog_hours=start_blocked_backlog_hours,
            extra_staff_days_needed=clean_hours(shortfall / settings.productive_hours_per_staff_per_day),
            recommended_action=(
                f"Work remains unfinished after Day {settings.target_end_day}. "
                f"Add staff, reserve carpet crew capacity, reduce scope, or release blocked rooms earlier."
            ),
            based_on_start_of_run=True,
            capacity_to_deadline_hours=capacity_to_deadline_hours,
            total_backlog_hours=total_backlog_hours,
        )

    if start_blocked_backlog_hours > 0 and start_available_backlog_hours == 0:
        return RecommendationSummary(
            status_label="WAITING ON ACCESS",
            bottleneck_type="Room Availability",
            available_backlog_hours=start_available_backlog_hours,
            blocked_backlog_hours=start_blocked_backlog_hours,
            extra_staff_days_needed=0.0,
            recommended_action=(
                "All remaining work at the start of the run was blocked by room availability. "
                "Extra staffing will not help until rooms release."
            ),
            based_on_start_of_run=True,
            capacity_to_deadline_hours=capacity_to_deadline_hours,
            total_backlog_hours=total_backlog_hours,
        )

    if finish_day > settings.target_end_day:
        shortfall = clean_hours(max(0.0, total_backlog_hours - capacity_to_deadline_hours))
        extra_staff_days_needed = clean_hours(
            shortfall / settings.productive_hours_per_staff_per_day
        )
        return RecommendationSummary(
            status_label="BEHIND",
            bottleneck_type=(
                "Mixed (Access + Capacity)"
                if start_blocked_backlog_hours > 0
                else "Staffing Capacity"
            ),
            available_backlog_hours=start_available_backlog_hours,
            blocked_backlog_hours=start_blocked_backlog_hours,
            extra_staff_days_needed=extra_staff_days_needed,
            recommended_action=(
                f"Start-of-run backlog exceeds productive capacity to Day {settings.target_end_day}. "
                f"Add about {extra_staff_days_needed:.2f} extra staff-days, reduce scope, "
                f"or release blocked rooms earlier."
            ),
            based_on_start_of_run=True,
            capacity_to_deadline_hours=capacity_to_deadline_hours,
            total_backlog_hours=total_backlog_hours,
        )

    if start_blocked_backlog_hours > 0:
        return RecommendationSummary(
            status_label="ON TRACK WITH ACCESS CONSTRAINTS",
            bottleneck_type="Room Availability",
            available_backlog_hours=start_available_backlog_hours,
            blocked_backlog_hours=start_blocked_backlog_hours,
            extra_staff_days_needed=0.0,
            recommended_action=(
                "Schedule is on track, but later room release is controlling part of the finish. "
                "Earlier access would create a smoother schedule faster than adding staff."
            ),
            based_on_start_of_run=True,
            capacity_to_deadline_hours=capacity_to_deadline_hours,
            total_backlog_hours=total_backlog_hours,
        )

    return RecommendationSummary(
        status_label="ON TRACK",
        bottleneck_type="None",
        available_backlog_hours=start_available_backlog_hours,
        blocked_backlog_hours=start_blocked_backlog_hours,
        extra_staff_days_needed=0.0,
        recommended_action=(
            "Start-of-run backlog fits within available productive capacity before the deadline."
        ),
        based_on_start_of_run=True,
        capacity_to_deadline_hours=capacity_to_deadline_hours,
        total_backlog_hours=total_backlog_hours,
    )


def group_tasks_by_school(tasks: List[TaskItem]) -> Dict[str, List[TaskItem]]:
    grouped: Dict[str, List[TaskItem]] = defaultdict(list)
    for task in tasks:
        grouped[task.school_name].append(task)
    for school_name in grouped:
        grouped[school_name].sort(key=task_sort_key)
    return grouped


def school_sort_key_from_tasks(tasks: List[TaskItem]) -> List[Tuple[int, str]]:
    seen: Dict[str, int] = {}
    for task in tasks:
        if task.school_name not in seen:
            seen[task.school_name] = task.school_order
    return sorted((order, name) for name, order in seen.items())


def build_room_task_map(tasks: List[TaskItem]) -> Dict[Tuple[str, str, str, str], List[TaskItem]]:
    grouped: Dict[Tuple[str, str, str, str], List[TaskItem]] = defaultdict(list)
    for task in tasks:
        grouped[room_key(task)].append(task)
    return grouped


def non_carpet_done_for_room(room_tasks: List[TaskItem]) -> bool:
    for task in room_tasks:
        if is_non_carpet_task(task) and task.remaining_hours > 0:
            return False
    return True


def get_general_school(
    tasks: List[TaskItem],
    deferred_general_keys: Set[Tuple[str, str, str, str, str]],
) -> str:
    for _, school_name in school_sort_key_from_tasks(tasks):
        for task in tasks:
            if (
                task.school_name == school_name
                and is_non_carpet_task(task)
                and task.remaining_hours > 0
                and task.identity_key() not in deferred_general_keys
            ):
                return school_name
    return ""


def get_general_deferred_school(
    tasks: List[TaskItem],
    deferred_general_keys: Set[Tuple[str, str, str, str, str]],
    current_day: int,
) -> str:
    for _, school_name in school_sort_key_from_tasks(tasks):
        for task in tasks:
            if (
                task.school_name == school_name
                and is_non_carpet_task(task)
                and task.remaining_hours > 0
                and task.identity_key() in deferred_general_keys
                and task.available_day <= current_day
            ):
                return school_name
    return ""


def get_carpet_school(
    tasks: List[TaskItem],
    room_tasks_map: Dict[Tuple[str, str, str, str], List[TaskItem]],
    deferred_carpet_keys: Set[Tuple[str, str, str, str, str]],
    current_day: int,
) -> str:
    for _, school_name in school_sort_key_from_tasks(tasks):
        for task in tasks:
            if (
                task.school_name == school_name
                and is_carpet_task(task)
                and task.remaining_hours > 0
                and task.identity_key() not in deferred_carpet_keys
                and task.available_day <= current_day
                and non_carpet_done_for_room(room_tasks_map[room_key(task)])
            ):
                return school_name
    return ""


def get_carpet_deferred_school(
    tasks: List[TaskItem],
    room_tasks_map: Dict[Tuple[str, str, str, str], List[TaskItem]],
    deferred_carpet_keys: Set[Tuple[str, str, str, str, str]],
    current_day: int,
) -> str:
    for _, school_name in school_sort_key_from_tasks(tasks):
        for task in tasks:
            if (
                task.school_name == school_name
                and is_carpet_task(task)
                and task.remaining_hours > 0
                and task.identity_key() in deferred_carpet_keys
                and task.available_day <= current_day
                and non_carpet_done_for_room(room_tasks_map[room_key(task)])
            ):
                return school_name
    return ""


def split_general_school_tasks_for_day(
    school_tasks: List[TaskItem],
    current_day: int,
    deferred_general_keys: Set[Tuple[str, str, str, str, str]],
    use_deferred_only: bool,
) -> Tuple[List[TaskItem], List[TaskItem]]:
    available_today: List[TaskItem] = []
    blocked_today: List[TaskItem] = []

    for task in school_tasks:
        if not is_non_carpet_task(task) or task.remaining_hours <= 0:
            continue

        key = task.identity_key()
        is_deferred = key in deferred_general_keys

        if use_deferred_only and not is_deferred:
            continue
        if not use_deferred_only and is_deferred:
            continue

        if task.available_day <= current_day:
            available_today.append(task)
        else:
            blocked_today.append(task)

    # Site-level phase order:
    # 1) Deep clean every currently available room at this site.
    # 2) Then strip currently available strip/wax rooms.
    # 3) Then wax currently available strip/wax rooms.
    # 4) Then other non-carpet tasks like exterior.
    # Blocked future rooms do not stop work that is available now.
    if any(t.phase_name == "Deep Clean" for t in available_today):
        available_today = [t for t in available_today if t.phase_name == "Deep Clean"]
    elif any(t.phase_name == "Strip" for t in available_today):
        available_today = [t for t in available_today if t.phase_name == "Strip"]
    elif any(is_wax_task(t) for t in available_today):
        available_today = [t for t in available_today if is_wax_task(t)]

    available_today.sort(key=task_sort_key)
    return available_today, blocked_today


def split_carpet_school_tasks_for_day(
    school_tasks: List[TaskItem],
    room_tasks_map: Dict[Tuple[str, str, str, str], List[TaskItem]],
    current_day: int,
    deferred_carpet_keys: Set[Tuple[str, str, str, str, str]],
    use_deferred_only: bool,
) -> Tuple[List[TaskItem], List[TaskItem], List[TaskItem]]:
    ready_today: List[TaskItem] = []
    blocked_today: List[TaskItem] = []
    not_ready_today: List[TaskItem] = []

    for task in school_tasks:
        if not is_carpet_task(task) or task.remaining_hours <= 0:
            continue

        key = task.identity_key()
        is_deferred = key in deferred_carpet_keys

        if use_deferred_only and not is_deferred:
            continue
        if not use_deferred_only and is_deferred:
            continue

        if task.available_day > current_day:
            blocked_today.append(task)
            continue

        if non_carpet_done_for_room(room_tasks_map[room_key(task)]):
            ready_today.append(task)
        else:
            not_ready_today.append(task)

    return ready_today, blocked_today, not_ready_today


def schedule_task_list(
    tasks_to_run: List[TaskItem],
    crew_capacity: float,
    crew_type: str,
    minimum_partial_strip_wax_hours: float = 0.5,
) -> Tuple[List[WorkLogEntry], float]:
    remaining_capacity = clean_hours(crew_capacity)
    work_log: List[WorkLogEntry] = []
    used_capacity = 0.0

    for task in tasks_to_run:
        if remaining_capacity <= 0:
            break

        # Strip/wax can be partially carried over for big areas like gyms,
        # but do not start a tiny end-of-day sliver like 5 minutes.
        if is_strip_wax_task(task) and task.remaining_hours > remaining_capacity:
            if remaining_capacity < minimum_partial_strip_wax_hours:
                break

        hours_done = clean_hours(min(task.remaining_hours, remaining_capacity))
        if hours_done <= 0:
            continue

        task.remaining_hours = clean_hours(task.remaining_hours - hours_done)
        remaining_capacity = clean_hours(remaining_capacity - hours_done)
        used_capacity = clean_hours(used_capacity + hours_done)

        work_log.append(
            WorkLogEntry(
                school_name=task.school_name,
                building_name=task.building_name,
                zone_name=task.zone_name,
                room_name=task.room_name,
                phase_name=task.phase_name,
                hours_done=hours_done,
                available_day=task.available_day,
                note=task.notes,
                crew_type=crew_type,
            )
        )

    return work_log, used_capacity


def build_day_status_note(
    general_school: str,
    carpet_school: str,
    carpet_assist_school: str,
    general_available_today: List[TaskItem],
    general_blocked_today: List[TaskItem],
    carpet_ready_today: List[TaskItem],
    carpet_blocked_today: List[TaskItem],
    carpet_not_ready_today: List[TaskItem],
    general_use_deferred_only: bool,
    carpet_use_deferred_only: bool,
    carpet_staff: int,
    carpet_used_capacity: float,
) -> str:
    parts: List[str] = []

    if general_school:
        if general_use_deferred_only and general_available_today:
            parts.append(f"general crew cleaning deferred rooms at {general_school}")
        elif general_available_today and general_blocked_today:
            parts.append(
                f"general crew working available rooms at {general_school} while blocked rooms are deferred"
            )
        elif general_available_today:
            parts.append(f"general crew working {general_school}")
        elif general_blocked_today:
            parts.append(f"general crew waiting on release at {general_school}")
        else:
            parts.append("general crew idle")
    else:
        parts.append("general crew idle")

    if carpet_staff > 0:
        if carpet_school:
            if carpet_use_deferred_only and carpet_ready_today:
                parts.append(f"carpet crew cleaning deferred carpet rooms at {carpet_school}")
            elif carpet_ready_today and (carpet_blocked_today or carpet_not_ready_today):
                parts.append(
                    f"carpet crew working ready rooms at {carpet_school} while other carpet rooms wait for release or prep"
                )
            elif carpet_ready_today:
                parts.append(f"carpet crew working {carpet_school}")
            elif carpet_blocked_today or carpet_not_ready_today:
                parts.append(
                    f"carpet crew waiting on room release or prep completion at {carpet_school}"
                )
            else:
                parts.append("carpet crew idle")
        else:
            parts.append("carpet crew idle")
    else:
        if carpet_assist_school and carpet_used_capacity > 0:
            parts.append(f"general crew helping with carpet work at {carpet_assist_school}")
        elif carpet_ready_today or carpet_blocked_today or carpet_not_ready_today or carpet_school:
            parts.append("no carpet staff reserved today")
        else:
            parts.append("no carpet staff reserved today")

    return sentence_case("; ".join(parts)) + "."


def run_scheduler(
    rooms: List[Room],
    settings: ScheduleSettings,
    staffing_days: List[StaffingDay],
    progress_entries: List[ProgressEntry],
) -> ScheduleResult:
    tasks = build_task_items(rooms, settings)
    completed_hours_before_run = apply_progress_to_tasks(tasks, progress_entries)
    remaining_hours_at_start = clean_hours(sum(task.remaining_hours for task in tasks))

    start_available_backlog_hours, start_blocked_backlog_hours = backlog_split(
        tasks, settings.current_day
    )
    capacity_to_deadline_hours = remaining_capacity_through_deadline(
        staffing_days, settings
    )

    staffing_by_day = staffing_map(staffing_days)
    total_planned_hours = clean_hours(sum(task.total_hours for task in tasks))
    total_used_hours = 0.0
    day_results: List[ScheduleDayResult] = []

    deferred_general_keys: Set[Tuple[str, str, str, str, str]] = set()
    deferred_carpet_keys: Set[Tuple[str, str, str, str, str]] = set()

    current_day = settings.current_day
    max_day = max(
        settings.target_end_day,
        max((item.day for item in staffing_days), default=settings.target_end_day),
    ) + 60

    no_progress_days = 0

    # Site transition is real crew-clock time used to gather tools, load carts/vehicles,
    # move to the next campus and stage for the next site. It intentionally does NOT
    # consume cleanup/lockup time, which remains protected for security checks.
    site_transition_clock_hours = getattr(settings, "transition_hours_per_school", 0.0)
    if site_transition_clock_hours <= 0:
        site_transition_clock_hours = 2.0

    min_clock_hours_to_start_transition = getattr(
        settings,
        "minimum_clock_hours_to_start_site_transition",
        1.0,
    )

    last_general_site = ""
    pending_transition_to_school = ""
    pending_transition_remaining_labor_hours = 0.0
    final_transition_target = "tool storage / home sites"
    final_site_transition_started = False
    final_site_transition_complete = False

    while current_day <= max_day:
        remaining_backlog = clean_hours(sum(task.remaining_hours for task in tasks))
        if remaining_backlog <= 0 and final_site_transition_complete:
            break

        tasks_by_school = group_tasks_by_school(tasks)
        room_tasks_map = build_room_task_map(tasks)

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

        general_staff = staff_info.general_crew_staff()
        carpet_staff = staff_info.carpet_crew_staff()
        effective_staff = staff_info.effective_cleaning_staff()

        general_capacity = clean_hours(
            general_staff * settings.productive_hours_per_staff_per_day
        )
        carpet_capacity = clean_hours(
            carpet_staff * settings.productive_hours_per_staff_per_day
        )
        daily_capacity = clean_hours(general_capacity + carpet_capacity)

        general_used_capacity = 0.0
        carpet_used_capacity = 0.0
        work_log: List[WorkLogEntry] = []

        general_school = get_general_school(tasks, deferred_general_keys)
        general_use_deferred_only = False
        if not general_school:
            general_school = get_general_deferred_school(tasks, deferred_general_keys, current_day)
            general_use_deferred_only = True

        carpet_school = ""
        carpet_use_deferred_only = False
        carpet_ready_today: List[TaskItem] = []
        carpet_blocked_today: List[TaskItem] = []
        carpet_not_ready_today: List[TaskItem] = []

        if carpet_capacity > 0:
            carpet_school = get_carpet_school(
                tasks,
                room_tasks_map,
                deferred_carpet_keys,
                current_day,
            )
            if not carpet_school:
                carpet_school = get_carpet_deferred_school(
                    tasks,
                    room_tasks_map,
                    deferred_carpet_keys,
                    current_day,
                )
                carpet_use_deferred_only = True
        else:
            carpet_school = get_carpet_school(
                tasks,
                room_tasks_map,
                deferred_carpet_keys,
                current_day,
            )
            if not carpet_school:
                carpet_school = get_carpet_deferred_school(
                    tasks,
                    room_tasks_map,
                    deferred_carpet_keys,
                    current_day,
                )
                carpet_use_deferred_only = True

        general_available_today: List[TaskItem] = []
        general_blocked_today: List[TaskItem] = []
        carpet_assist_school = ""
        access_downtime_warning = ""

        remaining_general_capacity = general_capacity
        minimum_partial_strip_wax_hours = getattr(
            settings,
            "minimum_partial_strip_wax_hours",
            0.5,
        )

        def transition_labor_hours_for_current_crew() -> float:
            # Transition is entered as crew-clock hours. Convert it to labor-hours so
            # it consumes the same capacity model as other full-crew work.
            crew_size = max(1, general_staff or effective_staff or 1)
            return site_transition_clock_hours * crew_size

        def min_transition_start_labor_hours() -> float:
            crew_size = max(1, general_staff or effective_staff or 1)
            return min_clock_hours_to_start_transition * crew_size

        def append_transition_work(target_school: str, labor_hours: float, remaining_after: float) -> None:
            note = (
                f"Site transition toward {target_school}. Includes gathering tools, loading equipment, "
                f"travel and staging. Cleanup/lockup time is separate and still protected."
            )
            if remaining_after > 0:
                note += f" Remaining transition after this block: {remaining_after:.2f} labor-hours."
            work_log.append(
                WorkLogEntry(
                    school_name="MULTI-SCHOOL",
                    building_name="",
                    zone_name="",
                    room_name="TRANSITION",
                    phase_name="Site Transition",
                    hours_done=labor_hours,
                    available_day=current_day,
                    note=note,
                    crew_type="General",
                )
            )

        # General crew can move through more than one site in a day, but only after
        # paying the transition cost. If a site is blocked by unavailable rooms,
        # the crew does not silently jump ahead; the day gets an access warning.
        while remaining_general_capacity > 0:
            if pending_transition_remaining_labor_hours > 0:
                used_now = clean_hours(min(pending_transition_remaining_labor_hours, remaining_general_capacity))
                pending_transition_remaining_labor_hours = clean_hours(
                    pending_transition_remaining_labor_hours - used_now
                )
                remaining_general_capacity = clean_hours(remaining_general_capacity - used_now)
                general_used_capacity = clean_hours(general_used_capacity + used_now)
                append_transition_work(
                    pending_transition_to_school,
                    used_now,
                    pending_transition_remaining_labor_hours,
                )

                if pending_transition_remaining_labor_hours > 0:
                    break

                completed_transition_target = pending_transition_to_school
                pending_transition_to_school = ""
                if completed_transition_target == final_transition_target:
                    final_site_transition_complete = True
                else:
                    last_general_site = completed_transition_target
                continue

            general_school = get_general_school(tasks, deferred_general_keys)
            general_use_deferred_only = False
            if not general_school:
                general_school = get_general_deferred_school(tasks, deferred_general_keys, current_day)
                general_use_deferred_only = True

            if not general_school:
                remaining_regular_backlog = clean_hours(sum(task.remaining_hours for task in tasks))
                if remaining_regular_backlog <= 0 and not final_site_transition_complete:
                    # All scheduled site work is complete. Still schedule a final tool-return
                    # transition so shared tools and equipment get back to their proper homes.
                    if not final_site_transition_started:
                        if remaining_general_capacity < min_transition_start_labor_hours():
                            break
                        pending_transition_to_school = final_transition_target
                        pending_transition_remaining_labor_hours = transition_labor_hours_for_current_crew()
                        final_site_transition_started = True
                        continue
                break

            # Changing sites requires a transition block first. Only start it when
            # there is at least one crew-clock hour available before cleanup.
            if last_general_site and general_school != last_general_site:
                if remaining_general_capacity < min_transition_start_labor_hours():
                    break

                pending_transition_to_school = general_school
                pending_transition_remaining_labor_hours = transition_labor_hours_for_current_crew()
                continue

            general_school_tasks = tasks_by_school[general_school]
            general_available_today, general_blocked_today = split_general_school_tasks_for_day(
                school_tasks=general_school_tasks,
                current_day=current_day,
                deferred_general_keys=deferred_general_keys,
                use_deferred_only=general_use_deferred_only,
            )

            if not general_use_deferred_only and general_blocked_today and general_available_today:
                for task in general_blocked_today:
                    deferred_general_keys.add(task.identity_key())

            if not general_available_today:
                if general_blocked_today:
                    next_release_day = min(task.available_day for task in general_blocked_today)
                    unused_clock_hours = (
                        remaining_general_capacity / max(1, general_staff or effective_staff or 1)
                    )
                    if unused_clock_hours >= 1.0:
                        access_downtime_warning = (
                            f"ACCESS DOWNTIME WARNING: {unused_clock_hours:.2f} crew-clock hours unused at "
                            f"{general_school} because rooms are restricted until Day {next_release_day}. "
                            f"Consider assigning alternate work. Check Day Outlook for details."
                        )
                break

            general_log, used_now = schedule_task_list(
                tasks_to_run=general_available_today,
                crew_capacity=remaining_general_capacity,
                crew_type="General",
                minimum_partial_strip_wax_hours=minimum_partial_strip_wax_hours,
            )

            if used_now <= 0:
                break

            general_used_capacity = clean_hours(general_used_capacity + used_now)
            remaining_general_capacity = clean_hours(remaining_general_capacity - used_now)
            work_log.extend(general_log)
            last_general_site = general_school

        if carpet_school:
            carpet_school_tasks = tasks_by_school[carpet_school]
            carpet_ready_today, carpet_blocked_today, carpet_not_ready_today = split_carpet_school_tasks_for_day(
                school_tasks=carpet_school_tasks,
                room_tasks_map=room_tasks_map,
                current_day=current_day,
                deferred_carpet_keys=deferred_carpet_keys,
                use_deferred_only=carpet_use_deferred_only,
            )

            if (
                not carpet_use_deferred_only
                and (carpet_blocked_today or carpet_not_ready_today)
                and carpet_ready_today
            ):
                for task in carpet_blocked_today + carpet_not_ready_today:
                    deferred_carpet_keys.add(task.identity_key())

            carpet_ready_today.sort(key=task_sort_key)

            if carpet_capacity > 0:
                carpet_log, carpet_used_capacity = schedule_task_list(
                    tasks_to_run=carpet_ready_today,
                    crew_capacity=carpet_capacity,
                    crew_type="Carpet",
                )
                work_log.extend(carpet_log)

        general_unused_capacity = clean_hours(general_capacity - general_used_capacity)

        general_crew_can_do_carpet = getattr(settings, "general_crew_can_do_carpet", True)

        if carpet_capacity <= 0 and general_unused_capacity > 0 and general_crew_can_do_carpet:
            room_tasks_map = build_room_task_map(tasks)

            assist_use_deferred_only = False
            assist_school = get_carpet_school(
                tasks,
                room_tasks_map,
                deferred_carpet_keys,
                current_day,
            )
            if not assist_school:
                assist_school = get_carpet_deferred_school(
                    tasks,
                    room_tasks_map,
                    deferred_carpet_keys,
                    current_day,
                )
                assist_use_deferred_only = True

            if assist_school:
                assist_school_tasks = tasks_by_school[assist_school]
                assist_ready, assist_blocked, assist_not_ready = split_carpet_school_tasks_for_day(
                    school_tasks=assist_school_tasks,
                    room_tasks_map=room_tasks_map,
                    current_day=current_day,
                    deferred_carpet_keys=deferred_carpet_keys,
                    use_deferred_only=assist_use_deferred_only,
                )

                if (
                    not assist_use_deferred_only
                    and (assist_blocked or assist_not_ready)
                    and assist_ready
                ):
                    for task in assist_blocked + assist_not_ready:
                        deferred_carpet_keys.add(task.identity_key())

                assist_ready.sort(key=task_sort_key)

                assist_log, assist_used = schedule_task_list(
                    tasks_to_run=assist_ready,
                    crew_capacity=general_unused_capacity,
                    crew_type="General Assist",
                )
                if assist_used > 0:
                    carpet_assist_school = assist_school
                    general_used_capacity = clean_hours(general_used_capacity + assist_used)
                    work_log.extend(assist_log)

        used_capacity = clean_hours(general_used_capacity + carpet_used_capacity)

        total_used_hours = clean_hours(total_used_hours + used_capacity)
        unused_capacity = clean_hours(daily_capacity - used_capacity)
        general_unused_capacity = clean_hours(general_capacity - general_used_capacity)
        carpet_unused_capacity = clean_hours(carpet_capacity - carpet_used_capacity)

        if general_school and carpet_assist_school and general_school == carpet_assist_school and carpet_capacity <= 0:
            active_school_name = general_school
        else:
            school_labels: List[str] = []
            if general_school:
                school_labels.append(f"General: {general_school}")
            if carpet_capacity > 0 and carpet_school:
                school_labels.append(f"Carpet: {carpet_school}")
            elif carpet_assist_school:
                school_labels.append(f"General Assist Carpet: {carpet_assist_school}")

            if not school_labels:
                active_school_name = ""
            elif len(school_labels) == 1:
                active_school_name = school_labels[0].split(": ", 1)[1]
            else:
                active_school_name = " | ".join(school_labels)

        day_status_note = build_day_status_note(
            general_school=general_school,
            carpet_school=carpet_school,
            carpet_assist_school=carpet_assist_school,
            general_available_today=general_available_today,
            general_blocked_today=general_blocked_today,
            carpet_ready_today=carpet_ready_today,
            carpet_blocked_today=carpet_blocked_today,
            carpet_not_ready_today=carpet_not_ready_today,
            general_use_deferred_only=general_use_deferred_only,
            carpet_use_deferred_only=carpet_use_deferred_only,
            carpet_staff=carpet_staff,
            carpet_used_capacity=carpet_used_capacity if carpet_capacity > 0 else clean_hours(sum(item.hours_done for item in work_log if item.crew_type == "General Assist")),
        )

        exterior_used_as_access_work = any(
            item.phase_name == "Exterior"
            for item in work_log
        ) and any(
            task.school_name == general_school
            and task.phase_name not in {"Exterior", "Carpet", "Site Transition", "Transition / Logistics"}
            and task.remaining_hours > 0
            and task.available_day > current_day
            for task in tasks
        )

        if exterior_used_as_access_work:
            day_status_note = (
                f"{day_status_note} ACCESS WORKAROUND: Interior rooms are blocked, so exterior work is being used as productive filler."
            )
        if access_downtime_warning:
            day_status_note = f"{day_status_note} {access_downtime_warning}"

        day_results.append(
            ScheduleDayResult(
                day=current_day,
                effective_staff=effective_staff,
                daily_capacity=daily_capacity,
                used_capacity=used_capacity,
                unused_capacity=unused_capacity,
                work_log=work_log,
                active_school_name=active_school_name,
                status_note=day_status_note,
                general_staff=general_staff,
                carpet_staff=carpet_staff,
                general_capacity=general_capacity,
                carpet_capacity=carpet_capacity,
                general_used_capacity=general_used_capacity,
                carpet_used_capacity=carpet_used_capacity,
                general_unused_capacity=general_unused_capacity,
                carpet_unused_capacity=carpet_unused_capacity,
            )
        )

        if used_capacity > 0:
            no_progress_days = 0
        else:
            no_progress_days += 1

        if no_progress_days >= 10:
            break

        current_day += 1

    finish_day = day_results[-1].day if day_results else settings.current_day
    remaining_backlog_hours = clean_hours(sum(task.remaining_hours for task in tasks))
    met_deadline = remaining_backlog_hours <= 0 and finish_day <= settings.target_end_day

    recommendation = build_recommendation(
        start_available_backlog_hours=start_available_backlog_hours,
        start_blocked_backlog_hours=start_blocked_backlog_hours,
        capacity_to_deadline_hours=capacity_to_deadline_hours,
        finish_day=finish_day,
        settings=settings,
        remaining_backlog_hours=remaining_backlog_hours,
    )

    return ScheduleResult(
        schedule_name=settings.schedule_name,
        target_end_day=settings.target_end_day,
        current_day=settings.current_day,
        finish_day=finish_day,
        met_deadline=met_deadline,
        total_planned_hours=total_planned_hours,
        completed_hours_before_run=completed_hours_before_run,
        remaining_hours_at_start=remaining_hours_at_start,
        total_used_hours=total_used_hours,
        days=day_results,
        remaining_backlog_hours=remaining_backlog_hours,
        task_items=tasks,
        recommendation=recommendation,
    )