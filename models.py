from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Room:
    school_name: str
    school_order: int
    building_name: str
    zone_name: str
    room_name: str
    room_order: int

    total_room_sqft: float
    carpet_sqft: float
    tile_strip_wax_sqft: float
    scrub_only_vct_sqft: float

    available_day: int
    include_deep_clean: bool
    include_strip: bool
    include_wax: bool
    include_carpet: bool
    include_exterior: bool
    notes: str = ""


@dataclass
class School:
    name: str
    order: int
    rooms: List[Room] = field(default_factory=list)

    def add_room(self, room: Room) -> None:
        self.rooms.append(room)
        self.rooms.sort(key=lambda r: (r.room_order, r.room_name))


@dataclass
class ScheduleSettings:
    schedule_name: str
    target_end_day: int

    scheduled_shift_hours_per_day: float
    lunch_hours_per_day: float
    break_hours_per_day: float
    setup_hours_per_day: float
    cleanup_hours_per_day: float
    productive_hours_per_staff_per_day: float

    current_day: int
    include_deep_clean: bool
    include_strip: bool
    include_wax: bool
    include_carpet: bool
    include_exterior: bool
    deep_clean_rate_sqft_per_hour: float
    strip_rate_sqft_per_hour: float
    wax_rate_sqft_per_hour: float
    carpet_rate_sqft_per_hour: float
    exterior_rate_sqft_per_hour: float
    wax_coats: int
    transition_hours_per_school: float

    day_start_time: str = "7:30 AM"
    work_start_time: str = "7:45 AM"
    first_break_time: str = "10:00 AM"
    lunch_time: str = "12:00 PM"
    second_break_time: str = "2:00 PM"
    cleanup_start_time: str = "3:30 PM"
    day_end_time: str = "4:00 PM"

    schedule_start_date: str = "2026-06-01"
    target_end_date: str = ""
    work_on_weekends: bool = False
    paid_holidays_in_range: int = 0

    @property
    def non_productive_hours_per_staff_per_day(self) -> float:
        return (
            self.lunch_hours_per_day
            + self.break_hours_per_day
            + self.setup_hours_per_day
            + self.cleanup_hours_per_day
        )

    @property
    def total_accounted_hours_per_staff_per_day(self) -> float:
        return (
            self.productive_hours_per_staff_per_day
            + self.non_productive_hours_per_staff_per_day
        )

    def validate_or_normalize(self) -> None:
        calculated_productive = (
            self.scheduled_shift_hours_per_day
            - self.lunch_hours_per_day
            - self.break_hours_per_day
            - self.setup_hours_per_day
            - self.cleanup_hours_per_day
        )

        if calculated_productive < 0:
            calculated_productive = 0.0

        if self.productive_hours_per_staff_per_day <= 0:
            self.productive_hours_per_staff_per_day = calculated_productive

        if abs(self.productive_hours_per_staff_per_day - calculated_productive) > 0.01:
            self.productive_hours_per_staff_per_day = calculated_productive

        self.target_end_day = max(1, int(self.target_end_day))
        self.current_day = max(1, int(self.current_day))
        self.paid_holidays_in_range = max(0, int(self.paid_holidays_in_range))


@dataclass
class StaffingDay:
    day: int
    available_staff: int
    carpet_staff_reserved: int
    absences: int
    temporary_help: int

    def net_available_staff(self) -> int:
        value = self.available_staff - self.absences + self.temporary_help
        return max(0, value)

    def carpet_crew_staff(self) -> int:
        return max(0, min(self.carpet_staff_reserved, self.net_available_staff()))

    def general_crew_staff(self) -> int:
        return max(0, self.net_available_staff() - self.carpet_crew_staff())

    def effective_cleaning_staff(self) -> int:
        return self.general_crew_staff() + self.carpet_crew_staff()


@dataclass
class ProgressEntry:
    school_name: str
    room_name: str
    phase_name: str
    hours_completed: float
    building_name: str = ""
    zone_name: str = ""


@dataclass
class TaskItem:
    school_name: str
    school_order: int
    building_name: str
    zone_name: str
    room_name: str
    room_order: int
    phase_name: str
    available_day: int
    total_hours: float
    remaining_hours: float
    notes: str = ""

    def identity_key(self) -> tuple[str, str, str, str, str]:
        return (
            self.school_name,
            self.building_name,
            self.zone_name,
            self.room_name,
            self.phase_name,
        )

    def legacy_progress_key(self) -> tuple[str, str, str]:
        return (
            self.school_name,
            self.room_name,
            self.phase_name,
        )


@dataclass
class WorkLogEntry:
    school_name: str
    room_name: str
    phase_name: str
    hours_done: float
    building_name: str = ""
    zone_name: str = ""
    available_day: Optional[int] = None
    note: str = ""
    crew_type: str = "General"


@dataclass
class ScheduleDayResult:
    day: int
    effective_staff: int
    daily_capacity: float
    used_capacity: float
    unused_capacity: float
    work_log: List[WorkLogEntry] = field(default_factory=list)
    active_school_name: str = ""
    status_note: str = ""

    general_staff: int = 0
    carpet_staff: int = 0
    general_capacity: float = 0.0
    carpet_capacity: float = 0.0
    general_used_capacity: float = 0.0
    carpet_used_capacity: float = 0.0
    general_unused_capacity: float = 0.0
    carpet_unused_capacity: float = 0.0


@dataclass
class RecommendationSummary:
    status_label: str
    bottleneck_type: str
    available_backlog_hours: float
    blocked_backlog_hours: float
    extra_staff_days_needed: float
    recommended_action: str
    based_on_start_of_run: bool = True
    capacity_to_deadline_hours: float = 0.0
    total_backlog_hours: float = 0.0


@dataclass
class ScheduleResult:
    schedule_name: str
    target_end_day: int
    current_day: int
    finish_day: int
    met_deadline: bool
    total_planned_hours: float
    completed_hours_before_run: float
    remaining_hours_at_start: float
    total_used_hours: float
    days: List[ScheduleDayResult]
    remaining_backlog_hours: float
    task_items: List[TaskItem]
    recommendation: RecommendationSummary