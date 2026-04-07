from dataclasses import dataclass, field
from typing import List


@dataclass
class Room:
    school_name: str
    school_order: int
    building_name: str
    zone_name: str
    room_name: str
    room_order: int
    carpet_sqft: float
    tile_sqft: float
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
    hours_per_staff_per_day: float
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


@dataclass
class StaffingDay:
    day: int
    available_staff: int
    carpet_staff_reserved: int
    absences: int
    temporary_help: int

    def effective_cleaning_staff(self) -> int:
        value = (
            self.available_staff
            - self.carpet_staff_reserved
            - self.absences
            + self.temporary_help
        )
        return max(0, value)


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


@dataclass
class WorkLogEntry:
    school_name: str
    room_name: str
    phase_name: str
    hours_done: float


@dataclass
class ScheduleDayResult:
    day: int
    effective_staff: int
    daily_capacity: float
    used_capacity: float
    unused_capacity: float
    work_log: List[WorkLogEntry] = field(default_factory=list)


@dataclass
class ScheduleResult:
    schedule_name: str
    target_end_day: int
    finish_day: int
    met_deadline: bool
    total_planned_hours: float
    total_used_hours: float
    days: List[ScheduleDayResult]
    remaining_backlog_hours: float
    task_items: List[TaskItem]