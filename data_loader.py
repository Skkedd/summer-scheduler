import csv
from typing import Dict, List, Tuple

from models import ProgressEntry, Room, School, ScheduleSettings, StaffingDay


def parse_bool(value: str) -> bool:
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def parse_float(raw: Dict[str, str], key: str, default: float) -> float:
    value = raw.get(key, "")
    if value is None or str(value).strip() == "":
        return default
    return float(str(value).strip())


def parse_int(raw: Dict[str, str], key: str, default: int) -> int:
    value = raw.get(key, "")
    if value is None or str(value).strip() == "":
        return default
    return int(str(value).strip())


def parse_str(raw: Dict[str, str], key: str, default: str) -> str:
    value = raw.get(key, "")
    if value is None or str(value).strip() == "":
        return default
    return str(value).strip()


def load_settings(settings_file_path: str = "data/settings.csv") -> ScheduleSettings:
    raw: Dict[str, str] = {}

    with open(settings_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            raw[row["setting"].strip()] = row["value"].strip()

    legacy_hours_per_staff_per_day = parse_float(
        raw,
        "hours_per_staff_per_day",
        6.75,
    )

    scheduled_shift_hours_per_day = parse_float(
        raw,
        "scheduled_shift_hours_per_day",
        8.5,
    )
    lunch_hours_per_day = parse_float(
        raw,
        "lunch_hours_per_day",
        0.5,
    )
    break_hours_per_day = parse_float(
        raw,
        "break_hours_per_day",
        0.5,
    )
    setup_hours_per_day = parse_float(
        raw,
        "setup_hours_per_day",
        0.25,
    )
    cleanup_hours_per_day = parse_float(
        raw,
        "cleanup_hours_per_day",
        0.5,
    )

    calculated_productive_default = (
        scheduled_shift_hours_per_day
        - lunch_hours_per_day
        - break_hours_per_day
        - setup_hours_per_day
        - cleanup_hours_per_day
    )

    if calculated_productive_default < 0:
        calculated_productive_default = 0.0

    productive_hours_per_staff_per_day = parse_float(
        raw,
        "productive_hours_per_staff_per_day",
        calculated_productive_default
        if "hours_per_staff_per_day" not in raw
        else legacy_hours_per_staff_per_day,
    )

    settings = ScheduleSettings(
        schedule_name=parse_str(raw, "schedule_name", "Summer Schedule"),
        target_end_day=parse_int(raw, "target_end_day", 20),
        scheduled_shift_hours_per_day=scheduled_shift_hours_per_day,
        lunch_hours_per_day=lunch_hours_per_day,
        break_hours_per_day=break_hours_per_day,
        setup_hours_per_day=setup_hours_per_day,
        cleanup_hours_per_day=cleanup_hours_per_day,
        productive_hours_per_staff_per_day=productive_hours_per_staff_per_day,
        current_day=parse_int(raw, "current_day", 1),
        include_deep_clean=parse_bool(raw.get("include_deep_clean", "True")),
        include_strip=parse_bool(raw.get("include_strip", "True")),
        include_wax=parse_bool(raw.get("include_wax", "True")),
        include_carpet=parse_bool(raw.get("include_carpet", "True")),
        include_exterior=parse_bool(raw.get("include_exterior", "False")),
        deep_clean_rate_sqft_per_hour=parse_float(
            raw,
            "deep_clean_rate_sqft_per_hour",
            400.0,
        ),
        strip_rate_sqft_per_hour=parse_float(
            raw,
            "strip_rate_sqft_per_hour",
            300.0,
        ),
        wax_rate_sqft_per_hour=parse_float(
            raw,
            "wax_rate_sqft_per_hour",
            600.0,
        ),
        carpet_rate_sqft_per_hour=parse_float(
            raw,
            "carpet_rate_sqft_per_hour",
            500.0,
        ),
        exterior_rate_sqft_per_hour=parse_float(
            raw,
            "exterior_rate_sqft_per_hour",
            1000.0,
        ),
        wax_coats=parse_int(raw, "wax_coats", 3),
        transition_hours_per_school=parse_float(
            raw,
            "transition_hours_per_school",
            0.0,
        ),
        day_start_time=parse_str(raw, "day_start_time", "7:30 AM"),
        work_start_time=parse_str(raw, "work_start_time", "7:45 AM"),
        first_break_time=parse_str(raw, "first_break_time", "10:00 AM"),
        lunch_time=parse_str(raw, "lunch_time", "12:00 PM"),
        second_break_time=parse_str(raw, "second_break_time", "2:00 PM"),
        cleanup_start_time=parse_str(raw, "cleanup_start_time", "3:30 PM"),
        day_end_time=parse_str(raw, "day_end_time", "4:00 PM"),
    )

    settings.validate_or_normalize()
    return settings


def load_rooms(
    rooms_file_path: str = "data/rooms.csv",
) -> Tuple[List[Room], List[School]]:
    rooms: List[Room] = []
    schools_by_name: Dict[str, School] = {}

    with open(rooms_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)

        for row in reader:
            room = Room(
                school_name=row["school_name"].strip(),
                school_order=int(row["school_order"]),
                building_name=row["building_name"].strip(),
                zone_name=row["zone_name"].strip(),
                room_name=row["room_name"].strip(),
                room_order=int(row["room_order"]),
                carpet_sqft=float(row["carpet_sqft"]),
                tile_sqft=float(row["tile_sqft"]),
                available_day=int(row["available_day"]),
                include_deep_clean=parse_bool(row["include_deep_clean"]),
                include_strip=parse_bool(row["include_strip"]),
                include_wax=parse_bool(row["include_wax"]),
                include_carpet=parse_bool(row["include_carpet"]),
                include_exterior=parse_bool(row["include_exterior"]),
                notes=row.get("notes", "").strip(),
            )
            rooms.append(room)

            if room.school_name not in schools_by_name:
                schools_by_name[room.school_name] = School(
                    name=room.school_name,
                    order=room.school_order,
                )

            schools_by_name[room.school_name].add_room(room)

    schools = sorted(schools_by_name.values(), key=lambda s: (s.order, s.name))
    return rooms, schools


def load_staffing(staffing_file_path: str = "data/staffing.csv") -> List[StaffingDay]:
    staffing_days: List[StaffingDay] = []

    with open(staffing_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)

        for row in reader:
            staffing_days.append(
                StaffingDay(
                    day=int(row["day"]),
                    available_staff=int(row["available_staff"]),
                    carpet_staff_reserved=int(row["carpet_staff_reserved"]),
                    absences=int(row["absences"]),
                    temporary_help=int(row["temporary_help"]),
                )
            )

    staffing_days.sort(key=lambda s: s.day)
    return staffing_days


def load_progress(progress_file_path: str = "data/progress.csv") -> List[ProgressEntry]:
    progress_entries: List[ProgressEntry] = []

    with open(progress_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        fieldnames = reader.fieldnames or []

        has_building_name = "building_name" in fieldnames
        has_zone_name = "zone_name" in fieldnames

        for row in reader:
            progress_entries.append(
                ProgressEntry(
                    school_name=row["school_name"].strip(),
                    building_name=row["building_name"].strip()
                    if has_building_name and row.get("building_name")
                    else "",
                    zone_name=row["zone_name"].strip()
                    if has_zone_name and row.get("zone_name")
                    else "",
                    room_name=row["room_name"].strip(),
                    phase_name=row["phase_name"].strip(),
                    hours_completed=float(row["hours_completed"]),
                )
            )

    return progress_entries