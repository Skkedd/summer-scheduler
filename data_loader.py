import csv
from typing import Dict, List, Tuple

from models import Room, School, ScheduleSettings, StaffingDay


def parse_bool(value: str) -> bool:
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def load_settings(settings_file_path: str = "data/settings.csv") -> ScheduleSettings:
    raw: Dict[str, str] = {}

    with open(settings_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            raw[row["setting"].strip()] = row["value"].strip()

    return ScheduleSettings(
        schedule_name=raw.get("schedule_name", "Summer Schedule"),
        target_end_day=int(raw.get("target_end_day", 20)),
        hours_per_staff_per_day=float(raw.get("hours_per_staff_per_day", 8)),
        include_deep_clean=parse_bool(raw.get("include_deep_clean", "True")),
        include_strip=parse_bool(raw.get("include_strip", "True")),
        include_wax=parse_bool(raw.get("include_wax", "True")),
        include_carpet=parse_bool(raw.get("include_carpet", "True")),
        include_exterior=parse_bool(raw.get("include_exterior", "False")),
        deep_clean_rate_sqft_per_hour=float(
            raw.get("deep_clean_rate_sqft_per_hour", 400)
        ),
        strip_rate_sqft_per_hour=float(raw.get("strip_rate_sqft_per_hour", 300)),
        wax_rate_sqft_per_hour=float(raw.get("wax_rate_sqft_per_hour", 600)),
        carpet_rate_sqft_per_hour=float(raw.get("carpet_rate_sqft_per_hour", 500)),
        exterior_rate_sqft_per_hour=float(raw.get("exterior_rate_sqft_per_hour", 1000)),
        wax_coats=int(raw.get("wax_coats", 3)),
        transition_hours_per_school=float(raw.get("transition_hours_per_school", 0.0)),
    )


def load_rooms(rooms_file_path: str = "data/rooms.csv") -> Tuple[List[Room], List[School]]:
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