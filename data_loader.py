from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import load_workbook

from models import ProgressEntry, Room, School, ScheduleSettings, StaffingDay


def parse_bool(value) -> bool:
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


def _sheet_exists(wb, sheet_name: str) -> bool:
    return sheet_name in wb.sheetnames


def _normalize_header(value) -> str:
    return str(value).strip() if value is not None else ""


def _sheet_to_dict_rows(ws) -> List[Dict[str, str]]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [_normalize_header(cell) for cell in rows[0]]
    data_rows: List[Dict[str, str]] = []

    for row in rows[1:]:
        if row is None:
            continue

        record: Dict[str, str] = {}
        has_any_value = False

        for i, header in enumerate(headers):
            if not header:
                continue
            value = row[i] if i < len(row) else ""
            if value is not None and str(value).strip() != "":
                has_any_value = True
            record[header] = "" if value is None else str(value).strip()

        if has_any_value:
            data_rows.append(record)

    return data_rows


def _load_setup_raw(workbook_path: str) -> Dict[str, str]:
    wb = load_workbook(workbook_path, data_only=True)

    if not _sheet_exists(wb, "Setup"):
        raise ValueError("Workbook is missing required sheet: Setup")

    ws = wb["Setup"]
    raw: Dict[str, str] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        key = "" if row[0] is None else str(row[0]).strip()
        value = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()

        if not key:
            continue

        raw[key] = value

    return raw


def load_settings(workbook_path: str) -> ScheduleSettings:
    raw = _load_setup_raw(workbook_path)

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
    workbook_path: str,
) -> Tuple[List[Room], List[School]]:
    wb = load_workbook(workbook_path, data_only=True)

    if not _sheet_exists(wb, "Rooms"):
        raise ValueError("Workbook is missing required sheet: Rooms")

    ws = wb["Rooms"]
    rows = _sheet_to_dict_rows(ws)

    rooms: List[Room] = []
    schools_by_name: Dict[str, School] = {}

    for row in rows:
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


def load_staffing(workbook_path: str) -> List[StaffingDay]:
    wb = load_workbook(workbook_path, data_only=True)

    if not _sheet_exists(wb, "Staffing"):
        raise ValueError("Workbook is missing required sheet: Staffing")

    ws = wb["Staffing"]
    rows = _sheet_to_dict_rows(ws)

    staffing_days: List[StaffingDay] = []

    for row in rows:
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


def load_progress(workbook_path: str) -> List[ProgressEntry]:
    wb = load_workbook(workbook_path, data_only=True)

    if not _sheet_exists(wb, "Progress"):
        return []

    ws = wb["Progress"]
    rows = _sheet_to_dict_rows(ws)

    progress_entries: List[ProgressEntry] = []

    for row in rows:
        school_name = row.get("school_name", "").strip()
        room_name = row.get("room_name", "").strip()
        phase_name = row.get("phase_name", "").strip()
        hours_completed = row.get("hours_completed", "").strip()

        if not school_name or not room_name or not phase_name or not hours_completed:
            continue

        progress_entries.append(
            ProgressEntry(
                school_name=school_name,
                building_name=row.get("building_name", "").strip(),
                zone_name=row.get("zone_name", "").strip(),
                room_name=room_name,
                phase_name=phase_name,
                hours_completed=float(hours_completed),
            )
        )

    return progress_entries


def validate_workbook(workbook_path: str) -> List[str]:
    errors: List[str] = []

    path = Path(workbook_path)
    if not path.exists():
        return [f"Workbook not found: {workbook_path}"]

    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        errors.append("Workbook must be an .xlsx or .xlsm file")

    try:
        wb = load_workbook(workbook_path, data_only=True)
    except Exception as exc:
        return [f"Could not open workbook: {exc}"]

    required_sheets = ["Setup", "Rooms", "Staffing"]
    for sheet in required_sheets:
        if sheet not in wb.sheetnames:
            errors.append(f"Missing required sheet: {sheet}")

    return errors