from __future__ import annotations

from datetime import timedelta
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import load_workbook

from calendar_math import parse_date_string
from models import ProgressEntry, Room, School, ScheduleSettings, StaffingDay


DISTRICT_FILE = "District Facility Data.xlsx"
ASSUMPTIONS_FILE = "Cleaning Planning Assumptions.xlsx"
RUN_INPUT_FILE = "Summer Scheduler Run Input.xlsx"


def parse_bool(value) -> bool:
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def parse_float(value, default: float = 0.0) -> float:
    if value is None or str(value).strip() == "":
        return default
    return float(str(value).strip())


def parse_int(value, default: int = 0) -> int:
    if value is None or str(value).strip() == "":
        return default
    return int(float(str(value).strip()))


def parse_str(value, default: str = "") -> str:
    if value is None or str(value).strip() == "":
        return default
    return str(value).strip()


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


def _resolve_base_dir(workbook_path: str | None) -> Path:
    if workbook_path:
        candidate = Path(workbook_path)
        if candidate.exists():
            if candidate.is_dir():
                return candidate.resolve()
            return candidate.parent.resolve()

        if candidate.suffix.lower() in {".xlsx", ".xlsm"}:
            return candidate.parent.resolve()

        return candidate.resolve()

    return (Path(__file__).resolve().parent / "data").resolve()


def _open_required_workbook(path: Path):
    if not path.exists():
        raise ValueError(f"Workbook not found: {path}")
    return load_workbook(path, data_only=True)


def _get_workbook_paths(workbook_path: str | None) -> Dict[str, Path]:
    base_dir = _resolve_base_dir(workbook_path)
    return {
        "district": base_dir / DISTRICT_FILE,
        "assumptions": base_dir / ASSUMPTIONS_FILE,
        "run_input": base_dir / RUN_INPUT_FILE,
    }


def _calculate_workdays(
    start_date_str: str,
    end_date_str: str,
    include_weekends: bool,
    paid_holidays: int,
) -> int:
    start_date = parse_date_string(start_date_str)
    end_date = parse_date_string(end_date_str)

    if end_date < start_date:
        return 0

    count = 0
    current = start_date
    while current <= end_date:
        if include_weekends or current.weekday() < 5:
            count += 1
        current += timedelta(days=1)

    count -= paid_holidays
    return max(count, 0)


def _sheet_exists(wb, sheet_name: str) -> bool:
    return sheet_name in wb.sheetnames


def _load_key_value_sheet(workbook_path: Path, sheet_name: str) -> Dict[str, str]:
    wb = _open_required_workbook(workbook_path)

    if not _sheet_exists(wb, sheet_name):
        raise ValueError(f"{workbook_path.name} is missing required sheet: {sheet_name}")

    ws = wb[sheet_name]
    raw: Dict[str, str] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        key = "" if row[0] is None else str(row[0]).strip()
        value = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
        if key:
            raw[key] = value

    return raw


def load_settings(workbook_path: str | None = None) -> ScheduleSettings:
    paths = _get_workbook_paths(workbook_path)
    assumptions = _load_key_value_sheet(paths["assumptions"], "Setup")
    run_input = _load_key_value_sheet(paths["run_input"], "Run Settings")

    scheduled_shift_hours_per_day = parse_float(
        assumptions.get("scheduled_shift_hours_per_day"),
        8.5,
    )
    lunch_hours_per_day = parse_float(assumptions.get("lunch_hours_per_day"), 0.5)
    break_hours_per_day = parse_float(assumptions.get("break_hours_per_day"), 0.5)
    setup_hours_per_day = parse_float(assumptions.get("setup_hours_per_day"), 0.25)
    cleanup_hours_per_day = parse_float(assumptions.get("cleanup_hours_per_day"), 0.5)

    calculated_productive_default = (
        scheduled_shift_hours_per_day
        - lunch_hours_per_day
        - break_hours_per_day
        - setup_hours_per_day
        - cleanup_hours_per_day
    )
    if calculated_productive_default < 0:
        calculated_productive_default = 0.0

    schedule_start_date = parse_str(run_input.get("schedule_start_date"), "2026-06-01")
    target_end_date = parse_str(run_input.get("target_end_date"), "")
    work_on_weekends = parse_bool(run_input.get("work_on_weekends", "False"))
    paid_holidays = parse_int(run_input.get("paid_holidays_in_range"), 0)

    explicit_target_end_day = parse_int(run_input.get("target_end_day"), 0)
    if explicit_target_end_day > 0:
        target_end_day = explicit_target_end_day
    elif target_end_date:
        target_end_day = _calculate_workdays(
            schedule_start_date,
            target_end_date,
            work_on_weekends,
            paid_holidays,
        )
    else:
        target_end_day = 20

    settings = ScheduleSettings(
        schedule_name=parse_str(run_input.get("schedule_name"), "Summer Schedule"),
        target_end_day=target_end_day,
        scheduled_shift_hours_per_day=scheduled_shift_hours_per_day,
        lunch_hours_per_day=lunch_hours_per_day,
        break_hours_per_day=break_hours_per_day,
        setup_hours_per_day=setup_hours_per_day,
        cleanup_hours_per_day=cleanup_hours_per_day,
        productive_hours_per_staff_per_day=parse_float(
            assumptions.get("productive_hours_per_staff_per_day"),
            calculated_productive_default,
        ),
        current_day=parse_int(run_input.get("current_day"), 1),
        include_deep_clean=parse_bool(run_input.get("include_deep_clean", "True")),
        include_strip=parse_bool(run_input.get("include_strip", "True")),
        include_wax=parse_bool(run_input.get("include_wax", "True")),
        include_carpet=parse_bool(run_input.get("include_carpet", "True")),
        include_exterior=parse_bool(run_input.get("include_exterior", "False")),
        deep_clean_rate_sqft_per_hour=parse_float(
            assumptions.get("deep_clean_rate_sqft_per_hour"),
            400.0,
        ),
        strip_rate_sqft_per_hour=parse_float(
            assumptions.get("strip_rate_sqft_per_hour"),
            300.0,
        ),
        wax_rate_sqft_per_hour=parse_float(
            assumptions.get("wax_rate_sqft_per_hour"),
            600.0,
        ),
        carpet_rate_sqft_per_hour=parse_float(
            assumptions.get("carpet_rate_sqft_per_hour"),
            500.0,
        ),
        exterior_rate_sqft_per_hour=parse_float(
            assumptions.get("exterior_rate_sqft_per_hour"),
            1000.0,
        ),
        wax_coats=parse_int(assumptions.get("wax_coats"), 3),
        transition_hours_per_school=parse_float(
            assumptions.get("transition_hours_per_school"),
            0.0,
        ),
        day_start_time=parse_str(assumptions.get("day_start_time"), "7:30 AM"),
        work_start_time=parse_str(assumptions.get("work_start_time"), "7:45 AM"),
        first_break_time=parse_str(assumptions.get("first_break_time"), "10:00 AM"),
        lunch_time=parse_str(assumptions.get("lunch_time"), "12:00 PM"),
        second_break_time=parse_str(assumptions.get("second_break_time"), "2:00 PM"),
        cleanup_start_time=parse_str(assumptions.get("cleanup_start_time"), "3:30 PM"),
        day_end_time=parse_str(assumptions.get("day_end_time"), "4:00 PM"),
        schedule_start_date=schedule_start_date,
        target_end_date=target_end_date,
        work_on_weekends=work_on_weekends,
        paid_holidays_in_range=paid_holidays,
    )

    settings.validate_or_normalize()
    return settings


def _resolve_floor_sqft(
    total_room_sqft: float,
    exact_sqft: float,
    fraction: float,
) -> float:
    if exact_sqft > 0:
        return exact_sqft
    if fraction > 0 and total_room_sqft > 0:
        return total_room_sqft * fraction
    return 0.0


def load_rooms(workbook_path: str | None = None) -> Tuple[List[Room], List[School]]:
    paths = _get_workbook_paths(workbook_path)

    district_wb = _open_required_workbook(paths["district"])
    run_input_wb = _open_required_workbook(paths["run_input"])

    if not _sheet_exists(district_wb, "Sites"):
        raise ValueError("District Facility Data.xlsx is missing required sheet: Sites")
    if not _sheet_exists(district_wb, "Rooms"):
        raise ValueError("District Facility Data.xlsx is missing required sheet: Rooms")
    if not _sheet_exists(run_input_wb, "Room Scope"):
        raise ValueError("Summer Scheduler Run Input.xlsx is missing required sheet: Room Scope")

    site_rows = _sheet_to_dict_rows(district_wb["Sites"])
    room_rows = _sheet_to_dict_rows(district_wb["Rooms"])
    scope_rows = _sheet_to_dict_rows(run_input_wb["Room Scope"])

    site_order_map: Dict[str, int] = {}
    for row in site_rows:
        site_name = parse_str(row.get("Site Name"))
        if site_name:
            site_order_map[site_name] = parse_int(row.get("Site Order"), 999)

    scope_map: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}
    for row in scope_rows:
        key = (
            parse_str(row.get("Site Name")),
            parse_str(row.get("Building Name")),
            parse_str(row.get("Zone Name")),
            parse_str(row.get("Room Name")),
        )
        scope_map[key] = row

    rooms: List[Room] = []
    schools_by_name: Dict[str, School] = {}

    for row in room_rows:
        school_name = parse_str(row.get("Site Name"))
        building_name = parse_str(row.get("Building Name"))
        zone_name = parse_str(row.get("Zone Name"))
        room_name = parse_str(row.get("Room Name"))
        room_order = parse_int(row.get("Room Order"), 999)
        total_room_sqft = parse_float(row.get("Total Room SqFt"), 0.0)

        carpet_sqft = _resolve_floor_sqft(
            total_room_sqft=total_room_sqft,
            exact_sqft=parse_float(row.get("Carpet SqFt"), 0.0),
            fraction=parse_float(row.get("Carpet Fraction"), 0.0),
        )
        tile_strip_wax_sqft = _resolve_floor_sqft(
            total_room_sqft=total_room_sqft,
            exact_sqft=parse_float(row.get("Strip/Wax Tile SqFt"), 0.0),
            fraction=parse_float(row.get("Strip/Wax Tile Fraction"), 0.0),
        )
        scrub_only_vct_sqft = _resolve_floor_sqft(
            total_room_sqft=total_room_sqft,
            exact_sqft=parse_float(row.get("Scrub-Only VCT SqFt"), 0.0),
            fraction=parse_float(row.get("Scrub-Only VCT Fraction"), 0.0),
        )

        key = (school_name, building_name, zone_name, room_name)
        scope = scope_map.get(key, {})

        room = Room(
            school_name=school_name,
            school_order=site_order_map.get(school_name, 999),
            building_name=building_name,
            zone_name=zone_name,
            room_name=room_name,
            room_order=room_order,
            total_room_sqft=total_room_sqft,
            carpet_sqft=carpet_sqft,
            tile_strip_wax_sqft=tile_strip_wax_sqft,
            scrub_only_vct_sqft=scrub_only_vct_sqft,
            available_day=parse_int(scope.get("Available Day"), 1),
            include_deep_clean=parse_bool(scope.get("Include Deep Clean", "True")),
            include_strip=parse_bool(scope.get("Include Strip", "True")),
            include_wax=parse_bool(scope.get("Include Wax", "True")),
            include_carpet=parse_bool(scope.get("Include Carpet", "True")),
            include_exterior=parse_bool(scope.get("Include Exterior", "False")),
            notes=parse_str(scope.get("Notes"), ""),
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


def load_staffing(workbook_path: str | None = None) -> List[StaffingDay]:
    paths = _get_workbook_paths(workbook_path)
    wb = _open_required_workbook(paths["run_input"])

    if not _sheet_exists(wb, "Staffing"):
        raise ValueError("Summer Scheduler Run Input.xlsx is missing required sheet: Staffing")

    rows = _sheet_to_dict_rows(wb["Staffing"])
    staffing_days: List[StaffingDay] = []

    for row in rows:
        staffing_days.append(
            StaffingDay(
                day=parse_int(row.get("Day"), 1),
                available_staff=parse_int(row.get("Available Staff"), 0),
                carpet_staff_reserved=parse_int(row.get("Carpet Staff Reserved"), 0),
                absences=parse_int(row.get("Absences"), 0),
                temporary_help=parse_int(row.get("Temporary Help"), 0),
            )
        )

    staffing_days.sort(key=lambda s: s.day)
    return staffing_days


def load_progress(workbook_path: str | None = None) -> List[ProgressEntry]:
    paths = _get_workbook_paths(workbook_path)
    wb = _open_required_workbook(paths["run_input"])

    if not _sheet_exists(wb, "Progress"):
        return []

    rows = _sheet_to_dict_rows(wb["Progress"])
    progress_entries: List[ProgressEntry] = []

    for row in rows:
        school_name = parse_str(row.get("Site Name"))
        room_name = parse_str(row.get("Room Name"))
        phase_name = parse_str(row.get("Task"))
        hours_completed = parse_str(row.get("Hours Completed"))

        if not school_name or not room_name or not phase_name or not hours_completed:
            continue

        progress_entries.append(
            ProgressEntry(
                school_name=school_name,
                building_name=parse_str(row.get("Building Name"), ""),
                zone_name=parse_str(row.get("Zone Name"), ""),
                room_name=room_name,
                phase_name=phase_name,
                hours_completed=float(hours_completed),
            )
        )

    return progress_entries


def validate_workbook(workbook_path: str | None = None) -> List[str]:
    errors: List[str] = []
    paths = _get_workbook_paths(workbook_path)

    for _, path in paths.items():
        if not path.exists():
            errors.append(f"Missing required workbook: {path.name}")

    if errors:
        return errors

    try:
        district_wb = load_workbook(paths["district"], data_only=True)
        assumptions_wb = load_workbook(paths["assumptions"], data_only=True)
        run_input_wb = load_workbook(paths["run_input"], data_only=True)
    except Exception as exc:
        return [f"Could not open workbook set: {exc}"]

    for sheet in ["Sites", "Rooms"]:
        if sheet not in district_wb.sheetnames:
            errors.append(f"District Facility Data.xlsx missing required sheet: {sheet}")

    if "Setup" not in assumptions_wb.sheetnames:
        errors.append("Cleaning Planning Assumptions.xlsx missing required sheet: Setup")

    for sheet in ["Run Settings", "Room Scope", "Staffing"]:
        if sheet not in run_input_wb.sheetnames:
            errors.append(f"Summer Scheduler Run Input.xlsx missing required sheet: {sheet}")

    if not errors:
        try:
            run_raw = _load_key_value_sheet(paths["run_input"], "Run Settings")
            assumptions_raw = _load_key_value_sheet(paths["assumptions"], "Setup")

            parse_date_string(parse_str(run_raw.get("schedule_start_date"), ""))
            target_end_date = parse_str(run_raw.get("target_end_date"), "")
            if target_end_date:
                parse_date_string(target_end_date)

            parse_float(assumptions_raw.get("deep_clean_rate_sqft_per_hour"), 400.0)
            parse_float(assumptions_raw.get("strip_rate_sqft_per_hour"), 300.0)
            parse_float(assumptions_raw.get("wax_rate_sqft_per_hour"), 600.0)
            parse_float(assumptions_raw.get("carpet_rate_sqft_per_hour"), 500.0)
        except Exception as exc:
            errors.append(f"Invalid workbook content: {exc}")

    return errors