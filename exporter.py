from __future__ import annotations

import os
from pathlib import Path
from typing import Optional, Sequence, Set

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from calendar_math import format_date_label, workday_to_date
from models import ScheduleResult, ScheduleSettings


DISTRICT_FILE = "District Facility Data.xlsx"
ASSUMPTIONS_FILE = "Cleaning Planning Assumptions.xlsx"
RUN_INPUT_FILE = "Summer Scheduler Run Input.xlsx"
FULL_RESULTS_FILE = "Full Schedule Results.xlsx"
WORKER_EXPORT_FILE = "Worker Schedule Export.xlsx"

HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
INSTRUCTION_FILL = PatternFill(fill_type="solid", fgColor="FFF4CC")


def ensure_output_folder(folder_path: str = "output") -> None:
    os.makedirs(folder_path, exist_ok=True)


def _bold_header(ws, row_num: int = 1) -> None:
    for cell in ws[row_num]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_instruction_row(ws, row_num: int = 1) -> None:
    for cell in ws[row_num]:
        cell.font = Font(bold=True)
        cell.fill = INSTRUCTION_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def _freeze_header(ws, cell: str = "A2") -> None:
    ws.freeze_panes = cell


def _set_column_widths(ws, widths: dict[str, float]) -> None:
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _autosize_from_headers(ws, min_width: int = 14, max_width: int = 34) -> None:
    for col_idx, cell in enumerate(ws[1], start=1):
        value = str(cell.value or "")
        width = min(max(len(value) + 4, min_width), max_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _selected_site_set(selected_sites: Optional[Sequence[str]]) -> Optional[Set[str]]:
    if not selected_sites:
        return None
    return {item.strip() for item in selected_sites if str(item).strip()}


def _task_matches_sites(task, selected_sites: Optional[Set[str]]) -> bool:
    if selected_sites is None:
        return True
    return task.school_name in selected_sites


def _worklog_matches_sites(item, selected_sites: Optional[Set[str]]) -> bool:
    if selected_sites is None:
        return True
    return item.school_name in selected_sites


def _filtered_days(result: ScheduleResult, selected_sites: Optional[Set[str]]):
    filtered = []
    for day in result.days:
        matching_log = [item for item in day.work_log if _worklog_matches_sites(item, selected_sites)]
        if matching_log:
            filtered.append((day, matching_log))
    return filtered


def _filtered_tasks(result: ScheduleResult, selected_sites: Optional[Set[str]]):
    return [task for task in result.task_items if _task_matches_sites(task, selected_sites)]


def export_result_workbook(
    result: ScheduleResult,
    settings: ScheduleSettings,
    folder_path: str = "output",
    selected_sites: Optional[Sequence[str]] = None,
) -> str:
    ensure_output_folder(folder_path)
    file_path = os.path.join(folder_path, FULL_RESULTS_FILE)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    selected_site_set = _selected_site_set(selected_sites)
    filtered_tasks = _filtered_tasks(result, selected_site_set)
    filtered_days = _filtered_days(result, selected_site_set)

    ws = wb.create_sheet("Summary")
    ws.append(["Field", "Value"])
    _bold_header(ws)

    filtered_total_planned = round(sum(task.total_hours for task in filtered_tasks), 2)
    filtered_remaining = round(sum(task.remaining_hours for task in filtered_tasks), 2)
    filtered_used = round(sum(item.hours_done for _, log in filtered_days for item in log), 2)

    summary_rows = [
        ("Schedule Name", result.schedule_name),
        ("Schedule Start Date", settings.schedule_start_date),
        ("Target End Date", settings.target_end_date),
        ("Work On Weekends", settings.work_on_weekends),
        ("Paid Holidays In Range", settings.paid_holidays_in_range),
        ("Target End Day", result.target_end_day),
        ("Current Day", result.current_day),
        ("Projected Finish Day", result.finish_day),
        (
            "Projected Finish Date",
            format_date_label(
                workday_to_date(
                    settings.schedule_start_date,
                    result.finish_day,
                    settings.work_on_weekends,
                )
            ),
        ),
        ("Deadline Met", result.met_deadline),
        ("Export Scope", ", ".join(sorted(selected_site_set)) if selected_site_set else "All Sites"),
        ("Filtered Total Planned Hours", filtered_total_planned),
        ("Filtered Total Used Hours", filtered_used),
        ("Filtered Remaining Backlog Hours", filtered_remaining),
        ("Recommendation Status", result.recommendation.status_label),
        ("Bottleneck Type", result.recommendation.bottleneck_type),
        ("Recommended Action", result.recommendation.recommended_action),
    ]
    for row in summary_rows:
        ws.append(list(row))
    _set_column_widths(ws, {"A": 30, "B": 50})

    ws = wb.create_sheet("Days")
    ws.append(
        [
            "Workday Number",
            "Work Date",
            "Active Site",
            "Day Note",
            "Effective Staff",
            "Cleaning Staff",
            "Carpet Staff",
            "Cleaning Capacity Hours",
            "Carpet Capacity Hours",
            "Total Capacity Hours",
            "Used Hours",
            "Unused Hours",
            "Matching Work Items",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)

    for day, matching_log in filtered_days:
        work_date = workday_to_date(
            settings.schedule_start_date,
            day.day,
            settings.work_on_weekends,
        )
        ws.append(
            [
                day.day,
                format_date_label(work_date),
                day.active_school_name,
                day.status_note,
                day.effective_staff,
                day.general_staff,
                day.carpet_staff,
                round(day.general_capacity, 2),
                round(day.carpet_capacity, 2),
                round(day.daily_capacity, 2),
                round(sum(item.hours_done for item in matching_log), 2),
                round(day.daily_capacity - sum(item.hours_done for item in matching_log), 2),
                len(matching_log),
            ]
        )
    _autosize_from_headers(ws)

    ws = wb.create_sheet("Work Log")
    ws.append(
        [
            "Workday Number",
            "Work Date",
            "Crew Type",
            "Site",
            "Building",
            "Zone",
            "Room",
            "Task",
            "Available Day",
            "Hours Done",
            "Notes",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)

    for day, matching_log in filtered_days:
        work_date = workday_to_date(
            settings.schedule_start_date,
            day.day,
            settings.work_on_weekends,
        )
        for item in matching_log:
            ws.append(
                [
                    day.day,
                    format_date_label(work_date),
                    item.crew_type,
                    item.school_name,
                    item.building_name,
                    item.zone_name,
                    item.room_name,
                    item.phase_name,
                    item.available_day if item.available_day is not None else "",
                    round(item.hours_done, 2),
                    item.note,
                ]
            )
    _autosize_from_headers(ws)

    ws = wb.create_sheet("Tasks")
    ws.append(
        [
            "Site",
            "Building",
            "Zone",
            "Room",
            "Task",
            "Available Day",
            "Available Date",
            "Total Hours",
            "Remaining Hours",
            "Status",
            "Notes",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)

    for task in filtered_tasks:
        if task.remaining_hours <= 0:
            status = "Complete"
        elif task.remaining_hours < task.total_hours:
            status = "In Progress"
        else:
            status = "Not Started"

        available_date = workday_to_date(
            settings.schedule_start_date,
            task.available_day,
            settings.work_on_weekends,
        )

        ws.append(
            [
                task.school_name,
                task.building_name,
                task.zone_name,
                task.room_name,
                task.phase_name,
                task.available_day,
                format_date_label(available_date),
                round(task.total_hours, 2),
                round(task.remaining_hours, 2),
                status,
                task.notes,
            ]
        )
    _autosize_from_headers(ws)

    wb.save(file_path)
    return file_path


def export_worker_schedule_workbook(
    result: ScheduleResult,
    settings: ScheduleSettings,
    folder_path: str = "output",
    selected_sites: Optional[Sequence[str]] = None,
) -> str:
    ensure_output_folder(folder_path)
    file_path = os.path.join(folder_path, WORKER_EXPORT_FILE)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    selected_site_set = _selected_site_set(selected_sites)
    filtered_days = _filtered_days(result, selected_site_set)

    ws = wb.create_sheet("Daily Assignment Overview")
    ws.append(
        [
            "Date",
            "Workday Number",
            "Site",
            "Crew Type",
            "Main Task",
            "Hours",
            "Notes",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)

    for day, matching_log in filtered_days:
        work_date = workday_to_date(
            settings.schedule_start_date,
            day.day,
            settings.work_on_weekends,
        )

        for item in matching_log:
            main_task = f"{item.phase_name} - {item.room_name}"
            ws.append(
                [
                    format_date_label(work_date),
                    day.day,
                    item.school_name,
                    item.crew_type,
                    main_task,
                    round(item.hours_done, 2),
                    item.note,
                ]
            )
    _autosize_from_headers(ws)

    ws = wb.create_sheet("Detailed Daily Work")
    ws.append(
        [
            "Date",
            "Site",
            "Building",
            "Zone",
            "Room",
            "Work Type",
            "Crew Type",
            "Hours",
            "Notes",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)

    for day, matching_log in filtered_days:
        work_date = workday_to_date(
            settings.schedule_start_date,
            day.day,
            settings.work_on_weekends,
        )

        for item in matching_log:
            ws.append(
                [
                    format_date_label(work_date),
                    item.school_name,
                    item.building_name,
                    item.zone_name,
                    item.room_name,
                    item.phase_name,
                    item.crew_type,
                    round(item.hours_done, 2),
                    item.note,
                ]
            )
    _autosize_from_headers(ws)

    ws = wb.create_sheet("Summary")
    ws.append(["Field", "Value"])
    _bold_header(ws)
    _set_column_widths(ws, {"A": 28, "B": 50})

    ws.append(["Schedule Name", result.schedule_name])
    ws.append(["Export Scope", ", ".join(sorted(selected_site_set)) if selected_site_set else "All Sites"])
    ws.append(["Projected Finish Day", result.finish_day])
    ws.append(
        [
            "Projected Finish Date",
            format_date_label(
                workday_to_date(
                    settings.schedule_start_date,
                    result.finish_day,
                    settings.work_on_weekends,
                )
            ),
        ]
    )
    ws.append(["Deadline Met", result.met_deadline])
    ws.append(["Recommendation Status", result.recommendation.status_label])
    ws.append(["Recommended Action", result.recommendation.recommended_action])

    wb.save(file_path)
    return file_path


def create_input_template(file_path: str = "data/Summer Scheduler Run Input.xlsx") -> str:
    target = Path(file_path)
    base_dir = target.parent if target.suffix else target
    base_dir.mkdir(parents=True, exist_ok=True)

    district_path = base_dir / DISTRICT_FILE
    assumptions_path = base_dir / ASSUMPTIONS_FILE
    run_input_path = base_dir / RUN_INPUT_FILE

    _create_district_facility_data_template(district_path)
    _create_cleaning_planning_assumptions_template(assumptions_path)
    _create_summer_scheduler_run_input_template(run_input_path)

    return str(run_input_path)


def _create_district_facility_data_template(path: Path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Instructions"
    ws["A1"] = "District Facility Data"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "What this file is for"
    ws["A3"].font = Font(bold=True)
    ws["A4"] = "Use this workbook for slow-changing district facts such as sites, buildings, rooms and flooring makeup."
    ws["A6"] = "Flooring entry rules"
    ws["A6"].font = Font(bold=True)
    ws["A7"] = "Enter exact square footage when known."
    ws["A8"] = "Fractions are allowed when estimating, such as 0.33 or 0.50."
    ws["A9"] = "If both square footage and fraction are entered, square footage wins."
    ws["A10"] = "Scrub-Only VCT is its own flooring category and should not generate strip/wax work."
    ws.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Sites")
    ws.append(["Site Name", "Site Order", "Notes"])
    _bold_header(ws)
    _freeze_header(ws)
    ws.append(["WES", 1, "Wright Elementary"])
    ws.append(["JXW", 2, "JX Wilson"])
    ws.append(["RLS", 3, "RL Stevens"])
    _set_column_widths(ws, {"A": 18, "B": 12, "C": 30})

    ws = wb.create_sheet("Rooms")
    ws.append(
        [
            "Site Name",
            "Building Name",
            "Zone Name",
            "Room Name",
            "Room Order",
            "Total Room SqFt",
            "Carpet SqFt",
            "Carpet Fraction",
            "Strip/Wax Tile SqFt",
            "Strip/Wax Tile Fraction",
            "Scrub-Only VCT SqFt",
            "Scrub-Only VCT Fraction",
            "Room Use",
            "Notes",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)
    ws.append(
        [
            "WES",
            "Main",
            "Classrooms",
            "Room 1",
            1,
            870,
            "",
            0.33,
            "",
            0.67,
            "",
            "",
            "Classroom",
            "",
        ]
    )
    _set_column_widths(
        ws,
        {
            "A": 14,
            "B": 18,
            "C": 18,
            "D": 18,
            "E": 12,
            "F": 16,
            "G": 14,
            "H": 16,
            "I": 18,
            "J": 20,
            "K": 18,
            "L": 20,
            "M": 16,
            "N": 24,
        },
    )
    ws.insert_rows(1)
    ws["A1"] = (
        "Enter exact square footage when known. Fractions are allowed when estimating. "
        "If both are entered, square footage wins."
    )
    ws.merge_cells("A1:N1")
    _style_instruction_row(ws, 1)
    _freeze_header(ws, "A3")

    wb.save(path)


def _create_cleaning_planning_assumptions_template(path: Path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Instructions"
    ws["A1"] = "Cleaning Planning Assumptions"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "What this file is for"
    ws["A3"].font = Font(bold=True)
    ws["A4"] = "Use this workbook for editable planning logic such as rates, time model and wax coats."
    ws["A5"] = "Update this when real-world production data improves your assumptions."
    ws.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Setup")
    ws.append(["Setting", "Value", "Description"])
    _bold_header(ws)
    _freeze_header(ws)

    rows = [
        ("scheduled_shift_hours_per_day", 8.5, "Paid shift length"),
        ("lunch_hours_per_day", 0.5, "Lunch time"),
        ("break_hours_per_day", 0.5, "Total break time"),
        ("setup_hours_per_day", 0.25, "Setup/opening time"),
        ("cleanup_hours_per_day", 0.5, "Cleanup/lockup time"),
        ("productive_hours_per_staff_per_day", 6.75, "Real productive hours per worker"),
        ("deep_clean_rate_sqft_per_hour", 400, "Deep clean production rate"),
        ("strip_rate_sqft_per_hour", 300, "Strip production rate"),
        ("wax_rate_sqft_per_hour", 600, "Wax production rate"),
        ("carpet_rate_sqft_per_hour", 500, "Carpet production rate"),
        ("exterior_rate_sqft_per_hour", 1000, "Exterior production rate"),
        ("wax_coats", 3, "Number of wax coats"),
        ("transition_hours_per_school", 0.0, "Inter-school logistics time"),
        ("day_start_time", "7:30 AM", "Arrival/open"),
        ("work_start_time", "7:45 AM", "Work starts"),
        ("first_break_time", "10:00 AM", "First break"),
        ("lunch_time", "12:00 PM", "Lunch"),
        ("second_break_time", "2:00 PM", "Second break"),
        ("cleanup_start_time", "3:30 PM", "Cleanup starts"),
        ("day_end_time", "4:00 PM", "End of day"),
    ]
    for row in rows:
        ws.append(list(row))

    _set_column_widths(ws, {"A": 34, "B": 18, "C": 42})

    wb.save(path)


def _create_summer_scheduler_run_input_template(path: Path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Instructions"
    ws["A1"] = "Summer Scheduler Run Input"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "What this file is for"
    ws["A3"].font = Font(bold=True)
    ws["A4"] = "Use this workbook for current run / current summer data such as dates, staffing, holidays and progress."
    ws["A5"] = "This is the file you are most likely to update and rerun throughout the summer."
    ws.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Run Settings")
    ws.append(["Setting", "Value", "Description"])
    _bold_header(ws)
    _freeze_header(ws)
    rows = [
        ("schedule_name", "Summer Schedule", "Name shown in app and exports"),
        ("schedule_start_date", "2026-06-01", "First workday of run"),
        ("target_end_date", "2026-06-26", "Real calendar target date"),
        ("target_end_day", "", "Optional. Leave blank to auto-calculate from target_end_date"),
        ("paid_holidays_in_range", 0, "Weekday holidays inside the run span"),
        ("work_on_weekends", False, "True if weekends count as workdays"),
        ("current_day", 1, "Usually 1 for a fresh run"),
        ("include_deep_clean", True, "Run deep clean tasks"),
        ("include_strip", True, "Run strip tasks"),
        ("include_wax", True, "Run wax tasks"),
        ("include_carpet", True, "Run carpet tasks"),
        ("include_exterior", False, "Run exterior tasks"),
    ]
    for row in rows:
        ws.append(list(row))
    _set_column_widths(ws, {"A": 30, "B": 18, "C": 46})

    ws = wb.create_sheet("Room Scope")
    ws.append(
        [
            "Site Name",
            "Building Name",
            "Zone Name",
            "Room Name",
            "Available Day",
            "Include Deep Clean",
            "Include Strip",
            "Include Wax",
            "Include Carpet",
            "Include Exterior",
            "Notes",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)
    ws.append(["WES", "Main", "Classrooms", "Room 1", 1, True, True, True, True, False, ""])
    _set_column_widths(
        ws,
        {
            "A": 14,
            "B": 18,
            "C": 18,
            "D": 18,
            "E": 14,
            "F": 18,
            "G": 14,
            "H": 14,
            "I": 16,
            "J": 16,
            "K": 24,
        },
    )

    ws = wb.create_sheet("Staffing")
    ws.append(
        [
            "Day",
            "Available Staff",
            "Carpet Staff Reserved",
            "Absences",
            "Temporary Help",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)
    for day in range(1, 41):
        ws.append([day, 4, 0, 0, 0])
    _set_column_widths(ws, {"A": 10, "B": 16, "C": 22, "D": 12, "E": 16})

    ws = wb.create_sheet("Progress")
    ws.append(
        [
            "Site Name",
            "Building Name",
            "Zone Name",
            "Room Name",
            "Task",
            "Hours Completed",
        ]
    )
    _bold_header(ws)
    _freeze_header(ws)
    _set_column_widths(ws, {"A": 14, "B": 18, "C": 18, "D": 18, "E": 20, "F": 18})

    wb.save(path)