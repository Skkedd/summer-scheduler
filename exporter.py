from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from models import ScheduleResult


def ensure_output_folder(folder_path: str = "output") -> None:
    os.makedirs(folder_path, exist_ok=True)


def _bold_header(ws, row_num: int = 1) -> None:
    for cell in ws[row_num]:
        cell.font = Font(bold=True)


def export_result_workbook(result: ScheduleResult, folder_path: str = "output") -> str:
    ensure_output_folder(folder_path)
    file_path = os.path.join(folder_path, "summer_scheduler_output.xlsx")

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    # Summary
    ws = wb.create_sheet("Summary")
    summary_rows = [
        ("schedule_name", result.schedule_name),
        ("target_end_day", result.target_end_day),
        ("current_day", result.current_day),
        ("finish_day", result.finish_day),
        ("met_deadline", result.met_deadline),
        ("total_planned_hours", round(result.total_planned_hours, 2)),
        ("completed_hours_before_run", round(result.completed_hours_before_run, 2)),
        ("remaining_hours_at_start", round(result.remaining_hours_at_start, 2)),
        ("total_used_hours", round(result.total_used_hours, 2)),
        ("remaining_backlog_hours", round(result.remaining_backlog_hours, 2)),
        ("recommendation_status", result.recommendation.status_label),
        ("bottleneck_type", result.recommendation.bottleneck_type),
        ("available_backlog_hours", round(result.recommendation.available_backlog_hours, 2)),
        ("blocked_backlog_hours", round(result.recommendation.blocked_backlog_hours, 2)),
        ("total_backlog_hours", round(result.recommendation.total_backlog_hours, 2)),
        ("capacity_to_deadline_hours", round(result.recommendation.capacity_to_deadline_hours, 2)),
        ("extra_staff_days_needed", round(result.recommendation.extra_staff_days_needed, 2)),
        ("recommended_action", result.recommendation.recommended_action),
    ]
    ws.append(["field", "value"])
    for row in summary_rows:
        ws.append(list(row))
    _bold_header(ws)

    # Days
    ws = wb.create_sheet("Days")
    ws.append(
        [
            "day",
            "active_school_name",
            "status_note",
            "effective_staff",
            "general_staff",
            "carpet_staff",
            "general_capacity",
            "carpet_capacity",
            "daily_capacity",
            "general_used_capacity",
            "carpet_used_capacity",
            "used_capacity",
            "general_unused_capacity",
            "carpet_unused_capacity",
            "unused_capacity",
            "work_items",
        ]
    )
    _bold_header(ws)
    for day in result.days:
        ws.append(
            [
                day.day,
                day.active_school_name,
                day.status_note,
                day.effective_staff,
                day.general_staff,
                day.carpet_staff,
                round(day.general_capacity, 2),
                round(day.carpet_capacity, 2),
                round(day.daily_capacity, 2),
                round(day.general_used_capacity, 2),
                round(day.carpet_used_capacity, 2),
                round(day.used_capacity, 2),
                round(day.general_unused_capacity, 2),
                round(day.carpet_unused_capacity, 2),
                round(day.unused_capacity, 2),
                len(day.work_log),
            ]
        )

    # Work Log
    ws = wb.create_sheet("WorkLog")
    ws.append(
        [
            "day",
            "crew_type",
            "school_name",
            "building_name",
            "zone_name",
            "room_name",
            "phase_name",
            "available_day",
            "hours_done",
            "note",
        ]
    )
    _bold_header(ws)
    for day in result.days:
        for item in day.work_log:
            ws.append(
                [
                    day.day,
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

    # Tasks
    ws = wb.create_sheet("Tasks")
    ws.append(
        [
            "school_name",
            "building_name",
            "zone_name",
            "room_name",
            "phase_name",
            "available_day",
            "total_hours",
            "remaining_hours",
            "status",
            "notes",
        ]
    )
    _bold_header(ws)
    for task in result.task_items:
        if task.remaining_hours <= 0:
            status = "Complete"
        elif task.remaining_hours < task.total_hours:
            status = "In Progress"
        else:
            status = "Not Started"

        ws.append(
            [
                task.school_name,
                task.building_name,
                task.zone_name,
                task.room_name,
                task.phase_name,
                task.available_day,
                round(task.total_hours, 2),
                round(task.remaining_hours, 2),
                status,
                task.notes,
            ]
        )

    wb.save(file_path)
    return file_path


def create_input_template(file_path: str = "data/summer_scheduler_template.xlsx") -> str:
    Path(file_path).parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Setup
    ws = wb.active
    ws.title = "Setup"
    ws.append(["setting", "value", "description"])
    setup_rows = [
        ("schedule_name", "Summer Schedule", "Name shown in the app and exports"),
        ("current_day", 1, "Usually 1 for a fresh run"),
        ("target_end_day", 20, "Target workday number"),
        ("scheduled_shift_hours_per_day", 8.5, "Paid shift length"),
        ("lunch_hours_per_day", 0.5, "Lunch time"),
        ("break_hours_per_day", 0.5, "Total break time"),
        ("setup_hours_per_day", 0.25, "Setup/opening time"),
        ("cleanup_hours_per_day", 0.5, "Cleanup/lockup time"),
        ("productive_hours_per_staff_per_day", 6.75, "Real productive time per worker"),
        ("include_deep_clean", True, "Run deep clean tasks"),
        ("include_strip", True, "Run strip tasks"),
        ("include_wax", True, "Run wax tasks"),
        ("include_carpet", True, "Run carpet tasks"),
        ("include_exterior", False, "Run exterior tasks"),
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
    for row in setup_rows:
        ws.append(list(row))
    _bold_header(ws)

    # Rooms
    ws = wb.create_sheet("Rooms")
    ws.append(
        [
            "school_name",
            "school_order",
            "building_name",
            "zone_name",
            "room_name",
            "room_order",
            "carpet_sqft",
            "tile_sqft",
            "available_day",
            "include_deep_clean",
            "include_strip",
            "include_wax",
            "include_carpet",
            "include_exterior",
            "notes",
        ]
    )
    _bold_header(ws)
    ws.append(
        [
            "WES",
            1,
            "Main",
            "Classrooms",
            "Room 1",
            1,
            574,
            296,
            1,
            True,
            True,
            True,
            True,
            False,
            "",
        ]
    )

    # Staffing
    ws = wb.create_sheet("Staffing")
    ws.append(
        [
            "day",
            "available_staff",
            "carpet_staff_reserved",
            "absences",
            "temporary_help",
        ]
    )
    _bold_header(ws)
    for day in range(1, 21):
        ws.append([day, 4, 0, 0, 0])

    # Progress
    ws = wb.create_sheet("Progress")
    ws.append(
        [
            "school_name",
            "building_name",
            "zone_name",
            "room_name",
            "phase_name",
            "hours_completed",
        ]
    )
    _bold_header(ws)

    wb.save(file_path)
    return file_path