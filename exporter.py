import csv
import os

from models import ScheduleResult


def ensure_output_folder(folder_path: str = "output") -> None:
    os.makedirs(folder_path, exist_ok=True)


def export_daily_schedule(result: ScheduleResult, folder_path: str = "output") -> str:
    ensure_output_folder(folder_path)
    file_path = os.path.join(folder_path, "daily_schedule.csv")

    with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(
            [
                "day",
                "effective_staff",
                "general_staff",
                "carpet_staff",
                "general_capacity_hours",
                "carpet_capacity_hours",
                "daily_capacity_productive_hours",
                "general_used_capacity_hours",
                "carpet_used_capacity_hours",
                "used_capacity_hours",
                "general_unused_capacity_hours",
                "carpet_unused_capacity_hours",
                "unused_capacity_hours",
                "active_school_name",
                "day_status_note",
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

        for day in result.days:
            if day.work_log:
                for item in day.work_log:
                    writer.writerow(
                        [
                            day.day,
                            day.effective_staff,
                            day.general_staff,
                            day.carpet_staff,
                            f"{day.general_capacity:.2f}",
                            f"{day.carpet_capacity:.2f}",
                            f"{day.daily_capacity:.2f}",
                            f"{day.general_used_capacity:.2f}",
                            f"{day.carpet_used_capacity:.2f}",
                            f"{day.used_capacity:.2f}",
                            f"{day.general_unused_capacity:.2f}",
                            f"{day.carpet_unused_capacity:.2f}",
                            f"{day.unused_capacity:.2f}",
                            day.active_school_name,
                            day.status_note,
                            item.crew_type,
                            item.school_name,
                            item.building_name,
                            item.zone_name,
                            item.room_name,
                            item.phase_name,
                            item.available_day if item.available_day is not None else "",
                            f"{item.hours_done:.2f}",
                            item.note,
                        ]
                    )
            else:
                writer.writerow(
                    [
                        day.day,
                        day.effective_staff,
                        day.general_staff,
                        day.carpet_staff,
                        f"{day.general_capacity:.2f}",
                        f"{day.carpet_capacity:.2f}",
                        f"{day.daily_capacity:.2f}",
                        f"{day.general_used_capacity:.2f}",
                        f"{day.carpet_used_capacity:.2f}",
                        f"{day.used_capacity:.2f}",
                        f"{day.general_unused_capacity:.2f}",
                        f"{day.carpet_unused_capacity:.2f}",
                        f"{day.unused_capacity:.2f}",
                        day.active_school_name,
                        day.status_note,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                    ]
                )

    return file_path


def export_task_summary(result: ScheduleResult, folder_path: str = "output") -> str:
    ensure_output_folder(folder_path)
    file_path = os.path.join(folder_path, "task_summary.csv")

    with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(
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

        for task in result.task_items:
            if task.remaining_hours <= 0:
                status = "Complete"
            elif task.remaining_hours < task.total_hours:
                status = "In Progress"
            else:
                status = "Not Started"

            writer.writerow(
                [
                    task.school_name,
                    task.building_name,
                    task.zone_name,
                    task.room_name,
                    task.phase_name,
                    task.available_day,
                    f"{task.total_hours:.2f}",
                    f"{task.remaining_hours:.2f}",
                    status,
                    task.notes,
                ]
            )

    return file_path


def export_summary(result: ScheduleResult, folder_path: str = "output") -> str:
    ensure_output_folder(folder_path)
    file_path = os.path.join(folder_path, "schedule_summary.csv")

    with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["schedule_name", result.schedule_name])
        writer.writerow(["target_end_day", result.target_end_day])
        writer.writerow(["current_day", result.current_day])
        writer.writerow(["finish_day", result.finish_day])
        writer.writerow(["met_deadline", result.met_deadline])
        writer.writerow(["total_planned_hours", f"{result.total_planned_hours:.2f}"])
        writer.writerow(
            ["completed_hours_before_run", f"{result.completed_hours_before_run:.2f}"]
        )
        writer.writerow(
            ["remaining_hours_at_start", f"{result.remaining_hours_at_start:.2f}"]
        )
        writer.writerow(
            ["total_used_hours_in_reforecast", f"{result.total_used_hours:.2f}"]
        )
        writer.writerow(
            ["remaining_backlog_hours", f"{result.remaining_backlog_hours:.2f}"]
        )
        writer.writerow(
            ["recommendation_based_on_start_of_run", result.recommendation.based_on_start_of_run]
        )
        writer.writerow(["status_label", result.recommendation.status_label])
        writer.writerow(["bottleneck_type", result.recommendation.bottleneck_type])
        writer.writerow(
            [
                "available_backlog_hours_at_start",
                f"{result.recommendation.available_backlog_hours:.2f}",
            ]
        )
        writer.writerow(
            [
                "blocked_backlog_hours_at_start",
                f"{result.recommendation.blocked_backlog_hours:.2f}",
            ]
        )
        writer.writerow(
            [
                "total_backlog_hours_at_start",
                f"{result.recommendation.total_backlog_hours:.2f}",
            ]
        )
        writer.writerow(
            [
                "productive_capacity_to_deadline_hours",
                f"{result.recommendation.capacity_to_deadline_hours:.2f}",
            ]
        )
        writer.writerow(
            [
                "extra_staff_days_needed",
                f"{result.recommendation.extra_staff_days_needed:.2f}",
            ]
        )
        writer.writerow(["recommended_action", result.recommendation.recommended_action])

    return file_path