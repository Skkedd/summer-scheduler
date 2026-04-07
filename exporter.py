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
                "daily_capacity",
                "used_capacity",
                "unused_capacity",
                "school_name",
                "room_name",
                "phase_name",
                "hours_done",
            ]
        )

        for day in result.days:
            if day.work_log:
                for item in day.work_log:
                    writer.writerow(
                        [
                            day.day,
                            day.effective_staff,
                            f"{day.daily_capacity:.2f}",
                            f"{day.used_capacity:.2f}",
                            f"{day.unused_capacity:.2f}",
                            item.school_name,
                            item.room_name,
                            item.phase_name,
                            f"{item.hours_done:.2f}",
                        ]
                    )
            else:
                writer.writerow(
                    [
                        day.day,
                        day.effective_staff,
                        f"{day.daily_capacity:.2f}",
                        f"{day.used_capacity:.2f}",
                        f"{day.unused_capacity:.2f}",
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
            status = "Complete" if task.remaining_hours <= 0 else "Remaining"
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
        writer.writerow(["finish_day", result.finish_day])
        writer.writerow(["met_deadline", result.met_deadline])
        writer.writerow(["total_planned_hours", f"{result.total_planned_hours:.2f}"])
        writer.writerow(["total_used_hours", f"{result.total_used_hours:.2f}"])
        writer.writerow(["remaining_backlog_hours", f"{result.remaining_backlog_hours:.2f}"])

    return file_path