from data_loader import load_rooms, load_settings, load_staffing
from exporter import export_daily_schedule, export_summary, export_task_summary
from scheduler import run_scheduler, summarize_hours_by_phase, summarize_hours_by_school


def main():
    settings = load_settings("data/settings.csv")
    rooms, schools = load_rooms("data/rooms.csv")
    staffing_days = load_staffing("data/staffing.csv")

    result = run_scheduler(
        rooms=rooms,
        schools=schools,
        settings=settings,
        staffing_days=staffing_days,
    )

    print("\n--- SCHEDULE INFO ---")
    print(f"Schedule name: {settings.schedule_name}")
    print(f"Target end day: {settings.target_end_day}")
    print(f"Hours per staff per day: {settings.hours_per_staff_per_day:.2f}")

    print("\n--- ACTIVE PHASES ---")
    print(f"Deep Clean: {'ON' if settings.include_deep_clean else 'OFF'}")
    print(f"Strip: {'ON' if settings.include_strip else 'OFF'}")
    print(f"Wax: {'ON' if settings.include_wax else 'OFF'}")
    print(f"Carpet: {'ON' if settings.include_carpet else 'OFF'}")
    print(f"Exterior: {'ON' if settings.include_exterior else 'OFF'}")

    print("\n--- SCHOOL ORDER ---")
    for school in schools:
        print(f"{school.order}. {school.name}")

    print("\n--- HOURS BY SCHOOL ---")
    by_school = summarize_hours_by_school(result.task_items)
    for school_name, hours in by_school.items():
        print(f"{school_name}: {hours:.2f} hrs")

    print("\n--- HOURS BY PHASE ---")
    by_phase = summarize_hours_by_phase(result.task_items)
    for phase_name, hours in by_phase.items():
        print(f"{phase_name}: {hours:.2f} hrs")

    print("\n--- RESULT ---")
    print(f"Total planned hours: {result.total_planned_hours:.2f}")
    print(f"Finish day: {result.finish_day}")
    print(f"Deadline met: {'YES' if result.met_deadline else 'NO'}")
    print(f"Remaining backlog hours: {result.remaining_backlog_hours:.2f}")

    print("\n--- DAILY SCHEDULE ---")
    for day in result.days:
        print(
            f"\nDay {day.day} | staff {day.effective_staff} | "
            f"capacity {day.daily_capacity:.2f} | used {day.used_capacity:.2f} | "
            f"unused {day.unused_capacity:.2f}"
        )

        if not day.work_log:
            print("  No work scheduled")
            continue

        for item in day.work_log:
            print(
                f"  {item.school_name} | {item.room_name} | "
                f"{item.phase_name} | {item.hours_done:.2f} hrs"
            )

    daily_path = export_daily_schedule(result)
    task_path = export_task_summary(result)
    summary_path = export_summary(result)

    print("\n--- EXPORTED FILES ---")
    print(daily_path)
    print(task_path)
    print(summary_path)


if __name__ == "__main__":
    main()