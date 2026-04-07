from sample_data import build_sample_schedule


def main():
    schedule = build_sample_schedule()

    print("\n--- School Breakdown ---")
    schedule.school_breakdown()

    print(f"Total hours: {schedule.total_hours():.2f}")
    print(f"Daily capacity: {schedule.daily_capacity():.2f}")
    print(f"Projected days needed: {schedule.projected_days_needed():.2f}")

    if schedule.will_meet_deadline():
        print("Status: ON TRACK")
        print(f"Days ahead: {schedule.days_over_under():.2f}")
    else:
        print("Status: BEHIND")
        print(f"Days behind: {abs(schedule.days_over_under()):.2f}")

    print("\n--- Availability Check ---")
    for day in [1, 5, 10]:
        available = schedule.available_hours_by_day(day)
        print(f"Day {day}: {available:.2f} hours available")

    print("\n--- Schedule Simulation ---")
    simulation = schedule.simulate_schedule()

    print(f"Projected finish day: Day {simulation['finish_day']}")

    for entry in simulation["daily_log"]:
        print(f"\nDay {entry['day']}")
        if entry["work_done"]:
            for item in entry["work_done"]:
                print(f"  {item}")
        else:
            print("  No work completed")
        print(f"  Unused capacity: {entry['unused_capacity']:.2f} hrs")


if __name__ == "__main__":
    main()