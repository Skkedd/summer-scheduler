from models import Room, School, SummerSchedule


def build_sample_schedule():
    school_a = School("Elementary A")
    school_b = School("Elementary B")
    school_c = School("Elementary C")

    # School A rooms
    school_a.add_room(Room("Room 1", "Elementary A", 800, 4))
    school_a.add_room(Room("Room 2", "Elementary A", 900, 5))
    school_a.add_room(Room("Room 3", "Elementary A", 700, 3.5))
    school_a.add_room(Room("Office", "Elementary A", 1200, 6))

    # School B rooms
    school_b.add_room(Room("Room 1", "Elementary B", 850, 4.5))
    school_b.add_room(Room("Room 2", "Elementary B", 950, 5))
    school_b.add_room(Room("Library", "Elementary B", 2000, 10))

    # School C rooms
    school_c.add_room(Room("Room 1", "Elementary C", 800, 4, available_day=10))
    school_c.add_room(Room("Room 2", "Elementary C", 800, 4, available_day=10))
    school_c.add_room(Room("Multi-use", "Elementary C", 3000, 12))

    schedule = SummerSchedule(
        "Summer 2026",
        work_days=20,
        staff_count=4
    )

    schedule.add_school(school_a)
    schedule.add_school(school_b)
    schedule.add_school(school_c)

    return schedule