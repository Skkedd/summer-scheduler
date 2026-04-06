# main.py

class Room:
    def __init__(self, name, school, sqft, estimated_hours):
        self.name = name
        self.school = school
        self.sqft = sqft
        self.estimated_hours = estimated_hours


class School:
    def __init__(self, name):
        self.name = name
        self.rooms = []

    def add_room(self, room):
        self.rooms.append(room)

    def total_hours(self):
        return sum(room.estimated_hours for room in self.rooms)


class SummerSchedule:
    def __init__(self, name, work_days, staff_count):
        self.name = name
        self.work_days = work_days
        self.staff_count = staff_count
        self.schools = []

    def add_school(self, school):
        self.schools.append(school)

    def total_hours(self):
        return sum(school.total_hours() for school in self.schools)

    def daily_capacity(self):
        return self.staff_count * 8  # 8 hours per person

    def projected_days_needed(self):
        return self.total_hours() / self.daily_capacity()


# ----- TEST DATA -----

school_a = School("School A")

school_a.add_room(Room("Room 1", "School A", 800, 4))
school_a.add_room(Room("Room 2", "School A", 900, 5))
school_a.add_room(Room("Room 3", "School A", 700, 3.5))

schedule = SummerSchedule("Summer 2026", work_days=40, staff_count=5)
schedule.add_school(school_a)

# ----- OUTPUT -----

print(f"Total hours: {schedule.total_hours():.2f}")
print(f"Daily capacity: {schedule.daily_capacity():.2f}")
print(f"Projected days needed: {schedule.projected_days_needed():.2f}")