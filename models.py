class Room:
    def __init__(self, name, school, sqft, estimated_hours, available_day=1):
        self.name = name
        self.school = school
        self.sqft = sqft
        self.estimated_hours = estimated_hours
        self.available_day = available_day
        self.remaining_hours = estimated_hours


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
        return self.staff_count * 8

    def projected_days_needed(self):
        return self.total_hours() / self.daily_capacity()

    def will_meet_deadline(self):
        return self.projected_days_needed() <= self.work_days

    def days_over_under(self):
        return self.work_days - self.projected_days_needed()

    def school_breakdown(self):
        for school in self.schools:
            print(f"{school.name}: {school.total_hours():.2f} hours")

    def available_hours_by_day(self, day):
        total = 0
        for school in self.schools:
            for room in school.rooms:
                if room.available_day <= day:
                    total += room.estimated_hours
        return total

    def reset_progress(self):
        for school in self.schools:
            for room in school.rooms:
                room.remaining_hours = room.estimated_hours

    def simulate_schedule(self):
        self.reset_progress()

        day = 1
        daily_log = []

        while True:
            remaining_total = sum(
                room.remaining_hours
                for school in self.schools
                for room in school.rooms
            )

            if remaining_total <= 0:
                break

            capacity = self.daily_capacity()
            worked_today = []

            for school in self.schools:
                for room in school.rooms:
                    if room.available_day <= day and room.remaining_hours > 0 and capacity > 0:
                        hours_done = min(room.remaining_hours, capacity)
                        room.remaining_hours -= hours_done
                        capacity -= hours_done

                        worked_today.append(
                            f"{school.name} - {room.name}: {hours_done:.2f} hrs"
                        )

            daily_log.append({
                "day": day,
                "work_done": worked_today,
                "unused_capacity": capacity
            })

            day += 1

            if day > self.work_days + 100:
                break

        return {
            "finish_day": day - 1,
            "daily_log": daily_log
        }