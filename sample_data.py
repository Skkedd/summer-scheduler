import csv
from models import Room, School, SummerSchedule


CARPET_RATE = 500
TILE_RATE = 300


def read_settings(settings_file_path="data/settings.csv"):
    settings = {}

    with open(settings_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)

        for row in reader:
            key = row["setting"].strip()
            value = row["value"].strip()
            settings[key] = value

    return {
        "schedule_name": settings.get("schedule_name", "Summer Schedule"),
        "work_days": int(settings.get("work_days", 20)),
        "staff_count": int(settings.get("staff_count", 4)),
        "include_carpet": settings.get("include_carpet", "True").lower() == "true",
        "include_tile": settings.get("include_tile", "True").lower() == "true",
    }


def calculate_hours(carpet_sqft, tile_sqft, include_carpet=True, include_tile=True):
    carpet_hours = (carpet_sqft / CARPET_RATE) if include_carpet else 0
    tile_hours = (tile_sqft / TILE_RATE) if include_tile else 0
    return carpet_hours + tile_hours


def build_schedule_from_csv(
    rooms_file_path="data/rooms.csv",
    settings_file_path="data/settings.csv",
):
    settings = read_settings(settings_file_path)
    schools_by_name = {}

    with open(rooms_file_path, mode="r", newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)

        for row in reader:
            school_name = row["school_name"].strip()
            school_order = int(row["school_order"])
            room_name = row["room_name"].strip()
            room_order = int(row["room_order"])
            carpet_sqft = float(row["carpet_sqft"])
            tile_sqft = float(row["tile_sqft"])
            available_day = int(row["available_day"])

            estimated_hours = calculate_hours(
                carpet_sqft=carpet_sqft,
                tile_sqft=tile_sqft,
                include_carpet=settings["include_carpet"],
                include_tile=settings["include_tile"],
            )

            if estimated_hours <= 0:
                continue

            if school_name not in schools_by_name:
                schools_by_name[school_name] = School(
                    name=school_name,
                    order=school_order,
                )

            room = Room(
                name=room_name,
                school=school_name,
                sqft=carpet_sqft + tile_sqft,
                estimated_hours=estimated_hours,
                available_day=available_day,
                order=room_order,
            )

            schools_by_name[school_name].add_room(room)

    schedule = SummerSchedule(
        name=settings["schedule_name"],
        work_days=settings["work_days"],
        staff_count=settings["staff_count"],
    )

    for school in schools_by_name.values():
        schedule.add_school(school)

    return schedule, settings