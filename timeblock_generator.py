from __future__ import annotations

from datetime import datetime, timedelta
from typing import List, Dict, Any

from models import ScheduleDayResult, ScheduleSettings


TIME_FMT = "%I:%M %p"


def _parse_time(value: str) -> datetime:
    return datetime.strptime(value.strip(), TIME_FMT)


def _fmt_time(dt: datetime) -> str:
    return dt.strftime("%I:%M %p").lstrip("0")


def _minutes_between(start: datetime, end: datetime) -> int:
    return max(0, int((end - start).total_seconds() // 60))


def generate_time_blocks(
    day: ScheduleDayResult,
    settings: ScheduleSettings,
) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []

    if not day.work_log:
        return [
            {
                "type": "idle",
                "start": settings.work_start_time,
                "end": settings.day_end_time,
                "label": "No scheduled work",
                "minutes": 0,
            }
        ]

    cursor = _parse_time(settings.work_start_time)

    first_break_time = _parse_time(settings.first_break_time)
    lunch_time = _parse_time(settings.lunch_time)
    second_break_time = _parse_time(settings.second_break_time)

    first_break_minutes = int((settings.break_hours_per_day / 2) * 60)
    lunch_minutes = int(settings.lunch_hours_per_day * 60)
    second_break_minutes = int((settings.break_hours_per_day / 2) * 60)

    events = [
        {
            "name": "Break",
            "time": first_break_time,
            "minutes": first_break_minutes,
            "used": False,
        },
        {
            "name": "Lunch",
            "time": lunch_time,
            "minutes": lunch_minutes,
            "used": False,
        },
        {
            "name": "Break",
            "time": second_break_time,
            "minutes": second_break_minutes,
            "used": False,
        },
    ]

    def insert_due_events() -> None:
        nonlocal cursor
        for event in events:
            if not event["used"] and cursor >= event["time"]:
                start = cursor
                end = cursor + timedelta(minutes=event["minutes"])
                blocks.append(
                    {
                        "type": event["name"].lower(),
                        "start": _fmt_time(start),
                        "end": _fmt_time(end),
                        "label": event["name"],
                        "minutes": event["minutes"],
                    }
                )
                cursor = end
                event["used"] = True

    for item in day.work_log:
        remaining_minutes = max(1, int(round(item.hours_done * 60)))

        task_label_parts = [item.phase_name]
        if item.room_name and item.room_name != "TRANSITION":
            task_label_parts.append(item.room_name)
        if item.school_name and item.school_name != "MULTI-SCHOOL":
            task_label_parts.append(f"({item.school_name})")

        task_label = " - ".join(task_label_parts[:2])
        if len(task_label_parts) > 2:
            task_label += f" {task_label_parts[2]}"

        while remaining_minutes > 0:
            insert_due_events()

            next_event = None
            for event in events:
                if not event["used"]:
                    next_event = event
                    break

            if next_event is None:
                segment_minutes = remaining_minutes
            else:
                minutes_until_event = _minutes_between(cursor, next_event["time"])
                if minutes_until_event <= 0:
                    insert_due_events()
                    continue
                segment_minutes = min(remaining_minutes, minutes_until_event)

            start = cursor
            end = cursor + timedelta(minutes=segment_minutes)

            blocks.append(
                {
                    "type": "work",
                    "start": _fmt_time(start),
                    "end": _fmt_time(end),
                    "label": task_label,
                    "minutes": segment_minutes,
                    "crew_type": item.crew_type,
                    "school_name": item.school_name,
                    "room_name": item.room_name,
                    "phase_name": item.phase_name,
                    "note": item.note or "",
                }
            )

            cursor = end
            remaining_minutes -= segment_minutes

    insert_due_events()

    return blocks


def format_time_blocks_for_text(
    day: ScheduleDayResult,
    settings: ScheduleSettings,
) -> str:
    blocks = generate_time_blocks(day, settings)

    lines: List[str] = []
    for block in blocks:
        lines.append(f'{block["start"]} - {block["end"]}  {block["label"]}')

    return "\n".join(lines)