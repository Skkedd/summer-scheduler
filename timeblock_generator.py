from __future__ import annotations

from datetime import datetime, timedelta
from typing import List, Dict, Any

from models import ScheduleDayResult, ScheduleSettings


TIME_FMT = "%I:%M %p"


def _parse_time(value: str) -> datetime:
    return datetime.strptime(value.strip(), TIME_FMT)


def _fmt_time(dt: datetime) -> str:
    return dt.strftime("%I:%M %p")


def _minutes_between(start: datetime, end: datetime) -> int:
    return max(0, int((end - start).total_seconds() // 60))


def _is_strip_wax(label: str) -> bool:
    return label == "Strip" or label.startswith("Wax")


def _clock_minutes_for_work(item, day: ScheduleDayResult) -> int:
    """
    Convert labor-hours into displayed clock-time.

    Deep clean/general work uses the effective day staff because the crew can work
    as a group or spread across nearby rooms. Strip/wax is modeled as 2-person
    crews, because that better matches the real floor process.
    """
    if _is_strip_wax(item.phase_name):
        staff = max(1, min(2, day.general_staff or day.effective_staff or 1))
    else:
        staff = max(1, day.effective_staff or day.general_staff or 1)

    return max(1, int(round((item.hours_done / staff) * 60)))


def generate_time_blocks(
    day: ScheduleDayResult,
    settings: ScheduleSettings,
) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []

    if not day.work_log:
        label = "No scheduled work"
        if "ACCESS DOWNTIME WARNING" in (day.status_note or ""):
            label = day.status_note
        return [
            {
                "type": "idle",
                "start": _fmt_time(_parse_time(settings.work_start_time)),
                "end": _fmt_time(_parse_time(settings.day_end_time)),
                "label": label,
                "minutes": 0,
            }
        ]

    cursor = _parse_time(settings.work_start_time)
    work_stop_time = _parse_time(settings.cleanup_start_time)
    day_end_time = _parse_time(settings.day_end_time)

    first_break_time = _parse_time(settings.first_break_time)
    lunch_time = _parse_time(settings.lunch_time)
    second_break_time = _parse_time(settings.second_break_time)

    first_break_minutes = int((settings.break_hours_per_day / 2) * 60)
    lunch_minutes = int(settings.lunch_hours_per_day * 60)
    second_break_minutes = int((settings.break_hours_per_day / 2) * 60)

    events = [
        {"name": "Break", "time": first_break_time, "minutes": first_break_minutes, "used": False},
        {"name": "Lunch", "time": lunch_time, "minutes": lunch_minutes, "used": False},
        {"name": "Break", "time": second_break_time, "minutes": second_break_minutes, "used": False},
    ]

    def insert_due_events() -> None:
        nonlocal cursor
        for event in events:
            if not event["used"] and cursor >= event["time"] and cursor < work_stop_time:
                start = cursor
                end = min(cursor + timedelta(minutes=event["minutes"]), work_stop_time)
                blocks.append(
                    {
                        "type": event["name"].lower(),
                        "start": _fmt_time(start),
                        "end": _fmt_time(end),
                        "label": event["name"],
                        "minutes": _minutes_between(start, end),
                    }
                )
                cursor = end
                event["used"] = True

    overflow_minutes = 0

    for item in day.work_log:
        remaining_minutes = _clock_minutes_for_work(item, day)

        task_label_parts = [item.phase_name]
        if item.room_name and item.room_name != "TRANSITION":
            task_label_parts.append(item.room_name)
        if item.school_name and item.school_name != "MULTI-SCHOOL":
            task_label_parts.append(f"({item.school_name})")

        task_label = " - ".join(task_label_parts[:2])
        if len(task_label_parts) > 2:
            task_label += f" {task_label_parts[2]}"

        if _is_strip_wax(item.phase_name):
            task_label += " [2-person crew]"

        while remaining_minutes > 0:
            insert_due_events()

            if cursor >= work_stop_time:
                overflow_minutes += remaining_minutes
                remaining_minutes = 0
                break

            next_event = None
            for event in events:
                if not event["used"]:
                    next_event = event
                    break

            if next_event is None:
                segment_minutes = min(remaining_minutes, _minutes_between(cursor, work_stop_time))
            else:
                minutes_until_event = _minutes_between(cursor, min(next_event["time"], work_stop_time))
                if minutes_until_event <= 0:
                    insert_due_events()
                    continue
                segment_minutes = min(remaining_minutes, minutes_until_event)

            if segment_minutes <= 0:
                overflow_minutes += remaining_minutes
                break

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

    if cursor < day_end_time:
        blocks.append(
            {
                "type": "cleanup",
                "start": _fmt_time(max(cursor, work_stop_time)),
                "end": _fmt_time(day_end_time),
                "label": "Cleanup / lockup",
                "minutes": _minutes_between(max(cursor, work_stop_time), day_end_time),
            }
        )

    if "ACCESS DOWNTIME WARNING" in (day.status_note or ""):
        blocks.append(
            {
                "type": "warning",
                "start": _fmt_time(work_stop_time),
                "end": _fmt_time(day_end_time),
                "label": day.status_note,
                "minutes": 0,
            }
        )

    if overflow_minutes > 0:
        blocks.append(
            {
                "type": "warning",
                "start": _fmt_time(work_stop_time),
                "end": _fmt_time(day_end_time),
                "label": f"Warning: {overflow_minutes} displayed minute(s) did not fit before cleanup",
                "minutes": overflow_minutes,
            }
        )

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
