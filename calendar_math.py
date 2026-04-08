from __future__ import annotations

from datetime import datetime, timedelta, date


def parse_date_string(value: str) -> date:
    raw = (value or "").strip()
    if not raw:
        raise ValueError("Date value is blank")

    formats = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%m/%d/%y",
        "%m-%d-%y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue

    raise ValueError(
        f"Could not parse date '{value}'. Use YYYY-MM-DD or MM/DD/YYYY."
    )


def normalize_start_date(start: date, work_on_weekends: bool) -> date:
    if work_on_weekends:
        return start

    while start.weekday() >= 5:
        start += timedelta(days=1)

    return start


def workday_to_date(
    start_date_str: str,
    workday_number: int,
    work_on_weekends: bool = False,
) -> date:
    if workday_number <= 0:
        raise ValueError("Workday number must be 1 or greater")

    current = normalize_start_date(
        parse_date_string(start_date_str),
        work_on_weekends=work_on_weekends,
    )

    if workday_number == 1:
        return current

    counted = 1
    while counted < workday_number:
        current += timedelta(days=1)

        if work_on_weekends or current.weekday() < 5:
            counted += 1

    return current


def format_date_label(value: date) -> str:
    return value.strftime("%A %B %d, %Y")


def format_workday_label(
    start_date_str: str,
    workday_number: int,
    work_on_weekends: bool = False,
) -> str:
    day_date = workday_to_date(
        start_date_str=start_date_str,
        workday_number=workday_number,
        work_on_weekends=work_on_weekends,
    )
    return f"Day {workday_number} - {format_date_label(day_date)}"