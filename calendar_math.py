from __future__ import annotations

from datetime import datetime, timedelta, date
from typing import Iterable, Optional, Set

try:
    from openpyxl.utils.datetime import from_excel
except Exception:  # pragma: no cover
    from_excel = None


def parse_date_string(value) -> date:
    """Parse workbook/user date values.

    Accepts normal strings, Excel/openpyxl date/datetime objects and Excel serial dates.
    """
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if from_excel is not None and value > 20000:
            converted = from_excel(value)
            return converted.date() if isinstance(converted, datetime) else converted

    raw = (str(value or "")).strip()
    if not raw:
        raise ValueError("Date value is blank")

    if " " in raw:
        raw = raw.split(" ", 1)[0]

    try:
        numeric = float(raw)
        if from_excel is not None and numeric > 20000:
            converted = from_excel(numeric)
            return converted.date() if isinstance(converted, datetime) else converted
    except Exception:
        pass

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


def normalize_holidays(holidays: Optional[Iterable[date]]) -> Set[date]:
    return {parse_date_string(item) for item in (holidays or set())}


def is_workday(
    value: date,
    work_on_weekends: bool = False,
    holidays: Optional[Iterable[date]] = None,
) -> bool:
    holiday_dates = normalize_holidays(holidays)
    if not work_on_weekends and value.weekday() >= 5:
        return False
    if value in holiday_dates:
        return False
    return True


def normalize_start_date(
    start: date,
    work_on_weekends: bool,
    holidays: Optional[Iterable[date]] = None,
) -> date:
    while not is_workday(start, work_on_weekends=work_on_weekends, holidays=holidays):
        start += timedelta(days=1)
    return start


def count_workdays(
    start_date_str: str,
    end_date_str: str,
    work_on_weekends: bool = False,
    holidays: Optional[Iterable[date]] = None,
) -> int:
    start_date = parse_date_string(start_date_str)
    end_date = parse_date_string(end_date_str)

    if end_date < start_date:
        return 0

    count = 0
    current = start_date
    while current <= end_date:
        if is_workday(current, work_on_weekends=work_on_weekends, holidays=holidays):
            count += 1
        current += timedelta(days=1)

    return count


def workday_to_date(
    start_date_str: str,
    workday_number: int,
    work_on_weekends: bool = False,
    holidays: Optional[Iterable[date]] = None,
) -> date:
    if workday_number <= 0:
        raise ValueError("Workday number must be 1 or greater")

    current = normalize_start_date(
        parse_date_string(start_date_str),
        work_on_weekends=work_on_weekends,
        holidays=holidays,
    )

    counted = 0
    while True:
        if is_workday(current, work_on_weekends=work_on_weekends, holidays=holidays):
            counted += 1
            if counted == workday_number:
                return current
        current += timedelta(days=1)


def format_date_label(value: date) -> str:
    return value.strftime("%A %B %d, %Y")


def format_workday_label(
    start_date_str: str,
    workday_number: int,
    work_on_weekends: bool = False,
    holidays: Optional[Iterable[date]] = None,
) -> str:
    day_date = workday_to_date(
        start_date_str=start_date_str,
        workday_number=workday_number,
        work_on_weekends=work_on_weekends,
        holidays=holidays,
    )
    return f"Day {workday_number} - {format_date_label(day_date)}"
