"""Time parsing helpers for survey intervals."""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta
from typing import Any

import pandas as pd


def parse_time(value: Any) -> time | None:
    """Parse common Excel/Pandas/string time values into a time."""
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, pd.Timestamp):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, (int, float)):
        if 0 <= float(value) < 1:
            seconds = int(round(float(value) * 24 * 60 * 60))
            return (datetime.min + timedelta(seconds=seconds)).time()
        return None

    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"(\d{1,2})[:.](\d{2})", text)
    if not match:
        return None
    hour = int(match.group(1))
    minute = int(match.group(2))
    if 0 <= hour <= 23 and 0 <= minute <= 59:
        return time(hour, minute)
    return None


def parse_interval(value: Any) -> tuple[time | None, time | None]:
    """Parse strings like 07:00-07:15 into start/end times."""
    if pd.isna(value):
        return None, None
    text = str(value)
    matches = re.findall(r"(\d{1,2})[:.](\d{2})", text)
    if len(matches) >= 2:
        start = time(int(matches[0][0]), int(matches[0][1]))
        end = time(int(matches[1][0]), int(matches[1][1]))
        return start, end
    return parse_time(value), None


def add_minutes(value: time | None, minutes: int) -> time | None:
    if value is None:
        return None
    base = datetime.combine(date.today(), value)
    return (base + timedelta(minutes=minutes)).time()


def time_to_minutes(value: Any) -> int | None:
    parsed = parse_time(value)
    if parsed is None:
        return None
    return parsed.hour * 60 + parsed.minute


def minutes_to_time(minutes: int) -> time:
    minutes = minutes % (24 * 60)
    return time(minutes // 60, minutes % 60)


SUMMARY_TIME_LABELS = {"total", "\u0e23\u0e27\u0e21", "\xe0\xb8\xa3\xe0\xb8\xa7\xe0\xb8\xa1"}


def _blankish(value: Any) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def normalize_interval_label(value: Any) -> str:
    if _blankish(value):
        return ""
    return str(value).strip().replace(".", ":")


def is_summary_time_label(value: Any) -> bool:
    label = normalize_interval_label(value).casefold()
    return not label or label in SUMMARY_TIME_LABELS


def hourly_interval_label_parts(value: Any) -> tuple[str, str] | None:
    """Return normalized start/end text for real one-hour interval labels."""

    if is_summary_time_label(value):
        return None
    start, end = parse_interval(value)
    if start is None or end is None:
        return None
    start_minutes = time_to_minutes(start)
    end_minutes = time_to_minutes(end)
    if start_minutes is None or end_minutes is None:
        return None
    if (end_minutes - start_minutes) % (24 * 60) != 60:
        return None
    return start.strftime("%H:%M"), end.strftime("%H:%M")


def is_hourly_interval_label(value: Any) -> bool:
    return hourly_interval_label_parts(value) is not None


def hourly_interval_rows(dataframe: pd.DataFrame, label_column: str | None = None) -> pd.DataFrame:
    """Return only rows whose first/label column is a real hourly interval."""

    if dataframe.empty or len(dataframe.columns) == 0:
        return dataframe.copy()
    column = label_column or str(dataframe.columns[0])
    if column not in dataframe.columns:
        return dataframe.iloc[0:0].copy()
    mask = dataframe[column].map(is_hourly_interval_label)
    return dataframe.loc[mask].copy()


def hourly_interval_options(dataframe: pd.DataFrame, label_column: str | None = None) -> list[tuple[str, str, str]]:
    """Return selectable one-hour interval labels as (label, start, end)."""

    rows = hourly_interval_rows(dataframe, label_column=label_column)
    if rows.empty or len(rows.columns) == 0:
        return []
    column = label_column or str(rows.columns[0])
    options: list[tuple[str, str, str]] = []
    for label in rows[column]:
        parts = hourly_interval_label_parts(label)
        if parts is None:
            continue
        start, end = parts
        options.append((f"{start}-{end}", start, end))
    return options
